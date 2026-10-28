/** @jest-environment jsdom */

jest.mock('../src/lib/document-parser.js', () => ({ parseDocument: jest.fn() }));
jest.mock('../src/lib/document-chunker.js', () => ({ chunkDocument: jest.fn() }));
jest.mock('../src/lib/context-extractor.js', () => ({ extractContext: () => ({ definitions: [], outline: [] }) }));
jest.mock('../src/lib/reassembler.js', () => ({
  bookmarkChunkRanges: jest.fn(), applyChunkResults: jest.fn(), cleanupBookmarks: jest.fn(), readChunkSnapshot: jest.fn(),
}));
jest.mock('../src/lib/orchestrator.js', () => ({ processChunksParallel: jest.fn() }));

const { runDocumentSkill } = require('../src/taskpane/word-actions.js');
const { parseDocument } = require('../src/lib/document-parser.js');
const { chunkDocument } = require('../src/lib/document-chunker.js');
const { bookmarkChunkRanges, applyChunkResults, cleanupBookmarks, readChunkSnapshot } = require('../src/lib/reassembler.js');
const { processChunksParallel } = require('../src/lib/orchestrator.js');

const chunks = ['applied', 'blocked', 'model-failed'].map((id, index) => ({
  id, startIndex: index, endIndex: index, tokenCount: 5,
  paragraphs: [{ text: `Paragraph ${index + 1}.`, revisionFingerprint: null }],
}));
const bookmarks = new Map(chunks.map((chunk) => [chunk.id, `_${chunk.id}`]));
const fulfilled = (chunk) => ({ status: 'fulfilled', chunk, chunkId: chunk.id, amendment: `Revised ${chunk.id} paragraph.` });
const deps = () => ({ log: jest.fn(), logWithRetry: jest.fn(), appState: {
  config: { backend: 'custom', providers: { custom: { model: 'fixture' } }, trackChangesEnabled: true },
  promptManager: { getActivePrompt: () => null },
} });

beforeEach(() => {
  jest.clearAllMocks();
  parseDocument.mockResolvedValue({ paragraphs: chunks.flatMap((chunk) => chunk.paragraphs), totalTokens: 15 });
  chunkDocument.mockReturnValue(chunks);
  bookmarkChunkRanges.mockResolvedValue(bookmarks);
  processChunksParallel.mockResolvedValue([
    fulfilled(chunks[0]), fulfilled(chunks[1]),
    { status: 'rejected', chunk: chunks[2], chunkId: 'model-failed', error: 'Upstream timeout.' },
  ]);
  cleanupBookmarks.mockResolvedValue(undefined);
});

test('application failures retain their bookmarks together with model failures and report a warning', async () => {
  const applicationResult = { amendmentsApplied: 1, commentsInserted: 0, errors: ['Native baseline changed.'],
    appliedChunkIds: ['applied'], attemptedChunkIds: ['applied', 'blocked'], failedChunkIds: ['blocked'],
    partialChunkIds: [], uncertainChunkIds: [], interrupted: false };
  applyChunkResults.mockResolvedValue(applicationResult);
  const dependencies = deps();
  const staged = await runDocumentSkill(dependencies, { category: 'amendment', promptTemplate: 'Polish the wording.', gateApply: true });
  expect(applyChunkResults).not.toHaveBeenCalled();
  expect(await staged.apply()).toBe(applicationResult);

  expect(cleanupBookmarks).toHaveBeenCalledWith(bookmarks, { keep: new Set(['_blocked', '_model-failed']) });
  expect(dependencies.log).toHaveBeenCalledWith(expect.stringContaining('1 section(s) failed during application'), 'warning');
  expect(dependencies.logWithRetry).toHaveBeenCalledWith(expect.stringContaining('1 chunk(s) failed'), 'warning', expect.any(Function));
});

test('partial or unverified application cannot be summarized as a wholly untouched document', async () => {
  processChunksParallel.mockResolvedValue([fulfilled(chunks[0]), fulfilled(chunks[1])]);
  applyChunkResults.mockResolvedValue({ amendmentsApplied: 0, commentsInserted: 0, errors: ['Read-back failed.'],
    appliedChunkIds: [], attemptedChunkIds: ['blocked'], failedChunkIds: ['blocked'],
    partialChunkIds: ['blocked'], uncertainChunkIds: ['blocked'], interrupted: false });
  const dependencies = deps();
  const staged = await runDocumentSkill(dependencies, { category: 'amendment', promptTemplate: 'Polish the wording.', gateApply: true });
  await staged.apply(['blocked']);

  expect(cleanupBookmarks).toHaveBeenCalledWith(bookmarks, { keep: new Set(['_blocked']) });
  expect(dependencies.log).toHaveBeenCalledWith(expect.stringContaining('some edits already reached Word; inspect the document and draft fresh edits'), 'warning');
  expect(dependencies.logWithRetry).not.toHaveBeenCalled();
});

test('interrupted application preserves bookmarks for untouched remaining sections', async () => {
  applyChunkResults.mockResolvedValue({ amendmentsApplied: 1, commentsInserted: 0, errors: [],
    appliedChunkIds: ['applied'], attemptedChunkIds: ['applied'], failedChunkIds: [],
    partialChunkIds: [], uncertainChunkIds: [], interrupted: true });
  const staged = await runDocumentSkill(deps(), { category: 'amendment', promptTemplate: 'Polish the wording.', gateApply: true });
  await staged.apply();
  expect(cleanupBookmarks).not.toHaveBeenCalled();
});

describe('targeted edit scope', () => {
  const doc = [
    { index: 0, text: 'Abstract', headingLevel: 1 },
    { index: 1, text: 'We study X.', headingLevel: 0 },
    { index: 2, text: 'Discussion', headingLevel: 1 },
    { index: 3, text: 'A causal link.', headingLevel: 0 },
  ];
  const twoChunks = [
    { id: 'abstract', tokenCount: 5, paragraphs: doc.slice(0, 2) },
    { id: 'discussion', tokenCount: 5, paragraphs: doc.slice(2) },
  ];
  beforeEach(() => {
    parseDocument.mockResolvedValue({ paragraphs: doc, totalTokens: 10 });
    chunkDocument.mockReturnValue(twoChunks);
    processChunksParallel.mockResolvedValue([
      { status: 'fulfilled', chunk: twoChunks[0], chunkId: 'abstract', amendment: 'Abstract\nWe examine X.' },
      { status: 'fulfilled', chunk: twoChunks[1], chunkId: 'discussion', amendment: 'Discussion\nA likely causal link.' },
    ]);
  });

  test('a request naming one part changes only that part', async () => {
    const staged = await runDocumentSkill(deps(), { category: 'amendment',
      promptTemplate: 'Rewrite only the Abstract paragraph. Do not change any other paragraph.', gateApply: true });
    expect(staged.results.map((r) => r.amendment)).toEqual(['Abstract\nWe examine X.', null]);
    expect(staged.scope).toEqual({ label: 'Abstract', revertedParagraphs: 1, droppedChunks: 1,
      notice: 'Limited to Abstract. 1 model edit(s) outside it were discarded.' });
  });

  test('an unlocatable target keeps every section and is reported', async () => {
    const staged = await runDocumentSkill(deps(), { category: 'amendment',
      promptTemplate: 'Rewrite only the Conclusion paragraph.', gateApply: true });
    expect(staged.results.map((r) => r.amendment)).toEqual(['Abstract\nWe examine X.', 'Discussion\nA likely causal link.']);
    expect(staged.scope).toMatchObject({ label: 'Conclusion', unresolved: true });
    expect(staged.scope.notice).toContain('could not be located unambiguously');
  });

  test('an unrestricted request is not confined', async () => {
    const staged = await runDocumentSkill(deps(), { category: 'amendment', promptTemplate: 'Polish the wording.', gateApply: true });
    expect(staged.scope).toBeNull();
    expect(staged.results.every((r) => r.amendment)).toBe(true);
  });
});

test('document review policy and full source context survive repeated failed-chunk retries', async () => {
  const failed = { status: 'rejected', chunk: chunks[0], chunkId: chunks[0].id, error: 'Verification unavailable.' };
  processChunksParallel.mockResolvedValue([failed]);
  const staged = await runDocumentSkill(deps(), { category: 'comment', promptTemplate: 'Review.',
    reviewPolicy: 'document-consistency', gateApply: true });
  const firstRetry = await staged.retryFailed();
  await firstRetry.retryFailed();
  expect(processChunksParallel).toHaveBeenCalledTimes(3);
  const originalContext = processChunksParallel.mock.calls[0][1].documentContext;
  for (const [, options] of processChunksParallel.mock.calls) {
    expect(options.reviewPolicy).toBe('document-consistency');
    expect(options.documentContext).toBe(originalContext);
  }
  expect(applyChunkResults).not.toHaveBeenCalled();
});

describe('failed section recovery', () => {
  const failedApply = { amendmentsApplied: 0, failedChunkIds: ['blocked'], uncertainChunkIds: ['blocked'],
    attemptedChunkIds: ['blocked'], errors: ['Read-back differs.'] };
  const succeeded = { amendmentsApplied: 1, failedChunkIds: [], errors: [], appliedChunkIds: ['blocked'] };
  let staged;
  let currentChunk;
  beforeEach(async () => {
    applyChunkResults.mockResolvedValue(failedApply);
    staged = await runDocumentSkill(deps(), { category: 'amendment', promptTemplate: 'Polish.', gateApply: true });
    await staged.apply(['blocked']);
    currentChunk = { ...chunks[1], paragraphs: [{ text: 'Current partially edited text.', revisionFingerprint: 'fresh' }] };
    readChunkSnapshot.mockResolvedValue(currentChunk);
    applyChunkResults.mockClear().mockResolvedValue(succeeded);
    cleanupBookmarks.mockClear();
    processChunksParallel.mockClear();
  });

  test('reapply retains the proposal baseline and touches only the failed section', async () => {
    await staged.recoverChunk('blocked', 'retry');
    expect(readChunkSnapshot).not.toHaveBeenCalled();
    const [selected, selectedBookmarks, options] = applyChunkResults.mock.calls[0];
    expect(selected.map((r) => r.chunkId)).toEqual(['blocked']);
    expect(selectedBookmarks).toEqual(new Map([['blocked', '_blocked']]));
    expect(options.chunkOriginals.get('blocked')).toEqual(['Paragraph 2.']);
    expect(cleanupBookmarks).toHaveBeenCalledWith(selectedBookmarks);
    expect(processChunksParallel).not.toHaveBeenCalled();
    await expect(staged.recoverChunk('blocked', 'retry')).rejects.toThrow('Only failed sections');
    await expect(staged.recoverChunk('applied', 'force')).rejects.toThrow('Only failed sections');
  });

  test('force rebases the existing target on current text and fingerprints without bypassing verification', async () => {
    await staged.recoverChunk('blocked', 'force');
    expect(readChunkSnapshot).toHaveBeenCalledWith(chunks[1], '_blocked');
    const [selected, , options] = applyChunkResults.mock.calls[0];
    expect(selected[0]).toMatchObject({ amendment: 'Revised blocked paragraph.', chunk: currentChunk, comment: null });
    expect(options.chunkOriginals.get('blocked')).toEqual(['Current partially edited text.']);
    expect(options.chunkRevisionFingerprints.get('blocked')).toEqual(['fresh']);
    expect(options.trackChangesEnabled).toBe(true);
    expect(processChunksParallel).not.toHaveBeenCalled();
  });

  test('regeneration uses fresh source and applies only the newly generated section', async () => {
    processChunksParallel.mockResolvedValue([{ ...fulfilled(currentChunk), amendment: 'Newly generated wording.' }]);
    const onProposal = jest.fn();
    await staged.recoverChunk('blocked', 'regenerate', { onProposal });
    expect(processChunksParallel).toHaveBeenCalledWith([currentChunk], expect.objectContaining({ concurrency: 1 }));
    expect(applyChunkResults.mock.calls[0][0][0].amendment).toBe('Newly generated wording.');
    expect(onProposal).toHaveBeenCalledWith({ before: 'Current partially edited text.', after: 'Newly generated wording.' });
  });

  test('a recovery failure keeps the bookmark and allows another deliberate attempt', async () => {
    applyChunkResults.mockResolvedValueOnce(failedApply);
    await staged.recoverChunk('blocked', 'force');
    expect(cleanupBookmarks).not.toHaveBeenCalled();
    await staged.recoverChunk('blocked', 'retry');
    expect(applyChunkResults.mock.calls[1][2].chunkOriginals.get('blocked')).toEqual(['Current partially edited text.']);
    expect(cleanupBookmarks).toHaveBeenCalledTimes(1);
  });

  test('generation failure never applies the old target; cancellation never starts recovery writes', async () => {
    processChunksParallel.mockResolvedValue([{ status: 'rejected', error: 'Provider unavailable.' }]);
    await expect(staged.recoverChunk('blocked', 'regenerate')).rejects.toThrow('Provider unavailable');
    expect(applyChunkResults).not.toHaveBeenCalled();
    const controller = new AbortController();
    controller.abort();
    await expect(staged.recoverChunk('blocked', 'force', { signal: controller.signal })).rejects.toMatchObject({ name: 'AbortError' });
    expect(applyChunkResults).not.toHaveBeenCalled();
    expect(cleanupBookmarks).not.toHaveBeenCalled();
  });

  test('edits during regeneration are guarded by the captured source, and concurrent recovery is refused', async () => {
    let finish;
    processChunksParallel.mockImplementationOnce(() => new Promise((resolve) => { finish = resolve; }));
    const pending = staged.recoverChunk('blocked', 'regenerate');
    while (!finish) await new Promise((resolve) => setTimeout(resolve, 0));
    await expect(staged.recoverChunk('blocked', 'force')).rejects.toThrow('already running');
    finish([{ ...fulfilled(currentChunk), amendment: 'Fresh rewrite.' }]);
    await pending;
    expect(applyChunkResults.mock.calls[0][2].chunkRevisionFingerprints.get('blocked')).toEqual(['fresh']);
  });
});

test('resuming the remaining selection retains bookmarks for earlier failed attempts', async () => {
  applyChunkResults.mockResolvedValueOnce({ amendmentsApplied: 0, failedChunkIds: ['blocked'],
    uncertainChunkIds: ['blocked'], interrupted: true, errors: ['Read-back failed.'] });
  const staged = await runDocumentSkill(deps(), { category: 'amendment', promptTemplate: 'Polish.', gateApply: true });
  await staged.apply(['blocked']);
  applyChunkResults.mockResolvedValueOnce({ amendmentsApplied: 1, failedChunkIds: [], interrupted: false, errors: [] });
  await staged.apply(['applied']);
  expect(cleanupBookmarks).toHaveBeenCalledWith(bookmarks, { keep: new Set(['_blocked', '_model-failed']) });
});
