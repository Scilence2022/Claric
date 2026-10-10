/** @jest-environment jsdom */

jest.mock('../src/lib/document-parser.js', () => ({ parseDocument: jest.fn() }));
jest.mock('../src/lib/document-chunker.js', () => ({ chunkDocument: jest.fn() }));
jest.mock('../src/lib/context-extractor.js', () => ({ extractContext: () => ({ definitions: [], outline: [] }) }));
jest.mock('../src/lib/reassembler.js', () => ({
  bookmarkChunkRanges: jest.fn(), applyChunkResults: jest.fn(), cleanupBookmarks: jest.fn(),
}));
jest.mock('../src/lib/orchestrator.js', () => ({ processChunksParallel: jest.fn() }));

const { runDocumentSkill } = require('../src/taskpane/word-actions.js');
const { parseDocument } = require('../src/lib/document-parser.js');
const { chunkDocument } = require('../src/lib/document-chunker.js');
const { bookmarkChunkRanges, applyChunkResults, cleanupBookmarks } = require('../src/lib/reassembler.js');
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
