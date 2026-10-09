/** @jest-environment jsdom */
jest.mock('../src/lib/document-parser.js', () => ({ parseDocument: jest.fn() }));
jest.mock('../src/lib/document-chunker.js', () => ({ chunkDocument: jest.fn() }));
jest.mock('../src/lib/context-extractor.js', () => ({ extractContext: () => ({ definitions: [], outline: [] }) }));
jest.mock('../src/lib/reassembler.js', () => ({ bookmarkChunkRanges: jest.fn(async () => new Map()), cleanupBookmarks: jest.fn(), applyChunkResults: jest.fn() }));
jest.mock('../src/lib/orchestrator.js', () => ({ processChunksParallel: jest.fn() }));
const { runDocumentSkill } = require('../src/taskpane/word-actions.js');
const { parseDocument } = require('../src/lib/document-parser.js');
const { chunkDocument } = require('../src/lib/document-chunker.js');
const { bookmarkChunkRanges } = require('../src/lib/reassembler.js');
const { processChunksParallel } = require('../src/lib/orchestrator.js');

const deps = () => ({ log: jest.fn(), appState: { config: { backend: 'custom', providers: { custom: { model: 'fixture' } } },
  promptManager: { getActivePrompt: () => null } } });
const args = { category: 'amendment', promptTemplate: 'Remove redundant spaces only.', gateApply: true, whitespaceOnly: true };

beforeEach(() => {
  jest.clearAllMocks();
  const chunk = { id: 'body', paragraphs: [{ text: 'Example body.' }], tokenCount: 5 };
  parseDocument.mockResolvedValue({ paragraphs: chunk.paragraphs, totalTokens: 5 });
  chunkDocument.mockReturnValue([chunk]);
  processChunksParallel.mockResolvedValue([{ status: 'fulfilled', chunk, chunkId: 'body', amendment: 'Example body.' }]);
});

test('fresh clean body text is a verified no-op without bookmarks or model requests', async () => {
  const result = await runDocumentSkill(deps(), args);
  expect(result).toMatchObject({ status: 'no_op', satisfied: true });
  expect(bookmarkChunkRanges).not.toHaveBeenCalled();
  expect(processChunksParallel).not.toHaveBeenCalled();
});

test('table-space candidates belong to the separate native table task', async () => {
  parseDocument.mockResolvedValue({ paragraphs: [{ text: 'Body text.' }, { text: 'Table  cell', inTable: true }], totalTokens: 5 });
  expect(await runDocumentSkill(deps(), args)).toMatchObject({ status: 'no_op', satisfied: true });
  expect(processChunksParallel).not.toHaveBeenCalled();
});

test('source candidates still run the model and rejected content changes remain retryable failures', async () => {
  const chunk = { id: 'body', paragraphs: [{ text: 'Example  body.' }], tokenCount: 5 };
  parseDocument.mockResolvedValue({ paragraphs: chunk.paragraphs, totalTokens: 5 });
  chunkDocument.mockReturnValue([chunk]);
  processChunksParallel.mockResolvedValue([{ status: 'fulfilled', chunk, chunkId: 'body', amendment: 'A new meaning.' }]);
  const result = await runDocumentSkill(deps(), args);
  expect(result.failedCount).toBe(1);
  expect(result.results[0]).toMatchObject({ status: 'rejected', whitespaceOnly: true });
  expect(result.retryFailed).toEqual(expect.any(Function));
});

test('generic amendment work always runs its model even when text has no space candidates', async () => {
  await runDocumentSkill(deps(), { ...args, whitespaceOnly: false });
  expect(processChunksParallel).toHaveBeenCalledTimes(1);
});

test('cancellation during the host parse prevents no-op claims and inference', async () => {
  const controller = new AbortController();
  parseDocument.mockImplementation(async () => { controller.abort(); return { paragraphs: [], totalTokens: 0 }; });
  await expect(runDocumentSkill(deps(), { ...args, signal: controller.signal })).rejects.toMatchObject({ name: 'AbortError' });
  expect(processChunksParallel).not.toHaveBeenCalled();
});
