jest.mock('../src/lib/word-diff/index.js', () => ({ applyTokenMapStrategy: jest.fn(), applySentenceDiffStrategy: jest.fn() }));
jest.mock('../src/lib/word-diff/char-diff.js', () => ({ hasCjk: () => false, applyCharDiffStrategy: jest.fn(async () => {}) }));
const { applyCharDiffStrategy } = require('../src/lib/word-diff/char-diff.js');
const { applyTokenMapStrategy } = require('../src/lib/word-diff/index.js');
const { applyChunkResults } = require('../src/lib/reassembler.js');

function world(texts) {
  const ranges = texts.map((text) => ({ text, load: jest.fn(), insertText: jest.fn(), insertComment: jest.fn() }));
  const paragraphs = texts.map((text, index) => ({ text, load: jest.fn(),
    parentTableOrNullObject: { isNullObject: true, load: jest.fn() }, getRange: () => ranges[index] }));
  const range = { text: texts.join('\r'), isNullObject: false, load: jest.fn(),
    paragraphs: { items: paragraphs, load: jest.fn() }, insertText: jest.fn() };
  const context = { sync: jest.fn(async () => {}), document: { changeTrackingMode: 'Off',
    getBookmarkRangeOrNullObject: () => range } };
  global.Word = { run: async (callback) => callback(context), ChangeTrackingMode: { trackAll: 'TrackAll', off: 'Off' }, InsertLocation: { replace: 'Replace' } };
  return { range, ranges };
}

async function apply(stored, current, amendment) {
  const w = world(current);
  const result = await applyChunkResults([{ chunkId: 'body', status: 'fulfilled', amendment, whitespaceOnly: true,
    chunk: { id: 'body', paragraphs: stored.map((text) => ({ text })) } }], new Map([['body', '_body']]),
  { trackChangesEnabled: true, lineDiffEnabled: false, log: jest.fn() });
  return { ...w, result };
}

beforeEach(() => jest.clearAllMocks());
afterEach(() => { delete global.Word; });

test('boundary-space-only changes use precise character edits instead of being trim-compared away', async () => {
  const { result, ranges } = await apply(['  Example body.  '], ['  Example body.  '], 'Example body.');
  expect(result.amendmentsApplied).toBe(1);
  expect(applyCharDiffStrategy).toHaveBeenCalledWith(expect.anything(), ranges[0], '  Example body.  ', 'Example body.', expect.anything(), { trackChanges: false, paragraph: true });
  expect(applyTokenMapStrategy).not.toHaveBeenCalled();
});

test('space-only changes remain paragraph-local and preserve meaningful spacing on other lines', async () => {
  const { result } = await apply(['First  body.', 'x = y'], ['First  body.', 'x = y'], 'First body.\nx = y');
  expect(result.amendmentsApplied).toBe(1);
  expect(applyCharDiffStrategy).toHaveBeenCalledTimes(1);
});

test('manual line breaks remain inside their original native paragraph during cleanup', async () => {
  const { result } = await apply(['First  line\nSecond  line', 'Another paragraph.'],
    ['First  line\nSecond  line', 'Another paragraph.'], 'First line\nSecond line\nAnother paragraph.');
  expect(result.amendmentsApplied).toBe(1);
  expect(applyCharDiffStrategy).toHaveBeenCalledTimes(1);
  expect(applyCharDiffStrategy.mock.calls[0][2]).toBe('First  line\nSecond  line');
  expect(applyCharDiffStrategy.mock.calls[0][3]).toBe('First line\nSecond line');
});

test('same-trimmed-content source drift invalidates whitespace proposals before writes', async () => {
  const { result } = await apply(['  Example body.  '], [' Example body. '], 'Example body.');
  expect(result.amendmentsApplied).toBe(0);
  expect(result.errors.join(' ')).toContain('source changed since staging');
  expect(applyCharDiffStrategy).not.toHaveBeenCalled();
});

test.each(['Rewritten body.', 'Examplebody.', 'Example\nbody.'])('apply rejects invalid cleanup instead of falling back to whole-range replacement: %s', async (amendment) => {
  const { result, range } = await apply(['Example body.'], ['Example body.'], amendment);
  expect(result.amendmentsApplied).toBe(0);
  expect(result.errors).toHaveLength(1);
  expect(applyCharDiffStrategy).not.toHaveBeenCalled();
  expect(range.insertText).not.toHaveBeenCalled();
});
