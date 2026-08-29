/**
 * Unit tests for src/lib/reassembler.js
 * Tests bookmarkChunkRanges, applyChunkResults, cleanupBookmarks exports.
 *
 * Covers:
 * - REASSEMBLY-01: Reverse-order amendment application
 * - Bookmark lifecycle: create -> use -> cleanup
 * - Comment insertion after amendments
 * - Partial failure handling
 */

// --- Mock Word API ---

/**
 * Creates a mock Word.run that captures its callback and provides a mock context.
 * Each call to Word.run creates a fresh context with the specified paragraphs.
 */
function createMockWordRun(paragraphItems, bookmarkRanges = {}) {
  const syncFn = jest.fn().mockResolvedValue(undefined);
  const insertedBookmarks = {};
  const deletedBookmarks = [];
  const insertedComments = [];
  const changeTrackingModes = [];

  function makeRange(text, id) {
    const range = {
      text,
      isNullObject: false,
      load: jest.fn().mockReturnValue(undefined),
      insertBookmark: jest.fn().mockImplementation((name) => {
        insertedBookmarks[name] = { text, rangeId: id };
      }),
      insertComment: jest.fn().mockImplementation((commentText) => {
        insertedComments.push({ commentText, rangeText: text, rangeId: id });
      }),
      expandTo: jest.fn().mockImplementation(function () { return this; }),
    };
    return range;
  }

  // Build paragraph mock items
  const items = paragraphItems.map((p, i) => {
    return {
      text: p.text,
      load: jest.fn(), // real Office.js paragraph proxies all expose load
      // Paragraphs in these mocks live outside tables unless flagged.
      parentTableOrNullObject: { isNullObject: !p.inTable, load: jest.fn() },
      getRange: jest.fn().mockImplementation((position) => {
        // Return a range-like object that can be used with expandTo
        const r = makeRange(p.text, `para-${i}-${position}`);
        r.expandTo = jest.fn().mockImplementation((otherRange) => {
          // Create a combined range with combined text
          return makeRange(`expanded-${i}`, `expanded-para-${i}`);
        });
        return r;
      }),
    };
  });

  const mockContext = {
    document: {
      body: {
        paragraphs: {
          items,
          load: jest.fn().mockReturnValue(undefined),
        },
      },
      getBookmarkRangeOrNullObject: jest.fn().mockImplementation((name) => {
        if (bookmarkRanges[name] || insertedBookmarks[name]) {
          const bm = bookmarkRanges[name] || insertedBookmarks[name];
          return makeRange(bm.text, `bookmark-${name}`);
        }
        return { isNullObject: true, load: jest.fn(), text: '' };
      }),
      deleteBookmark: jest.fn().mockImplementation((name) => {
        deletedBookmarks.push(name);
      }),
      changeTrackingMode: null,
    },
    sync: syncFn,
  };

  // Track changeTrackingMode assignments via setter
  let _trackingMode = null;
  Object.defineProperty(mockContext.document, 'changeTrackingMode', {
    get: () => _trackingMode,
    set: (val) => {
      _trackingMode = val;
      changeTrackingModes.push(val);
    },
  });

  const wordRun = jest.fn().mockImplementation(async (callback) => {
    await callback(mockContext);
  });

  return {
    wordRun,
    mockContext,
    syncFn,
    insertedBookmarks,
    deletedBookmarks,
    insertedComments,
    changeTrackingModes,
    items,
  };
}

// Mock Word global with ChangeTrackingMode
global.Word = {
  run: jest.fn(),
  ChangeTrackingMode: {
    trackAll: 'TrackAll',
    off: 'Off',
  },
  InsertLocation: {
    after: 'After',
    before: 'Before',
    replace: 'Replace',
  },
};

// Mock the vendored word-diff strategy layer
jest.mock('../src/lib/word-diff/index.js', () => ({
  applyTokenMapStrategy: jest.fn().mockResolvedValue(undefined),
  applySentenceDiffStrategy: jest.fn().mockResolvedValue(undefined),
}));

const { applyTokenMapStrategy, applySentenceDiffStrategy } = require('../src/lib/word-diff/index.js');
const { bookmarkChunkRanges, applyChunkResults, cleanupBookmarks, _normalizeLineEndings, _alignParagraphs, _findAnchorWindow } = require('../src/lib/reassembler.js');

// --- Mock Helpers ---

function mockChunk(id, index, text, startIndex, endIndex) {
  return {
    id,
    // One ParsedParagraph per line, mirroring the real chunker contract
    // (bookmarkChunkRanges validates boundary paragraphs against these).
    paragraphs: text.split('\n').map((t) => ({ text: t, headingLevel: 0 })),
    startIndex,
    endIndex,
    tokenCount: Math.ceil(text.length / 4),
    sectionTitle: '',
    overlapBefore: '',
  };
}

function makeChunkResult(chunkId, chunkIndex, status, opts = {}) {
  return {
    chunkId,
    chunkIndex,
    status,
    amendment: opts.amendment || null,
    comment: opts.comment || null,
    error: opts.error || null,
    chunk: opts.chunk || mockChunk(chunkId, chunkIndex, `text-${chunkIndex}`, chunkIndex * 3, chunkIndex * 3 + 2),
  };
}

// --- Test Suites ---

describe('bookmarkChunkRanges', () => {
  beforeEach(() => {
    jest.clearAllMocks();
  });

  test('creates bookmarks for each chunk using _wdp prefix naming convention', async () => {
    const paragraphs = [
      { text: 'Para 0' }, { text: 'Para 1' }, { text: 'Para 2' },
      { text: 'Para 3' }, { text: 'Para 4' }, { text: 'Para 5' },
    ];
    const mock = createMockWordRun(paragraphs);
    global.Word.run = mock.wordRun;

    const chunks = [
      mockChunk('chunk-0', 0, 'Para 0\nPara 1', 0, 1),
      mockChunk('chunk-1', 1, 'Para 2\nPara 3', 2, 3),
      mockChunk('chunk-2', 2, 'Para 4\nPara 5', 4, 5),
    ];

    const bookmarkMap = await bookmarkChunkRanges(chunks);

    expect(bookmarkMap).toBeInstanceOf(Map);
    expect(bookmarkMap.size).toBe(3);

    // All bookmark names should start with _wdp
    for (const [, bookmarkName] of bookmarkMap) {
      expect(bookmarkName).toMatch(/^_wdp/);
    }

    // All chunk IDs should be in the map
    expect(bookmarkMap.has('chunk-0')).toBe(true);
    expect(bookmarkMap.has('chunk-1')).toBe(true);
    expect(bookmarkMap.has('chunk-2')).toBe(true);
  });

  test('returns a Map from chunkId to bookmarkName', async () => {
    const paragraphs = [
      { text: 'Para 0' }, { text: 'Para 1' },
    ];
    const mock = createMockWordRun(paragraphs);
    global.Word.run = mock.wordRun;

    const chunks = [
      mockChunk('chunk-0', 0, 'Para 0\nPara 1', 0, 1),
    ];

    const bookmarkMap = await bookmarkChunkRanges(chunks);

    expect(bookmarkMap).toBeInstanceOf(Map);
    expect(bookmarkMap.size).toBe(1);
    const [key, value] = [...bookmarkMap.entries()][0];
    expect(key).toBe('chunk-0');
    expect(typeof value).toBe('string');
  });

  test('bookmark names are unique across chunks', async () => {
    const paragraphs = [
      { text: 'P0' }, { text: 'P1' }, { text: 'P2' },
      { text: 'P3' }, { text: 'P4' }, { text: 'P5' },
    ];
    const mock = createMockWordRun(paragraphs);
    global.Word.run = mock.wordRun;

    const chunks = [
      mockChunk('chunk-0', 0, 'P0\nP1', 0, 1),
      mockChunk('chunk-1', 1, 'P2\nP3', 2, 3),
      mockChunk('chunk-2', 2, 'P4\nP5', 4, 5),
    ];

    const bookmarkMap = await bookmarkChunkRanges(chunks);

    const names = [...bookmarkMap.values()];
    const uniqueNames = new Set(names);
    expect(uniqueNames.size).toBe(names.length);
  });
});

describe('applyChunkResults', () => {
  beforeEach(() => {
    jest.clearAllMocks();
  });

  test('processes chunks in reverse order (highest startIndex first)', async () => {
    const applicationOrder = [];
    applyTokenMapStrategy.mockImplementation(async (context, range, original, amended, log) => {
      applicationOrder.push(original);
    });

    const paragraphs = [
      { text: 'Para 0' }, { text: 'Para 1' }, { text: 'Para 2' },
      { text: 'Para 3' }, { text: 'Para 4' }, { text: 'Para 5' },
      { text: 'Para 6' }, { text: 'Para 7' }, { text: 'Para 8' },
    ];

    const bookmarkRanges = {
      '_wdpbm0': { text: 'Para 0\nPara 1\nPara 2' },
      '_wdpbm1': { text: 'Para 3\nPara 4\nPara 5' },
      '_wdpbm2': { text: 'Para 6\nPara 7\nPara 8' },
    };
    const mock = createMockWordRun(paragraphs, bookmarkRanges);
    global.Word.run = mock.wordRun;

    const chunk0 = mockChunk('chunk-0', 0, 'Para 0\nPara 1\nPara 2', 0, 2);
    const chunk1 = mockChunk('chunk-1', 1, 'Para 3\nPara 4\nPara 5', 3, 5);
    const chunk2 = mockChunk('chunk-2', 2, 'Para 6\nPara 7\nPara 8', 6, 8);

    const results = [
      makeChunkResult('chunk-0', 0, 'fulfilled', { amendment: 'Amended 0', chunk: chunk0 }),
      makeChunkResult('chunk-1', 1, 'fulfilled', { amendment: 'Amended 1', chunk: chunk1 }),
      makeChunkResult('chunk-2', 2, 'fulfilled', { amendment: 'Amended 2', chunk: chunk2 }),
    ];

    const bookmarkMap = new Map([
      ['chunk-0', '_wdpbm0'],
      ['chunk-1', '_wdpbm1'],
      ['chunk-2', '_wdpbm2'],
    ]);

    await applyChunkResults(results, bookmarkMap, {
      trackChangesEnabled: true,
      lineDiffEnabled: false,
      log: jest.fn(),
    });

    // Verify reverse order: chunk-2 (startIndex=6) first, then chunk-1 (3), then chunk-0 (0)
    expect(applicationOrder).toHaveLength(3);
    // The original text passed should reflect reverse order
    expect(applicationOrder[0]).toContain('Para 6');
    expect(applicationOrder[1]).toContain('Para 3');
    expect(applicationOrder[2]).toContain('Para 0');
  });

  test('calls applyTokenMapStrategy with correct original text and amended text', async () => {
    const paragraphs = [
      { text: 'Original clause text here' }, { text: 'More text' },
    ];

    const bookmarkRanges = {
      '_wdpbm0': { text: 'Original clause text here\nMore text' },
    };
    const mock = createMockWordRun(paragraphs, bookmarkRanges);
    global.Word.run = mock.wordRun;

    const chunk = mockChunk('chunk-0', 0, 'Original clause text here\nMore text', 0, 1);

    const results = [
      makeChunkResult('chunk-0', 0, 'fulfilled', { amendment: 'Revised clause text', chunk }),
    ];

    const bookmarkMap = new Map([['chunk-0', '_wdpbm0']]);

    await applyChunkResults(results, bookmarkMap, {
      trackChangesEnabled: true,
      lineDiffEnabled: false,
      log: jest.fn(),
    });

    expect(applyTokenMapStrategy).toHaveBeenCalled();
    const args = applyTokenMapStrategy.mock.calls[0];
    // args: (context, range, originalText, amendedText, log)
    expect(args[3]).toBe('Revised clause text');
  });

  test('uses applySentenceDiffStrategy when lineDiffEnabled=true', async () => {
    const paragraphs = [{ text: 'Some text' }];

    const bookmarkRanges = {
      '_wdpbm0': { text: 'Some text' },
    };
    const mock = createMockWordRun(paragraphs, bookmarkRanges);
    global.Word.run = mock.wordRun;

    const chunk = mockChunk('chunk-0', 0, 'Some text', 0, 0);

    const results = [
      makeChunkResult('chunk-0', 0, 'fulfilled', { amendment: 'Amended text', chunk }),
    ];

    const bookmarkMap = new Map([['chunk-0', '_wdpbm0']]);

    await applyChunkResults(results, bookmarkMap, {
      trackChangesEnabled: true,
      lineDiffEnabled: true,
      log: jest.fn(),
    });

    expect(applySentenceDiffStrategy).toHaveBeenCalled();
    expect(applyTokenMapStrategy).not.toHaveBeenCalled();
  });

  test('paragraph-level keep branch honors the toggle (per-paragraph sentence mode)', async () => {
    // A bookmark range that supports the paragraph-level path (.paragraphs),
    // so the keep branch — not the range-level fallback — runs.
    const paraItem = {
      text: 'Original clause text here',
      parentTableOrNullObject: { isNullObject: true, load: jest.fn() },
      load: jest.fn(),
      getRange: jest.fn().mockImplementation((position) => ({
        text: 'Original clause text here',
        isNullObject: false,
        load: jest.fn(),
        insertText: jest.fn(),
      })),
    };
    const bookmarkRange = {
      text: 'Original clause text here',
      isNullObject: false,
      load: jest.fn(),
      paragraphs: { items: [paraItem], load: jest.fn() },
    };
    const mock = createMockWordRun([], {});
    mock.mockContext.document.getBookmarkRangeOrNullObject
      .mockImplementation(() => bookmarkRange);
    global.Word.run = mock.wordRun;

    const chunk = mockChunk('chunk-0', 0, 'Original clause text here', 0, 0);
    const results = [
      makeChunkResult('chunk-0', 0, 'fulfilled', { amendment: 'Revised clause text', chunk }),
    ];

    await applyChunkResults(results, new Map([['chunk-0', '_wdpbm0']]), {
      trackChangesEnabled: true,
      lineDiffEnabled: true,
      log: jest.fn(),
    });

    expect(applySentenceDiffStrategy).toHaveBeenCalledTimes(1);
    const args = applySentenceDiffStrategy.mock.calls[0];
    expect(args[2]).toBe('Original clause text here');
    expect(args[3]).toBe('Revised clause text');
    // trackChanges:false — the paragraph loop owns the tracking mode.
    expect(args[5]).toEqual({ trackChanges: false });
    expect(applyTokenMapStrategy).not.toHaveBeenCalled();

    // Toggle OFF takes the token-map branch for the same keep edit.
    applySentenceDiffStrategy.mockClear();
    applyTokenMapStrategy.mockClear();
    await applyChunkResults(results, new Map([['chunk-0', '_wdpbm0']]), {
      trackChangesEnabled: true,
      lineDiffEnabled: false,
      log: jest.fn(),
    });
    expect(applyTokenMapStrategy).toHaveBeenCalledTimes(1);
    expect(applySentenceDiffStrategy).not.toHaveBeenCalled();
  });

  test('skips chunks with status="rejected"', async () => {
    const paragraphs = [
      { text: 'Para 0' }, { text: 'Para 1' },
      { text: 'Para 2' }, { text: 'Para 3' },
    ];

    const bookmarkRanges = {
      '_wdpbm0': { text: 'Para 0' },
      '_wdpbm1': { text: 'Para 2' },
    };
    const mock = createMockWordRun(paragraphs, bookmarkRanges);
    global.Word.run = mock.wordRun;

    const chunk0 = mockChunk('chunk-0', 0, 'Para 0', 0, 1);
    const chunk1 = mockChunk('chunk-1', 1, 'Para 2', 2, 3);

    const results = [
      makeChunkResult('chunk-0', 0, 'fulfilled', { amendment: 'Amended 0', chunk: chunk0 }),
      makeChunkResult('chunk-1', 1, 'rejected', { error: 'LLM timeout', chunk: chunk1 }),
    ];

    const bookmarkMap = new Map([
      ['chunk-0', '_wdpbm0'],
      ['chunk-1', '_wdpbm1'],
    ]);

    const result = await applyChunkResults(results, bookmarkMap, {
      trackChangesEnabled: true,
      lineDiffEnabled: false,
      log: jest.fn(),
    });

    // Only 1 amendment applied (the fulfilled one)
    expect(result.amendmentsApplied).toBe(1);
    expect(applyTokenMapStrategy).toHaveBeenCalledTimes(1);
  });

  test('skips chunks with status="cancelled"', async () => {
    const paragraphs = [{ text: 'Para 0' }];

    const bookmarkRanges = {
      '_wdpbm0': { text: 'Para 0' },
    };
    const mock = createMockWordRun(paragraphs, bookmarkRanges);
    global.Word.run = mock.wordRun;

    const chunk0 = mockChunk('chunk-0', 0, 'Para 0', 0, 0);

    const results = [
      makeChunkResult('chunk-0', 0, 'cancelled', { chunk: chunk0 }),
    ];

    const bookmarkMap = new Map([['chunk-0', '_wdpbm0']]);

    const result = await applyChunkResults(results, bookmarkMap, {
      trackChangesEnabled: true,
      lineDiffEnabled: false,
      log: jest.fn(),
    });

    expect(result.amendmentsApplied).toBe(0);
    expect(applyTokenMapStrategy).not.toHaveBeenCalled();
  });

  test('inserts comments on bookmarked ranges after all amendments', async () => {
    const amendmentCallOrder = [];

    applyTokenMapStrategy.mockImplementation(async () => {
      amendmentCallOrder.push(Date.now());
    });

    const paragraphs = [
      { text: 'Para 0' }, { text: 'Para 1' },
      { text: 'Para 2' }, { text: 'Para 3' },
    ];

    const bookmarkRanges = {
      '_wdpbm0': { text: 'Para 0\nPara 1' },
      '_wdpbm1': { text: 'Para 2\nPara 3' },
    };
    const mock = createMockWordRun(paragraphs, bookmarkRanges);
    // Track insertComment calls
    mock.wordRun.mockImplementation(async (callback) => {
      await callback(mock.mockContext);
    });
    global.Word.run = mock.wordRun;

    const chunk0 = mockChunk('chunk-0', 0, 'Para 0\nPara 1', 0, 1);
    const chunk1 = mockChunk('chunk-1', 1, 'Para 2\nPara 3', 2, 3);

    const results = [
      makeChunkResult('chunk-0', 0, 'fulfilled', { amendment: 'Amended 0', comment: 'Comment on chunk 0', chunk: chunk0 }),
      makeChunkResult('chunk-1', 1, 'fulfilled', { amendment: 'Amended 1', comment: 'Comment on chunk 1', chunk: chunk1 }),
    ];

    const bookmarkMap = new Map([
      ['chunk-0', '_wdpbm0'],
      ['chunk-1', '_wdpbm1'],
    ]);

    await applyChunkResults(results, bookmarkMap, {
      trackChangesEnabled: true,
      lineDiffEnabled: false,
      log: jest.fn(),
    });

    // Both amendments should be applied
    expect(applyTokenMapStrategy).toHaveBeenCalledTimes(2);

    // Comments should also be inserted (via insertComment on ranges)
    expect(mock.insertedComments.length).toBe(2);
  });

  test('returns counts (amendmentsApplied, commentsInserted, errors)', async () => {
    const paragraphs = [
      { text: 'Para 0' }, { text: 'Para 1' },
      { text: 'Para 2' }, { text: 'Para 3' },
      { text: 'Para 4' }, { text: 'Para 5' },
    ];

    const bookmarkRanges = {
      '_wdpbm0': { text: 'Para 0\nPara 1' },
      '_wdpbm1': { text: 'Para 2\nPara 3' },
      '_wdpbm2': { text: 'Para 4\nPara 5' },
    };
    const mock = createMockWordRun(paragraphs, bookmarkRanges);
    global.Word.run = mock.wordRun;

    const chunk0 = mockChunk('chunk-0', 0, 'Para 0\nPara 1', 0, 1);
    const chunk1 = mockChunk('chunk-1', 1, 'Para 2\nPara 3', 2, 3);
    const chunk2 = mockChunk('chunk-2', 2, 'Para 4\nPara 5', 4, 5);

    const results = [
      makeChunkResult('chunk-0', 0, 'fulfilled', { amendment: 'Amended 0', comment: 'Comment 0', chunk: chunk0 }),
      makeChunkResult('chunk-1', 1, 'rejected', { error: 'LLM error', chunk: chunk1 }),
      makeChunkResult('chunk-2', 2, 'fulfilled', { amendment: 'Amended 2', chunk: chunk2 }),
    ];

    const bookmarkMap = new Map([
      ['chunk-0', '_wdpbm0'],
      ['chunk-1', '_wdpbm1'],
      ['chunk-2', '_wdpbm2'],
    ]);

    const result = await applyChunkResults(results, bookmarkMap, {
      trackChangesEnabled: true,
      lineDiffEnabled: false,
      log: jest.fn(),
    });

    expect(result.amendmentsApplied).toBe(2);
    expect(result.commentsInserted).toBe(1);
    expect(result.errors).toHaveLength(1);
    expect(result.errors[0]).toContain('LLM error');
  });

  test('handles chunks with only comments (no amendment)', async () => {
    const paragraphs = [{ text: 'Para 0' }];

    const bookmarkRanges = {
      '_wdpbm0': { text: 'Para 0' },
    };
    const mock = createMockWordRun(paragraphs, bookmarkRanges);
    global.Word.run = mock.wordRun;

    const chunk0 = mockChunk('chunk-0', 0, 'Para 0', 0, 0);

    const results = [
      makeChunkResult('chunk-0', 0, 'fulfilled', { comment: 'Legal review comment', chunk: chunk0 }),
    ];

    const bookmarkMap = new Map([['chunk-0', '_wdpbm0']]);

    const result = await applyChunkResults(results, bookmarkMap, {
      trackChangesEnabled: true,
      lineDiffEnabled: false,
      log: jest.fn(),
    });

    expect(result.amendmentsApplied).toBe(0);
    expect(result.commentsInserted).toBe(1);
    expect(applyTokenMapStrategy).not.toHaveBeenCalled();
  });

  test('handles amendment application error gracefully (records in errors array)', async () => {
    applyTokenMapStrategy.mockRejectedValueOnce(new Error('Word API error'));

    const paragraphs = [{ text: 'Para 0' }];

    const bookmarkRanges = {
      '_wdpbm0': { text: 'Para 0' },
    };
    const mock = createMockWordRun(paragraphs, bookmarkRanges);
    global.Word.run = mock.wordRun;

    const chunk0 = mockChunk('chunk-0', 0, 'Para 0', 0, 0);

    const results = [
      makeChunkResult('chunk-0', 0, 'fulfilled', { amendment: 'Amended text', chunk: chunk0 }),
    ];

    const bookmarkMap = new Map([['chunk-0', '_wdpbm0']]);

    const result = await applyChunkResults(results, bookmarkMap, {
      trackChangesEnabled: true,
      lineDiffEnabled: false,
      log: jest.fn(),
    });

    expect(result.amendmentsApplied).toBe(0);
    expect(result.errors).toHaveLength(1);
    expect(result.errors[0]).toContain('Word API error');
  });

  test('a truncated amendment is refused, never downgraded to a range-level strategy', async () => {
    // Regression: the < 30% guard inside _applyParagraphLevelAmendment used
    // to throw a plain Error, which the strategy dispatcher treated as any
    // other paragraph-level failure and "fell back" to range-level diffing —
    // writing the SAME truncated text in cruder form. A truncation verdict
    // must end the chunk untouched instead.
    const paraItem = {
      text: 'First original paragraph with substantial content',
      parentTableOrNullObject: { isNullObject: true, load: jest.fn() },
      load: jest.fn(),
    };
    const bookmarkRange = {
      text: 'First original paragraph with substantial content',
      isNullObject: false,
      load: jest.fn(),
      paragraphs: { items: [paraItem], load: jest.fn() },
    };
    const mock = createMockWordRun([], {});
    mock.mockContext.document.getBookmarkRangeOrNullObject
      .mockImplementation(() => bookmarkRange);
    global.Word.run = mock.wordRun;

    const chunk = mockChunk('chunk-0', 0, 'First original paragraph with substantial content', 0, 0);
    const results = [
      makeChunkResult('chunk-0', 0, 'fulfilled', { amendment: 'Tiny', chunk }),
    ];
    const onChunkApplied = jest.fn();

    const result = await applyChunkResults(results, new Map([['chunk-0', '_wdpbm0']]), {
      trackChangesEnabled: true,
      lineDiffEnabled: false,
      log: jest.fn(),
      onChunkApplied,
    });

    // No range-level strategy ran with the truncated text.
    expect(applyTokenMapStrategy).not.toHaveBeenCalled();
    expect(applySentenceDiffStrategy).not.toHaveBeenCalled();
    expect(result.amendmentsApplied).toBe(0);
    expect(result.errors).toHaveLength(1);
    expect(result.errors[0]).toContain('truncated');
    expect(onChunkApplied).toHaveBeenCalledWith('chunk-0', expect.objectContaining({ applied: false, error: true }));
  });
});

describe('cleanupBookmarks', () => {
  beforeEach(() => {
    jest.clearAllMocks();
  });

  test('removes all bookmarks from the document', async () => {
    const paragraphs = [{ text: 'Para 0' }];
    const mock = createMockWordRun(paragraphs);
    global.Word.run = mock.wordRun;

    const bookmarkMap = new Map([
      ['chunk-0', '_wdpbm0'],
      ['chunk-1', '_wdpbm1'],
      ['chunk-2', '_wdpbm2'],
    ]);

    await cleanupBookmarks(bookmarkMap);

    expect(mock.deletedBookmarks).toContain('_wdpbm0');
    expect(mock.deletedBookmarks).toContain('_wdpbm1');
    expect(mock.deletedBookmarks).toContain('_wdpbm2');
    expect(mock.deletedBookmarks).toHaveLength(3);
  });

  test('keeps bookmarks named in the keep set (failed chunks stay retryable)', async () => {
    const mock = createMockWordRun([]);
    global.Word.run = mock.wordRun;

    const bookmarkMap = new Map([
      ['chunk-0', '_wdpbm0'],
      ['chunk-1', '_wdpbm1'],
    ]);

    await cleanupBookmarks(bookmarkMap, { keep: new Set(['_wdpbm1']) });

    expect(mock.deletedBookmarks).toEqual(['_wdpbm0']);
  });

  test('handles errors on individual bookmark deletion without stopping', async () => {
    const paragraphs = [{ text: 'Para 0' }];
    const mock = createMockWordRun(paragraphs);

    let callCount = 0;
    mock.mockContext.document.deleteBookmark = jest.fn().mockImplementation((name) => {
      callCount++;
      if (name === '_wdpbm1') {
        throw new Error('Bookmark not found');
      }
      mock.deletedBookmarks.push(name);
    });

    global.Word.run = mock.wordRun;

    const bookmarkMap = new Map([
      ['chunk-0', '_wdpbm0'],
      ['chunk-1', '_wdpbm1'],
      ['chunk-2', '_wdpbm2'],
    ]);

    // Should not throw
    await cleanupBookmarks(bookmarkMap);

    // Should still attempt all 3 deletions
    expect(callCount).toBe(3);
    // The non-erroring ones should still be deleted
    expect(mock.deletedBookmarks).toContain('_wdpbm0');
    expect(mock.deletedBookmarks).toContain('_wdpbm2');
  });
});

describe('_normalizeLineEndings', () => {
  test('converts \\r to \\n', () => {
    expect(_normalizeLineEndings('hello\rworld')).toBe('hello\nworld');
  });

  test('converts \\r\\n to \\n', () => {
    expect(_normalizeLineEndings('hello\r\nworld')).toBe('hello\nworld');
  });

  test('preserves existing \\n', () => {
    expect(_normalizeLineEndings('hello\nworld')).toBe('hello\nworld');
  });

  test('handles mixed line endings', () => {
    expect(_normalizeLineEndings('a\rb\r\nc\nd')).toBe('a\nb\nc\nd');
  });

  test('handles empty string', () => {
    expect(_normalizeLineEndings('')).toBe('');
  });
});

describe('_alignParagraphs', () => {
  test('identical paragraphs: all keep', () => {
    const orig = ['Para 1', 'Para 2', 'Para 3'];
    const amended = ['Para 1', 'Para 2', 'Para 3'];
    const ops = _alignParagraphs(orig, amended);

    expect(ops).toEqual([
      { type: 'keep', origIdx: 0, newIdx: 0 },
      { type: 'keep', origIdx: 1, newIdx: 1 },
      { type: 'keep', origIdx: 2, newIdx: 2 },
    ]);
  });

  test('paragraph deleted: produces delete op', () => {
    const orig = ['Para 1', 'Para 2', 'Para 3'];
    const amended = ['Para 1', 'Para 3'];
    const ops = _alignParagraphs(orig, amended);

    const types = ops.map(o => o.type);
    expect(types).toContain('delete');
    expect(types.filter(t => t === 'keep')).toHaveLength(2);
    // The deleted paragraph should be origIdx 1
    const deleteOp = ops.find(o => o.type === 'delete');
    expect(deleteOp.origIdx).toBe(1);
  });

  test('paragraph inserted: produces insert op', () => {
    const orig = ['Para 1', 'Para 3'];
    const amended = ['Para 1', 'Para 2', 'Para 3'];
    const ops = _alignParagraphs(orig, amended);

    const types = ops.map(o => o.type);
    expect(types).toContain('insert');
    expect(types.filter(t => t === 'keep')).toHaveLength(2);
    const insertOp = ops.find(o => o.type === 'insert');
    expect(insertOp.newIdx).toBe(1);
  });

  test('paragraph with minor edits: matched as keep (similarity-based)', () => {
    const orig = ['Original text here with some content'];
    const amended = ['Modified text here with some content'];
    const ops = _alignParagraphs(orig, amended);

    // High similarity (shared words) -> matched as keep with text changes
    expect(ops).toHaveLength(1);
    expect(ops[0].type).toBe('keep');
    expect(ops[0].origIdx).toBe(0);
    expect(ops[0].newIdx).toBe(0);
  });

  test('completely different paragraph: appears as delete+insert pair', () => {
    const orig = ['Alpha beta gamma'];
    const amended = ['Zeta eta theta'];
    const ops = _alignParagraphs(orig, amended);

    // No shared words -> no similarity -> delete + insert
    expect(ops).toHaveLength(2);
    expect(ops[0].type).toBe('delete');
    expect(ops[1].type).toBe('insert');
  });

  test('CJK paragraph with a one-comma edit: matched as keep (bigram similarity)', () => {
    // CJK text has no whitespace, so word-overlap similarity is always 0;
    // without the bigram path this aligned as delete+insert (whole-paragraph
    // redline) instead of a keep with char-level edits inside.
    const orig = ['但有一个女孩选择不备份自己与母亲的回忆。'];
    const amended = ['但有一个女孩，选择不备份自己与母亲的回忆。'];
    const ops = _alignParagraphs(orig, amended);

    expect(ops).toEqual([{ type: 'keep', origIdx: 0, newIdx: 0 }]);
  });

  test('CJK paragraph that is genuinely rewritten: still delete+insert', () => {
    const orig = ['今天天气很好，我们一起去公园散步。'];
    const amended = ['他最喜欢在深夜写代码，旁边放着一杯咖啡。'];
    const ops = _alignParagraphs(orig, amended);

    expect(ops).toHaveLength(2);
    expect(ops[0].type).toBe('delete');
    expect(ops[1].type).toBe('insert');
  });

  test('handles empty arrays', () => {
    expect(_alignParagraphs([], [])).toEqual([]);
    expect(_alignParagraphs(['a'], [])).toEqual([{ type: 'delete', origIdx: 0 }]);
    expect(_alignParagraphs([], ['b'])).toEqual([{ type: 'insert', newIdx: 0 }]);
  });

  test('trims text for comparison', () => {
    const orig = ['  Para 1  ', '  Para 2  '];
    const amended = ['Para 1', 'Para 2'];
    const ops = _alignParagraphs(orig, amended);

    // Should match on trimmed text
    expect(ops).toEqual([
      { type: 'keep', origIdx: 0, newIdx: 0 },
      { type: 'keep', origIdx: 1, newIdx: 1 },
    ]);
  });
});

describe('_findAnchorWindow', () => {
  test('no drift: stored sequence matches the whole range at offset 0', () => {
    expect(_findAnchorWindow(['A', 'B', 'C'], ['A', 'B', 'C'])).toEqual({ start: 0, end: 3 });
  });

  test('prefix drift: paragraph absorbed at range start (title insertion)', () => {
    expect(_findAnchorWindow(['Title', 'A', 'B'], ['A', 'B'])).toEqual({ start: 1, end: 3 });
  });

  test('suffix drift: paragraph absorbed at range end', () => {
    expect(_findAnchorWindow(['A', 'B', 'Appended'], ['A', 'B'])).toEqual({ start: 0, end: 2 });
  });

  test('prefix + suffix drift', () => {
    expect(_findAnchorWindow(['Title', 'A', 'B', 'Tail'], ['A', 'B'])).toEqual({ start: 1, end: 3 });
  });

  test('middle insertion breaks contiguity: returns null', () => {
    expect(_findAnchorWindow(['A', 'Inserted', 'B'], ['A', 'B'])).toBeNull();
  });

  test('stored longer than current: returns null', () => {
    expect(_findAnchorWindow(['A'], ['A', 'B'])).toBeNull();
  });

  test('empty stored sequence: returns null', () => {
    expect(_findAnchorWindow(['A'], [])).toBeNull();
  });

  test('compares trimmed text', () => {
    expect(_findAnchorWindow(['Title', '  A  ', ' B '], ['A', 'B'])).toEqual({ start: 1, end: 3 });
  });
});

// --- Re-anchoring scenario tests ---
// Regression: a staged card's bookmark range absorbs paragraphs inserted
// after staging (e.g. a title from another card). Applying the staged
// amendment then treated the absorbed paragraph as LLM-deleted and removed
// it. The apply step must re-anchor to the stored original paragraphs.

/**
 * Builds a mock Word.run whose bookmark range exposes a real paragraphs
 * collection (unlike createMockWordRun, which exercises the fallback path).
 * expandTo on a paragraph Start range produces a narrowed range whose
 * paragraphs are the [paraIdx..other.paraIdx] window, mirroring Word.
 */
function createParagraphAwareMockRun(currentParaTexts) {
  const syncFn = jest.fn().mockResolvedValue(undefined);

  // Entries may be plain strings (outside tables) or { text, inTable } specs.
  const specs = currentParaTexts.map((p) => (typeof p === 'string' ? { text: p, inTable: false } : p));
  const paraTexts = specs.map((s) => s.text);

  function makeNarrowedRange(startIdx, endIdx) {
    const windowItems = items.slice(startIdx, endIdx + 1);
    return {
      text: windowItems.map((p) => p.text).join('\n'),
      isNullObject: false,
      load: jest.fn(),
      paragraphs: { items: windowItems, load: jest.fn() },
    };
  }

  function makeSubRange(paraIdx) {
    return {
      paraIdx,
      text: paraTexts[paraIdx],
      load: jest.fn(),
      insertText: jest.fn(),
      expandTo: jest.fn((other) => makeNarrowedRange(paraIdx, other.paraIdx)),
    };
  }

  const items = specs.map((spec, i) => ({
    text: spec.text,
    load: jest.fn(),
    delete: jest.fn(),
    insertParagraph: jest.fn(),
    parentTableOrNullObject: { isNullObject: !spec.inTable, load: jest.fn() },
    getRange: jest.fn(() => makeSubRange(i)),
  }));

  const bookmarkRange = {
    text: paraTexts.join('\n'),
    isNullObject: false,
    load: jest.fn(),
    paragraphs: { items, load: jest.fn() },
    insertComment: jest.fn(),
  };

  const mockContext = {
    document: {
      body: { paragraphs: { items: [], load: jest.fn() } },
      getBookmarkRangeOrNullObject: jest.fn(() => bookmarkRange),
      deleteBookmark: jest.fn(),
    },
    sync: syncFn,
  };

  let _trackingMode = null;
  Object.defineProperty(mockContext.document, 'changeTrackingMode', {
    get: () => _trackingMode,
    set: (val) => { _trackingMode = val; },
  });

  const wordRun = jest.fn().mockImplementation(async (callback) => {
    await callback(mockContext);
  });

  return { wordRun, mockContext, items, bookmarkRange };
}

function driftChunk(id, paraTexts, startIndex, endIndex) {
  return {
    id,
    paragraphs: paraTexts.map((text, k) => ({ index: startIndex + k, text, headingLevel: 0 })),
    startIndex,
    endIndex,
    tokenCount: 100,
  };
}

describe('applyChunkResults re-anchoring', () => {
  beforeEach(() => {
    jest.clearAllMocks();
  });

  test('prefix drift: absorbed title paragraph is not deleted by the amendment', async () => {
    // Card 1 inserted a title at document start; the chunk-0 bookmark range
    // absorbed it. Card 2's amendment was generated against the original
    // paragraphs only and must not touch the absorbed title.
    const mock = createParagraphAwareMockRun(['A New Title', 'Para one text.', 'Para two text.']);
    global.Word.run = mock.wordRun;

    const chunk = driftChunk('chunk-0', ['Para one text.', 'Para two text.'], 0, 1);
    const results = [
      makeChunkResult('chunk-0', 0, 'fulfilled', {
        amendment: 'Para one text.\nPara two revised text.',
        chunk,
      }),
    ];
    const bookmarkMap = new Map([['chunk-0', '_wdpbm0']]);

    const outcome = await applyChunkResults(results, bookmarkMap, {
      trackChangesEnabled: true,
      lineDiffEnabled: false,
      log: jest.fn(),
    });

    expect(outcome.errors).toHaveLength(0);
    expect(outcome.amendmentsApplied).toBe(1);
    // The title paragraph must be untouched: no delete, no diff, no insert.
    expect(mock.items[0].delete).not.toHaveBeenCalled();
    expect(mock.items[0].insertParagraph).not.toHaveBeenCalled();
    expect(mock.items[0].getRange).not.toHaveBeenCalled();
    // Paragraph two changed -> exactly one word-level diff call on it.
    expect(applyTokenMapStrategy).toHaveBeenCalledTimes(1);
    expect(applyTokenMapStrategy.mock.calls[0][2]).toBe('Para two text.');
    expect(applyTokenMapStrategy.mock.calls[0][3]).toBe('Para two revised text.');
  });

  test('suffix drift: absorbed trailing paragraph is not deleted', async () => {
    const mock = createParagraphAwareMockRun(['Para one text.', 'Para two text.', 'Appended paragraph']);
    global.Word.run = mock.wordRun;

    const chunk = driftChunk('chunk-0', ['Para one text.', 'Para two text.'], 0, 1);
    const results = [
      makeChunkResult('chunk-0', 0, 'fulfilled', {
        amendment: 'Para one text.\nPara two revised text.',
        chunk,
      }),
    ];
    const bookmarkMap = new Map([['chunk-0', '_wdpbm0']]);

    const outcome = await applyChunkResults(results, bookmarkMap, {
      trackChangesEnabled: true,
      lineDiffEnabled: false,
      log: jest.fn(),
    });

    expect(outcome.errors).toHaveLength(0);
    expect(outcome.amendmentsApplied).toBe(1);
    expect(mock.items[2].delete).not.toHaveBeenCalled();
    expect(applyTokenMapStrategy).toHaveBeenCalledTimes(1);
  });

  test('no contiguous match (middle insertion): chunk skipped, nothing deleted', async () => {
    const mock = createParagraphAwareMockRun(['Para one text.', 'Inserted in middle', 'Para two text.']);
    global.Word.run = mock.wordRun;

    const chunk = driftChunk('chunk-0', ['Para one text.', 'Para two text.'], 0, 1);
    const results = [
      makeChunkResult('chunk-0', 0, 'fulfilled', {
        amendment: 'Para one text.\nPara two revised text.',
        chunk,
      }),
    ];
    const bookmarkMap = new Map([['chunk-0', '_wdpbm0']]);

    const outcome = await applyChunkResults(results, bookmarkMap, {
      trackChangesEnabled: true,
      lineDiffEnabled: false,
      log: jest.fn(),
    });

    expect(outcome.amendmentsApplied).toBe(0);
    expect(outcome.errors).toHaveLength(1);
    expect(outcome.errors[0]).toMatch(/no longer matches/);
    expect(applyTokenMapStrategy).not.toHaveBeenCalled();
    for (const item of mock.items) {
      expect(item.delete).not.toHaveBeenCalled();
    }
  });

  test('no drift: amendment applies to the whole range as before', async () => {
    const mock = createParagraphAwareMockRun(['Para one text.', 'Para two text.']);
    global.Word.run = mock.wordRun;

    const chunk = driftChunk('chunk-0', ['Para one text.', 'Para two text.'], 0, 1);
    const results = [
      makeChunkResult('chunk-0', 0, 'fulfilled', {
        amendment: 'Para one text.\nPara two revised text.',
        chunk,
      }),
    ];
    const bookmarkMap = new Map([['chunk-0', '_wdpbm0']]);

    const outcome = await applyChunkResults(results, bookmarkMap, {
      trackChangesEnabled: true,
      lineDiffEnabled: false,
      log: jest.fn(),
    });

    expect(outcome.errors).toHaveLength(0);
    expect(outcome.amendmentsApplied).toBe(1);
    expect(applyTokenMapStrategy).toHaveBeenCalledTimes(1);
  });

  test('chunkOriginals option supplies staged texts when the result chunk has no paragraphs (retry path)', async () => {
    const mock = createParagraphAwareMockRun(['A New Title', 'Para one text.', 'Para two text.']);
    global.Word.run = mock.wordRun;

    const results = [
      makeChunkResult('chunk-0', 0, 'fulfilled', {
        amendment: 'Para one text.\nPara two revised text.',
        chunk: { id: 'chunk-0', startIndex: 0, endIndex: 1 }, // rebuilt retry chunk: no paragraphs
      }),
    ];
    const bookmarkMap = new Map([['chunk-0', '_wdpbm0']]);
    const chunkOriginals = new Map([['chunk-0', ['Para one text.', 'Para two text.']]]);

    const outcome = await applyChunkResults(results, bookmarkMap, {
      trackChangesEnabled: true,
      lineDiffEnabled: false,
      log: jest.fn(),
      chunkOriginals,
    });

    expect(outcome.errors).toHaveLength(0);
    expect(outcome.amendmentsApplied).toBe(1);
    expect(mock.items[0].delete).not.toHaveBeenCalled();
    expect(applyTokenMapStrategy).toHaveBeenCalledTimes(1);
  });

  test('re-anchor check failure: chunk skipped safely instead of falling back to the raw range', async () => {
    const mock = createParagraphAwareMockRun(['A New Title', 'Para one text.', 'Para two text.']);
    // Simulate a Word API failure while reading the range's paragraphs.
    mock.bookmarkRange.paragraphs.load = jest.fn(() => { throw new Error('boom'); });
    global.Word.run = mock.wordRun;

    const chunk = driftChunk('chunk-0', ['Para one text.', 'Para two text.'], 0, 1);
    const results = [
      makeChunkResult('chunk-0', 0, 'fulfilled', {
        amendment: 'Para one text.\nPara two revised text.',
        chunk,
      }),
    ];
    const bookmarkMap = new Map([['chunk-0', '_wdpbm0']]);

    const outcome = await applyChunkResults(results, bookmarkMap, {
      trackChangesEnabled: true,
      lineDiffEnabled: false,
      log: jest.fn(),
    });

    expect(outcome.amendmentsApplied).toBe(0);
    expect(outcome.errors).toHaveLength(1);
    expect(outcome.errors[0]).toMatch(/re-anchor check failed/);
    expect(applyTokenMapStrategy).not.toHaveBeenCalled();
    for (const item of mock.items) {
      expect(item.delete).not.toHaveBeenCalled();
    }
  });

  test('blank paragraphs are ignored by anchoring and never deleted (title + blank lines)', async () => {
    // The exact user scenario: the range absorbed a tracked-inserted title
    // AND the document has blank spacer paragraphs (which the parser skips,
    // so the staged sequence never contains them).
    const mock = createParagraphAwareMockRun(['A New Title', 'Para one text.', '', 'Para two text.', '']);
    global.Word.run = mock.wordRun;

    const chunk = driftChunk('chunk-0', ['Para one text.', 'Para two text.'], 0, 1);
    const results = [
      makeChunkResult('chunk-0', 0, 'fulfilled', {
        amendment: 'Para one text.\nPara two revised text.',
        chunk,
      }),
    ];
    const bookmarkMap = new Map([['chunk-0', '_wdpbm0']]);

    const outcome = await applyChunkResults(results, bookmarkMap, {
      trackChangesEnabled: true,
      lineDiffEnabled: false,
      log: jest.fn(),
    });

    expect(outcome.errors).toHaveLength(0);
    expect(outcome.amendmentsApplied).toBe(1);
    // Title and both blank paragraphs must survive untouched.
    expect(mock.items[0].delete).not.toHaveBeenCalled();
    expect(mock.items[2].delete).not.toHaveBeenCalled();
    expect(mock.items[4].delete).not.toHaveBeenCalled();
    expect(applyTokenMapStrategy).toHaveBeenCalledTimes(1);
    expect(applyTokenMapStrategy.mock.calls[0][2]).toBe('Para two text.');
  });

  test('blank paragraphs are preserved even without drift', async () => {
    const mock = createParagraphAwareMockRun(['Para one text.', '', 'Para two text.']);
    global.Word.run = mock.wordRun;

    const chunk = driftChunk('chunk-0', ['Para one text.', 'Para two text.'], 0, 1);
    const results = [
      makeChunkResult('chunk-0', 0, 'fulfilled', {
        amendment: 'Para one text.\nPara two revised text.',
        chunk,
      }),
    ];
    const bookmarkMap = new Map([['chunk-0', '_wdpbm0']]);

    const outcome = await applyChunkResults(results, bookmarkMap, {
      trackChangesEnabled: true,
      lineDiffEnabled: false,
      log: jest.fn(),
    });

    expect(outcome.errors).toHaveLength(0);
    expect(outcome.amendmentsApplied).toBe(1);
    expect(mock.items[1].delete).not.toHaveBeenCalled();
    expect(applyTokenMapStrategy).toHaveBeenCalledTimes(1);
  });

  test('all-blank range: nothing to amend, no changes counted', async () => {
    const mock = createParagraphAwareMockRun(['', '']);
    global.Word.run = mock.wordRun;

    // Chunk without paragraphs: no stored texts -> no anchoring, straight
    // to the paragraph-level path, which finds only blanks.
    const results = [
      makeChunkResult('chunk-0', 0, 'fulfilled', {
        amendment: 'some new text',
        chunk: { id: 'chunk-0', startIndex: 0, endIndex: 1 },
      }),
    ];
    const bookmarkMap = new Map([['chunk-0', '_wdpbm0']]);

    const outcome = await applyChunkResults(results, bookmarkMap, {
      trackChangesEnabled: true,
      lineDiffEnabled: false,
      log: jest.fn(),
    });

    expect(outcome.errors).toHaveLength(0);
    expect(outcome.amendmentsApplied).toBe(0);
    expect(outcome.noChangeCount).toBe(1);
    expect(applyTokenMapStrategy).not.toHaveBeenCalled();
  });

  test('aborted signal pauses between chunks: interrupted with partial appliedChunkIds', async () => {
    const paragraphs = [
      { text: 'Para 0' }, { text: 'Para 1' }, { text: 'Para 2' },
      { text: 'Para 3' }, { text: 'Para 4' }, { text: 'Para 5' },
    ];
    const bookmarkRanges = {
      '_wdpbm0': { text: 'Para 0\nPara 1\nPara 2' },
      '_wdpbm1': { text: 'Para 3\nPara 4\nPara 5' },
    };
    const mock = createMockWordRun(paragraphs, bookmarkRanges);
    global.Word.run = mock.wordRun;

    const results = [
      makeChunkResult('chunk-0', 0, 'fulfilled', {
        amendment: 'Amended 0', chunk: mockChunk('chunk-0', 0, 'Para 0\nPara 1\nPara 2', 0, 2),
      }),
      makeChunkResult('chunk-1', 1, 'fulfilled', {
        amendment: 'Amended 1', chunk: mockChunk('chunk-1', 1, 'Para 3\nPara 4\nPara 5', 3, 5),
      }),
    ];
    const bookmarkMap = new Map([
      ['chunk-0', '_wdpbm0'],
      ['chunk-1', '_wdpbm1'],
    ]);

    // Reverse order applies chunk-1 first, then chunk-0. Abort after the
    // first chunk lands so the second is skipped.
    const controller = new AbortController();
    const appliedChunkIds = [];
    const applied = [];
    const outcome = await applyChunkResults(results, bookmarkMap, {
      trackChangesEnabled: true,
      lineDiffEnabled: false,
      log: jest.fn(),
      signal: controller.signal,
      onChunkApplied: (id, status) => {
        appliedChunkIds.push(id);
        applied.push(status);
        controller.abort(); // pause right after the first chunk finishes
      },
    });

    expect(outcome.interrupted).toBe(true);
    expect(outcome.appliedChunkIds).toHaveLength(1);
    expect(outcome.appliedChunkIds[0]).toBe('chunk-1'); // reverse order → first
    expect(applied).toHaveLength(1);
    expect(applied[0]).toMatchObject({ applied: true });
    // The remaining chunk was NOT attempted (signal aborted before its turn).
    expect(outcome.amendmentsApplied).toBe(1);
  });

  test('no signal / no interruption returns interrupted:false', async () => {
    const mock = createParagraphAwareMockRun(['a']);
    global.Word.run = mock.wordRun;
    const results = [
      makeChunkResult('chunk-0', 0, 'fulfilled', {
        amendment: 'b', chunk: { id: 'chunk-0', startIndex: 0, endIndex: 0 },
      }),
    ];
    const outcome = await applyChunkResults(results, new Map([['chunk-0', '_wdpbm0']]), {
      trackChangesEnabled: true,
      lineDiffEnabled: false,
      log: jest.fn(),
    });
    expect(outcome.interrupted).toBe(false);
    expect(outcome.appliedChunkIds).toEqual(['chunk-0']);
  });
});

// --- Table paragraph guards ---

describe('applyChunkResults table-paragraph guards', () => {
  beforeEach(() => {
    jest.clearAllMocks();
  });

  test('in-cell keep edits still diff granularly; insert anchored at a table paragraph is skipped', async () => {
    const mock = createParagraphAwareMockRun([
      'Intro',
      { text: 'Cell A1 text', inTable: true },
      'Outro',
    ]);
    global.Word.run = mock.wordRun;
    const log = jest.fn();

    const chunk = driftChunk('chunk-0', ['Intro', 'Cell A1 text', 'Outro'], 0, 2);
    const results = [
      makeChunkResult('chunk-0', 0, 'fulfilled', {
        amendment: 'Intro\nCell A1 text revised\nBrand new paragraph\nOutro',
        chunk,
      }),
    ];
    const bookmarkMap = new Map([['chunk-0', '_wdpbm0']]);

    const outcome = await applyChunkResults(results, bookmarkMap, {
      trackChangesEnabled: true,
      lineDiffEnabled: false,
      log,
    });

    expect(outcome.errors).toHaveLength(0);
    expect(outcome.amendmentsApplied).toBe(1);
    // The in-cell text edit went through the granular strategy as usual.
    expect(applyTokenMapStrategy).toHaveBeenCalledTimes(1);
    expect(applyTokenMapStrategy.mock.calls[0][2]).toBe('Cell A1 text');
    expect(applyTokenMapStrategy.mock.calls[0][3]).toBe('Cell A1 text revised');
    // The new paragraph would anchor inside the table — skipped, not inserted.
    for (const item of mock.items) {
      expect(item.insertParagraph).not.toHaveBeenCalled();
      expect(item.delete).not.toHaveBeenCalled();
    }
    const warnings = log.mock.calls.filter((c) => c[1] === 'warning').map((c) => c[0]).join('\n');
    expect(warnings).toMatch(/Skipping insert after para 1/);
  });

  test('delete op on a table paragraph is skipped (cell content is not a row)', async () => {
    const mock = createParagraphAwareMockRun([
      'Intro',
      { text: 'Cell A1 text', inTable: true },
      'Outro',
    ]);
    global.Word.run = mock.wordRun;
    const log = jest.fn();

    const chunk = driftChunk('chunk-0', ['Intro', 'Cell A1 text', 'Outro'], 0, 2);
    const results = [
      makeChunkResult('chunk-0', 0, 'fulfilled', {
        amendment: 'Intro\nOutro',
        chunk,
      }),
    ];
    const bookmarkMap = new Map([['chunk-0', '_wdpbm0']]);

    const outcome = await applyChunkResults(results, bookmarkMap, {
      trackChangesEnabled: true,
      lineDiffEnabled: false,
      log,
    });

    expect(outcome.errors).toHaveLength(0);
    expect(mock.items[1].delete).not.toHaveBeenCalled();
    expect(applyTokenMapStrategy).not.toHaveBeenCalled();
    const warnings = log.mock.calls.filter((c) => c[1] === 'warning').map((c) => c[0]).join('\n');
    expect(warnings).toMatch(/skipping delete/);
  });

  test('insert anchored at a non-table paragraph still inserts', async () => {
    const mock = createParagraphAwareMockRun([
      'Intro',
      { text: 'Cell A1 text', inTable: true },
      'Outro',
    ]);
    global.Word.run = mock.wordRun;

    const chunk = driftChunk('chunk-0', ['Intro', 'Cell A1 text', 'Outro'], 0, 2);
    const results = [
      makeChunkResult('chunk-0', 0, 'fulfilled', {
        amendment: 'Intro\nCell A1 text\nOutro\nBrand new paragraph',
        chunk,
      }),
    ];
    const bookmarkMap = new Map([['chunk-0', '_wdpbm0']]);

    const outcome = await applyChunkResults(results, bookmarkMap, {
      trackChangesEnabled: true,
      lineDiffEnabled: false,
      log: jest.fn(),
    });

    expect(outcome.errors).toHaveLength(0);
    expect(mock.items[2].insertParagraph).toHaveBeenCalledWith('Brand new paragraph', 'After');
  });
});
