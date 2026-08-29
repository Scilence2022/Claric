/**
 * Bookmark hygiene tests (reassembler).
 *
 * Covers two hardening fixes:
 * 1. bookmarkChunkRanges anchor validation — chunk paragraph indexes come
 *    from a PREVIOUS Word.run (parseDocument). If paragraphs were inserted
 *    or deleted in between (a concurrent card apply, a comment-queue flush,
 *    the user typing), the index silently points at the wrong paragraph and
 *    amendments land on the wrong text. The boundaries must be verified
 *    against the chunk's expected first/last paragraph texts.
 * 2. reapOrphanChunkBookmarks — chunk bookmarks are only removed by an
 *    in-memory apply/discard closure; after a taskpane reload those are
 *    gone and _wdp* bookmarks would leak into the document forever.
 */

const {
    bookmarkChunkRanges,
    reapOrphanChunkBookmarks,
} = require('../src/lib/reassembler.js');

/** Builds a Word.run mock with paragraph text proxies and bookmark capture. */
function setupWord({ paragraphTexts = [], bookmarkNames = [] } = {}) {
    const inserted = [];
    const deleted = [];

    const paraProxies = paragraphTexts.map((text) => ({
        text,
        load: jest.fn(),
        getRange: jest.fn((location) => ({
            location,
            expandTo: jest.fn(() => ({
                insertBookmark: jest.fn((name) => inserted.push(name)),
            })),
        })),
    }));

    const context = {
        document: {
            body: {
                paragraphs: { items: paraProxies, load: jest.fn() },
                getRange: jest.fn(() => ({
                    getBookmarks: jest.fn(() => ({ value: bookmarkNames })),
                })),
            },
            deleteBookmark: jest.fn((name) => deleted.push(name)),
        },
        sync: jest.fn().mockResolvedValue(undefined),
    };
    global.Word = { run: async (fn) => fn(context), ChangeTrackingMode: { off: 'Off' } };
    return { inserted, deleted };
}

function makeChunk(id, startIndex, endIndex, texts) {
    return {
        id,
        startIndex,
        endIndex,
        paragraphs: texts.map((text) => ({ text })),
    };
}

afterEach(() => {
    delete global.Word;
});

describe('bookmarkChunkRanges anchor validation', () => {
    test('bookmarks ranges when boundary paragraphs still match the chunk texts', async () => {
        setupWord({ paragraphTexts: ['First para.', 'Middle para.', 'Last para.'] });
        const chunks = [makeChunk('c1', 0, 2, ['First para.', 'Middle para.', 'Last para.'])];

        const map = await bookmarkChunkRanges(chunks);

        expect(map.get('c1')).toMatch(/^_wdp/);
    });

    test('throws honestly when paragraph indexes shifted (document changed)', async () => {
        // Chunk was parsed when paragraph 0 read "Original opening."; by
        // bookmark time it reads "Inserted title." — the index drifted.
        setupWord({ paragraphTexts: ['Inserted title.', 'Middle para.', 'Last para.'] });
        const chunks = [makeChunk('c1', 0, 2, ['Original opening.', 'Middle para.', 'Last para.'])];

        await expect(bookmarkChunkRanges(chunks)).rejects.toThrow(/Document changed while staging/);
    });

    test('throws when the document shrank below a chunk index', async () => {
        setupWord({ paragraphTexts: ['Only one paragraph left.'] });
        const chunks = [makeChunk('c1', 0, 5, ['A', 'B', 'C', 'D', 'E', 'F'])];

        await expect(bookmarkChunkRanges(chunks)).rejects.toThrow(/no longer exists/);
    });
});

describe('reapOrphanChunkBookmarks', () => {
    test('deletes only hidden _wdp* bookmarks and reports the count', async () => {
        const { deleted } = setupWord({
            bookmarkNames: ['_wdp18f3a0', 'ChapterRef', '_wdp19abcd'],
        });

        const removed = await reapOrphanChunkBookmarks();

        expect(removed).toBe(2);
        expect(deleted.sort()).toEqual(['_wdp18f3a0', '_wdp19abcd']);
    });

    test('is a no-op on hosts without Range.getBookmarks (graceful skip)', async () => {
        const { deleted } = setupWord({});
        // Strip the API to emulate an older host.
        global.Word.run = async (fn) => fn({
            document: {
                body: {
                    getRange: () => ({}), // no getBookmarks
                },
                deleteBookmark: (n) => deleted.push(n),
            },
            sync: () => Promise.resolve(),
        });
        const log = jest.fn();

        const removed = await reapOrphanChunkBookmarks(log);

        expect(removed).toBe(0);
        expect(deleted).toHaveLength(0);
        expect(log).toHaveBeenCalledWith(expect.stringContaining('skipped'), 'warning');
    });
});
