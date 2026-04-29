/**
 * Wave 0 RED specs for src/lib/comment-extractor.js : extractCommentsOnRange (Phase 05.2-R1).
 *
 * These tests are intentionally failing — the function under test is not yet implemented.
 * Plan 02 (Wave 1) lands the implementation and turns these GREEN.
 *
 * Failure mode while RED: `extractCommentsOnRange is not a function` (destructured
 * undefined export from an existing module — no MODULE_NOT_FOUND, no SyntaxError).
 *
 * Mirrors the Jest harness from tests/comment-extractor.spec.js:1-12.
 */

const { JSDOM } = require('jsdom');
const { extractCommentsOnRange } = require('../src/lib/comment-extractor.js');

// Provide DOMParser for OOXML tests (node test environment lacks it)
if (typeof globalThis.DOMParser === 'undefined') {
    const dom = new JSDOM('');
    globalThis.DOMParser = dom.window.DOMParser;
}

// ============================================================================
// LocationRelation enums (STYLE.md "No Magic Strings")
// Mirrors INTERSECTING_RELATIONS / EXCLUDED set Plan 02 will export.
// ============================================================================

/** @type {ReadonlyArray<string>} 9 relations a comment must be KEPT for. */
const TEST_INTERSECTING_RELATIONS = Object.freeze([
    'Equal',
    'Inside',
    'InsideStart',
    'InsideEnd',
    'Contains',
    'ContainsStart',
    'ContainsEnd',
    'OverlapsBefore',
    'OverlapsAfter',
]);

/** @type {ReadonlyArray<string>} 5 relations a comment must be DROPPED for. */
const TEST_EXCLUDED_RELATIONS = Object.freeze([
    'Before',
    'After',
    'AdjacentBefore',
    'AdjacentAfter',
    'Unrelated',
]);

// ============================================================================
// Mock factory
// ============================================================================

/**
 * Build an Office.js-shaped mock comment whose getRange().compareLocationWith()
 * returns a deferred ClientResult-shaped object resolving to `locationRelation`
 * after `context.sync()`.
 *
 * @param {object} opts
 * @param {string} [opts.id]
 * @param {string} [opts.content]
 * @param {string} [opts.authorName]
 * @param {string} [opts.creationDate]
 * @param {boolean} [opts.resolved]
 * @param {string} [opts.locationRelation]
 * @param {Array<object>} [opts.replies]
 * @returns {object} Office.js-shaped mock Comment
 */
function createMockComment({
    id = 'c0',
    content = '',
    authorName = '',
    creationDate = '2026-01-01T00:00:00Z',
    resolved = false,
    locationRelation = 'Inside',
    replies = [],
} = {}) {
    const comparedLocation = { value: undefined, _pending: locationRelation };
    const range = {
        load: jest.fn(),
        compareLocationWith: jest.fn(() => comparedLocation),
    };

    // The reply collection itself
    const repliesCollection = {
        items: replies,
        load: jest.fn(),
    };

    const comment = {
        id,
        content,
        authorName,
        creationDate,
        resolved,
        load: jest.fn(),
        getRange: jest.fn(() => range),
        getReplies: jest.fn(() => repliesCollection),
        replies: repliesCollection,
        _comparedLocation: comparedLocation, // exposed for assertions
        _range: range,
    };
    return comment;
}

/**
 * Build a mock CommentReply with no `resolved` property (per research correction #2).
 *
 * @param {object} opts
 * @returns {object} Office.js-shaped mock CommentReply
 */
function createMockReply({ id = 'r0', content = '', authorName = '', creationDate = '2026-01-01T00:00:00Z' } = {}) {
    return {
        id,
        content,
        authorName,
        creationDate,
        // Spy used by the "does NOT load resolved on CommentReply" test.
        load: jest.fn(),
    };
}

/**
 * Build a context whose `sync()` resolves all pending compareLocationWith values
 * (mimics ClientResult settling after sync).
 *
 * @param {Array<object>} commentItems - mock Comment objects
 * @param {object} [opts]
 * @param {object} [opts.selectionRange] - optional explicit selectionRange spy
 * @returns {{context: object, body: object, selectionRange: object, syncCalls: {n: number}}}
 */
function createContext(commentItems, { selectionRange } = {}) {
    const syncCalls = { n: 0 };
    const body = {
        getComments: jest.fn(() => ({ items: commentItems, load: jest.fn() })),
    };
    const sel = selectionRange || {
        // Anti-pattern guard: extractor MUST NOT call selectionRange.getComments()
        // (research § Anti-Patterns). Spy installed so the assertion is exact.
        getComments: jest.fn(() => {
            throw new Error('selectionRange.getComments() must not be called — use body.getComments()');
        }),
    };
    const context = {
        document: { body },
        sync: jest.fn(async () => {
            syncCalls.n += 1;
            // Settle every comment's deferred locationRelation (ClientResult.value
            // becomes available only after sync per research Pitfall 1).
            for (const c of commentItems) {
                if (c && c._comparedLocation && c._comparedLocation._pending !== undefined) {
                    c._comparedLocation.value = c._comparedLocation._pending;
                }
            }
        }),
    };
    return { context, body, selectionRange: sel, syncCalls };
}

beforeEach(() => {
    global.Word = {
        run: jest.fn(async (rangeOrCallback, maybeCallback) => {
            // Support both Word.run(callback) and Word.run(range, callback) shapes
            // per research § Pattern 1 / Open Q2.
            if (typeof rangeOrCallback === 'function') {
                return rangeOrCallback(/* test never reaches here directly */);
            }
            return maybeCallback(/* test never reaches here directly */);
        }),
    };
});

afterEach(() => {
    delete global.Word;
});

// ============================================================================
// Tests
// ============================================================================

describe('extractCommentsOnRange (Phase 05.2-R1)', () => {
    it.each(TEST_INTERSECTING_RELATIONS)(
        'intersection filter: includes comments whose LocationRelation is %s (kept)',
        async (relation) => {
            const comment = createMockComment({ id: 'c-' + relation, locationRelation: relation });
            const { context, selectionRange } = createContext([comment]);

            const result = await extractCommentsOnRange(context, selectionRange);

            expect(Array.isArray(result)).toBe(true);
            expect(result).toHaveLength(1);
            expect(result[0].id).toBe('c-' + relation);
        }
    );

    it.each(TEST_EXCLUDED_RELATIONS)(
        'intersection filter: drops comments whose LocationRelation is %s (excluded)',
        async (relation) => {
            const comment = createMockComment({ id: 'c-' + relation, locationRelation: relation });
            const { context, selectionRange } = createContext([comment]);

            const result = await extractCommentsOnRange(context, selectionRange);

            expect(Array.isArray(result)).toBe(true);
            expect(result).toHaveLength(0);
        }
    );

    it('resolved comments are included (no filter on comment.resolved)', async () => {
        const comment = createMockComment({
            id: 'c-resolved',
            resolved: true,
            locationRelation: 'Inside',
        });
        const { context, selectionRange } = createContext([comment]);

        const result = await extractCommentsOnRange(context, selectionRange);

        expect(result).toHaveLength(1);
        expect(result[0].resolved).toBe(true);
    });

    it('reply order: replies sorted by creationDate ascending defensively', async () => {
        const replyLate = createMockReply({ id: 'r-late', creationDate: '2026-03-01T00:00:00Z' });
        const replyMid = createMockReply({ id: 'r-mid', creationDate: '2026-02-01T00:00:00Z' });
        const replyEarly = createMockReply({ id: 'r-early', creationDate: '2026-01-01T00:00:00Z' });
        // Provided in REVERSE chronological order — extractor must re-sort ascending.
        const comment = createMockComment({
            id: 'c-thread',
            locationRelation: 'Inside',
            replies: [replyLate, replyMid, replyEarly],
        });
        const { context, selectionRange } = createContext([comment]);

        const result = await extractCommentsOnRange(context, selectionRange);

        expect(result).toHaveLength(1);
        const ids = result[0].replies.map((r) => r.id);
        expect(ids).toEqual(['r-early', 'r-mid', 'r-late']);
    });

    it('does NOT load resolved on CommentReply (research correction #2)', async () => {
        const reply = createMockReply({ id: 'r-only' });
        const comment = createMockComment({
            id: 'c-thread',
            locationRelation: 'Inside',
            replies: [reply],
        });
        const { context, selectionRange } = createContext([comment]);

        await extractCommentsOnRange(context, selectionRange);

        // Inspect every load() call made on the reply: none of the load specs
        // may include 'resolved' (CommentReply has no resolved property).
        for (const call of reply.load.mock.calls) {
            const arg = call[0];
            if (typeof arg === 'string') {
                expect(arg.split(/[,\s/]+/)).not.toContain('resolved');
            } else if (arg && typeof arg === 'object') {
                expect(Object.prototype.hasOwnProperty.call(arg, 'resolved')).toBe(false);
            }
        }
    });

    it('uses body.getComments(), NOT selectionRange.getComments()', async () => {
        const comment = createMockComment({ id: 'c-1', locationRelation: 'Inside' });
        const { context, body, selectionRange } = createContext([comment]);

        await extractCommentsOnRange(context, selectionRange);

        expect(body.getComments).toHaveBeenCalled();
        expect(selectionRange.getComments).not.toHaveBeenCalled();
    });

    it('reads compareLocationWith via .value AFTER context.sync() (research Pitfall 1)', async () => {
        const comment = createMockComment({ id: 'c-1', locationRelation: 'Inside' });
        const { context, selectionRange, syncCalls } = createContext([comment]);

        // Before extractor runs: deferred value is unset.
        expect(comment._comparedLocation.value).toBeUndefined();

        const result = await extractCommentsOnRange(context, selectionRange);

        // After extractor runs: at least one sync occurred, .value populated, comment kept.
        expect(syncCalls.n).toBeGreaterThanOrEqual(1);
        expect(comment._comparedLocation.value).toBe('Inside');
        expect(result).toHaveLength(1);
    });
});
