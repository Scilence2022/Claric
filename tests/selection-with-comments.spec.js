/**
 * Specs for src/lib/selection-with-comments.js : formatSelectionWithComments
 * (Phase 05.2-R2 — splicer behavior).
 *
 * Covers the OOXML splicer: how comment anchors overlapping the selection are
 * rendered as annotation envelopes — fully inside, head/tail/both truncated,
 * nested ranges, resolved status, tracked changes, and no-comments passthrough.
 *
 * 7 OOXML fixtures inlined as JS template strings per RESEARCH.md
 * § OOXML Fixture Sketches (lines 706–786) — fixtures readable next to
 * assertions, no separate .xml files.
 */

const { JSDOM } = require('jsdom');

if (typeof globalThis.DOMParser === 'undefined') {
    const dom = new JSDOM('');
    globalThis.DOMParser = dom.window.DOMParser;
}

const { formatSelectionWithComments } = require('../src/lib/selection-with-comments.js');

// ---------------------------------------------------------------------------
// OOXML namespace constants (mirror src/lib/comment-extractor.js:16-17)
// ---------------------------------------------------------------------------

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const PKG_NS = 'http://schemas.microsoft.com/office/2006/xmlPackage';

/**
 * Wrap a `<w:body>` payload in the canonical pkg:package envelope so the
 * existing extractDocumentBody helper from comment-extractor.js can resolve
 * the body element. Per RESEARCH.md fixture 1, every fixture must be a
 * complete pkg:package wrap.
 *
 * @param {string} bodyInner - raw inner body XML (one or more <w:p>...).
 * @returns {string} full pkg:package OOXML string.
 */
function wrapPackage(bodyInner) {
    return (
        `<pkg:package xmlns:pkg="${PKG_NS}">` +
        `<pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">` +
        `<pkg:xmlData>` +
        `<w:document xmlns:w="${W_NS}"><w:body>${bodyInner}</w:body></w:document>` +
        `</pkg:xmlData></pkg:part></pkg:package>`
    );
}

// ---------------------------------------------------------------------------
// 7 inlined OOXML fixtures — RESEARCH.md § OOXML Fixture Sketches
// ---------------------------------------------------------------------------

/** Fixture 1: single comment, fully inside selection. */
const FIXTURE_1_FULLY_INSIDE = wrapPackage(
    `<w:p>` +
        `<w:r><w:t xml:space="preserve">Selection text </w:t></w:r>` +
        `<w:commentRangeStart w:id="0"/>` +
        `<w:r><w:t>anchor text</w:t></w:r>` +
        `<w:commentRangeEnd w:id="0"/>` +
        `<w:r><w:t xml:space="preserve"> more text</w:t></w:r>` +
    `</w:p>`
);

/** Fixture 2: head-outside (only end marker present). */
const FIXTURE_2_HEAD_OUTSIDE = wrapPackage(
    `<w:p>` +
        `<w:r><w:t>tail of comment anchor</w:t></w:r>` +
        `<w:commentRangeEnd w:id="0"/>` +
        `<w:r><w:t xml:space="preserve"> rest of selection</w:t></w:r>` +
    `</w:p>`
);

/** Fixture 3: tail-outside (only start marker present). */
const FIXTURE_3_TAIL_OUTSIDE = wrapPackage(
    `<w:p>` +
        `<w:r><w:t xml:space="preserve">start of selection </w:t></w:r>` +
        `<w:commentRangeStart w:id="0"/>` +
        `<w:r><w:t>head of comment anchor</w:t></w:r>` +
    `</w:p>`
);

/** Fixture 4: both-outside (no markers; comment fully contains selection). */
const FIXTURE_4_BOTH_OUTSIDE = wrapPackage(
    `<w:p>` +
        `<w:r><w:t>this entire selection is inside a comment that started in a previous paragraph</w:t></w:r>` +
    `</w:p>`
);

/** Fixture 5: nested ranges (comment B inside comment A). */
const FIXTURE_5_NESTED = wrapPackage(
    `<w:p>` +
        `<w:commentRangeStart w:id="0"/>` +
        `<w:r><w:t xml:space="preserve">outer </w:t></w:r>` +
        `<w:commentRangeStart w:id="1"/>` +
        `<w:r><w:t>inner</w:t></w:r>` +
        `<w:commentRangeEnd w:id="1"/>` +
        `<w:r><w:t xml:space="preserve"> end</w:t></w:r>` +
        `<w:commentRangeEnd w:id="0"/>` +
    `</w:p>`
);

/** Fixture 6: comment markers inside <w:ins> (tracked-change interaction). */
const FIXTURE_6_TRACKED_CHANGES = wrapPackage(
    `<w:p>` +
        `<w:ins w:id="100" w:author="Alice">` +
            `<w:commentRangeStart w:id="0"/>` +
            `<w:r><w:t>inserted-and-commented text</w:t></w:r>` +
            `<w:commentRangeEnd w:id="0"/>` +
        `</w:ins>` +
    `</w:p>`
);

/** Fixture 7: no comments at all — splicer must passthrough plain selection text. */
const FIXTURE_7_NO_COMMENTS = wrapPackage(
    `<w:p>` +
        `<w:r><w:t>plain selected text without any comment markers</w:t></w:r>` +
    `</w:p>`
);

// ---------------------------------------------------------------------------
// Annotation-token enum (STYLE.md "No Magic Strings"). The module keeps its
// canonical strings private (only formatSelectionWithComments is exported),
// so the spec mirrors the expected vocabulary here.
// ---------------------------------------------------------------------------

const ANNOTATION_TOKENS = Object.freeze({
    OPEN_BRACKET: '[COMMENT',
    CLOSE_BRACKET: '[/COMMENT]',
    RESOLVED_MARK: '(resolved)',
    HEAD_TRUNC_MARK: 'anchor extends before selection',
    TAIL_TRUNC_MARK: 'anchor extends past selection',
    BOTH_TRUNC_MARK: 'comment fully contains selection',
});

// ---------------------------------------------------------------------------
// CommentThread fixture builder (matches Plan 02 extractor return shape).
// ---------------------------------------------------------------------------

/**
 * Build a CommentThread test fixture matching the shape extractCommentsOnRange
 * (Plan 02) will return.
 *
 * @param {object} opts
 * @param {string} [opts.id]
 * @param {string} [opts.content]
 * @param {string} [opts.author]
 * @param {string} [opts.date]
 * @param {boolean} [opts.resolved]
 * @param {string} [opts.locationRelation] - 'Inside' | 'OverlapsBefore' | 'OverlapsAfter' | 'Contains' | ...
 * @param {Array<object>} [opts.replies]
 * @returns {object} CommentThread
 */
function makeThread({
    id = 'c0',
    content = 'comment body',
    author = 'Reviewer',
    date = '2026-01-01T00:00:00Z',
    resolved = false,
    locationRelation = 'Inside',
    replies = [],
} = {}) {
    return { id, content, authorName: author, creationDate: date, resolved, locationRelation, replies };
}

// ---------------------------------------------------------------------------
// Tests — 8 splicer behaviors per VALIDATION.md R2 table
// ---------------------------------------------------------------------------

describe('formatSelectionWithComments (Phase 05.2-R2)', () => {
    it('fully inside: single-comment fully-inside renders annotation at correct offsets', () => {
        const thread = makeThread({
            id: 'c0',
            content: 'Reviewer note',
            author: 'Reviewer',
            date: '2026-01-15T00:00:00Z',
            locationRelation: 'Inside',
        });

        const result = formatSelectionWithComments(FIXTURE_1_FULLY_INSIDE, [thread]);

        expect(typeof result).toBe('string');
        expect(result).toContain(ANNOTATION_TOKENS.OPEN_BRACKET);
        expect(result).toContain(ANNOTATION_TOKENS.CLOSE_BRACKET);
        expect(result).toContain('Reviewer');
        expect(result).toContain('Reviewer note');
        // Underlying run text must survive.
        expect(result).toContain('Selection text');
        expect(result).toContain('anchor text');
        expect(result).toContain('more text');
    });

    it('head outside: start marker absent — renders truncation marker at start of annotation', () => {
        const thread = makeThread({
            id: 'c0',
            content: 'spans before selection',
            locationRelation: 'OverlapsBefore',
        });

        const result = formatSelectionWithComments(FIXTURE_2_HEAD_OUTSIDE, [thread]);

        expect(result).toContain(ANNOTATION_TOKENS.HEAD_TRUNC_MARK);
        expect(result).toContain('tail of comment anchor');
        expect(result).toContain('rest of selection');
    });

    it('tail outside: end marker absent — renders truncation marker at end of annotation', () => {
        const thread = makeThread({
            id: 'c0',
            content: 'spans past selection',
            locationRelation: 'OverlapsAfter',
        });

        const result = formatSelectionWithComments(FIXTURE_3_TAIL_OUTSIDE, [thread]);

        expect(result).toContain(ANNOTATION_TOKENS.TAIL_TRUNC_MARK);
        expect(result).toContain('start of selection');
        expect(result).toContain('head of comment anchor');
    });

    it('both outside: selection inside comment range — wraps entire selection in annotation envelope', () => {
        // commentRange.compareLocationWith(selectionRange) returns 'Contains'
        // when the comment range fully contains the selection (research §
        // RELATION_TO_CASE table; verified against learn.microsoft.com/word.locationrelation).
        // The comparator direction is comment-vs-selection, NOT selection-vs-comment,
        // so "selection inside comment range" means commentRange Contains selection.
        const thread = makeThread({
            id: 'c0',
            content: 'comment fully contains selection',
            locationRelation: 'Contains',
        });

        const result = formatSelectionWithComments(FIXTURE_4_BOTH_OUTSIDE, [thread]);

        expect(result).toContain(ANNOTATION_TOKENS.BOTH_TRUNC_MARK);
        expect(result).toContain('this entire selection is inside a comment');
    });

    it('no comments passthrough: returns plain text identical to selection', () => {
        const result = formatSelectionWithComments(FIXTURE_7_NO_COMMENTS, []);

        expect(typeof result).toBe('string');
        // No annotation tokens whatsoever.
        expect(result).not.toContain(ANNOTATION_TOKENS.OPEN_BRACKET);
        expect(result).not.toContain(ANNOTATION_TOKENS.CLOSE_BRACKET);
        // Plain text preserved.
        expect(result).toContain('plain selected text without any comment markers');
    });

    it('nested ranges: comment B inside comment A anchor renders both annotations correctly', () => {
        const outer = makeThread({
            id: 'c0',
            content: 'outer comment',
            author: 'Outer',
            date: '2026-01-01T00:00:00Z',
            locationRelation: 'Inside',
        });
        const inner = makeThread({
            id: 'c1',
            content: 'inner comment',
            author: 'Inner',
            date: '2026-01-02T00:00:00Z',
            locationRelation: 'Inside',
        });

        const result = formatSelectionWithComments(FIXTURE_5_NESTED, [outer, inner]);

        expect(result).toContain('outer comment');
        expect(result).toContain('inner comment');
        expect(result).toContain('Outer');
        expect(result).toContain('Inner');
        expect(result).toContain('outer');
        expect(result).toContain('inner');
        expect(result).toContain('end');
    });

    it('mixed status: resolved thread carries (resolved) marker; open thread does not', () => {
        // Plan 01 SUMMARY § Next Phase Readiness recommended tightening this
        // assertion away from a positional-distance heuristic (which couldn't
        // distinguish two co-anchored threads on a single w:id pair) toward
        // exact occurrence counting + structural co-location. We split into two
        // independent passes — one fixture, one thread per pass — which
        // mirrors real Word OOXML semantics (one Comment per w:id).
        const openThread = makeThread({
            id: 'c0',
            content: 'open thread',
            author: 'Alice',
            date: '2026-01-01T00:00:00Z',
            resolved: false,
            locationRelation: 'Inside',
        });
        const resolvedThread = makeThread({
            id: 'c0',
            content: 'resolved thread',
            author: 'Bob',
            date: '2026-01-02T00:00:00Z',
            resolved: true,
            locationRelation: 'Inside',
        });

        const openOnly = formatSelectionWithComments(FIXTURE_1_FULLY_INSIDE, [openThread]);
        const resolvedOnly = formatSelectionWithComments(FIXTURE_1_FULLY_INSIDE, [resolvedThread]);

        // Open thread output: zero RESOLVED_MARK occurrences.
        expect(openOnly).toContain('open thread');
        expect(openOnly).not.toContain(ANNOTATION_TOKENS.RESOLVED_MARK);

        // Resolved thread output: exactly one RESOLVED_MARK occurrence,
        // structurally co-located with the resolved thread's content (same
        // annotation envelope — appears between the opening '[COMMENT' and
        // the colon that ends the open token, before the body content).
        expect(resolvedOnly).toContain('resolved thread');
        const resolvedMarkCount = resolvedOnly.split(ANNOTATION_TOKENS.RESOLVED_MARK).length - 1;
        expect(resolvedMarkCount).toBe(1);
        const markIdx = resolvedOnly.indexOf(ANNOTATION_TOKENS.RESOLVED_MARK);
        const openIdx = resolvedOnly.indexOf(ANNOTATION_TOKENS.OPEN_BRACKET);
        const contentIdx = resolvedOnly.indexOf('resolved thread');
        expect(openIdx).toBeGreaterThanOrEqual(0);
        expect(markIdx).toBeGreaterThan(openIdx);
        expect(markIdx).toBeLessThan(contentIdx);
    });

    it('tracked changes: comment markers inside <w:ins> still anchor correctly', () => {
        const thread = makeThread({
            id: 'c0',
            content: 'comment on inserted text',
            author: 'Alice',
            locationRelation: 'Inside',
        });

        const result = formatSelectionWithComments(FIXTURE_6_TRACKED_CHANGES, [thread]);

        expect(result).toContain(ANNOTATION_TOKENS.OPEN_BRACKET);
        expect(result).toContain(ANNOTATION_TOKENS.CLOSE_BRACKET);
        // Underlying inserted-text content survives the splice.
        expect(result).toContain('inserted-and-commented text');
        expect(result).toContain('comment on inserted text');
    });
});

// Fixtures must be queryable from outside via the file (Plan 03 may import them
// for parallel coverage). They are also declared here for static-grep gating.
module.exports = {
    FIXTURE_1_FULLY_INSIDE,
    FIXTURE_2_HEAD_OUTSIDE,
    FIXTURE_3_TAIL_OUTSIDE,
    FIXTURE_4_BOTH_OUTSIDE,
    FIXTURE_5_NESTED,
    FIXTURE_6_TRACKED_CHANGES,
    FIXTURE_7_NO_COMMENTS,
    ANNOTATION_TOKENS,
};
