/**
 * Selection-with-Comments Module
 *
 * Pure-function splicer: takes pkg:package-wrapped OOXML from
 * selection.getOoxml() plus a pre-filtered/sorted CommentThread array (from
 * extractCommentsOnRange) and returns selection text with inline [COMMENT...]
 * annotations interleaved at w:commentRangeStart / w:commentRangeEnd marker
 * positions. No Office.js. No network. Hermetic-testable.
 *
 * @module selection-with-comments
 */

import { W_NS, extractDocumentBody, readRunText } from './comment-extractor.js';

/**
 * Marker-case enum covering the 4 ways a comment range can intersect a selection.
 * Frozen — STYLE.md "Enums for Fixed Values".
 */
const MARKER_CASE = Object.freeze({
    FULLY_INSIDE: 'fully_inside',
    HEAD_OUTSIDE: 'head_outside',
    TAIL_OUTSIDE: 'tail_outside',
    BOTH_OUTSIDE: 'both_outside',
});

/**
 * Annotation-token vocabulary. Frozen — STYLE.md "No Magic Strings".
 * Mirrors tests/selection-with-comments.spec.js ANNOTATION_TOKENS.
 */
const ANNOTATION_TOKENS = Object.freeze({
    OPEN_PREFIX: '[COMMENT',
    CLOSE_TOKEN: '[/COMMENT]',
    RESOLVED_MARK: '(resolved)',
    HEAD_TRUNC: 'anchor extends before selection',
    TAIL_TRUNC: 'anchor extends past selection',
    BOTH_TRUNC: 'comment fully contains selection',
    REPLY_PREFIX: 'reply ',
});

/**
 * Map Word.LocationRelation → MARKER_CASE.
 * Frozen — single source of truth. Verified against learn.microsoft.com/word.locationrelation.
 */
const RELATION_TO_CASE = Object.freeze({
    Equal: MARKER_CASE.FULLY_INSIDE,
    Inside: MARKER_CASE.FULLY_INSIDE,
    InsideStart: MARKER_CASE.FULLY_INSIDE,
    InsideEnd: MARKER_CASE.FULLY_INSIDE,
    OverlapsBefore: MARKER_CASE.HEAD_OUTSIDE,
    OverlapsAfter: MARKER_CASE.TAIL_OUTSIDE,
    Contains: MARKER_CASE.BOTH_OUTSIDE,
    ContainsStart: MARKER_CASE.TAIL_OUTSIDE,
    ContainsEnd: MARKER_CASE.HEAD_OUTSIDE,
});

/**
 * ISO-8601 (YYYY-MM-DD). 'unknown' on falsy; original string when unparseable.
 * @param {Date|string|null|undefined} d
 * @returns {string}
 */
function formatDate(d) {
    if (!d) return 'unknown';
    const dt = (d instanceof Date) ? d : new Date(d);
    if (isNaN(dt.getTime())) return String(d);
    return dt.toISOString().slice(0, 10);
}

/**
 * Build opening token: `[COMMENT[ (resolved)] — Author (YYYY-MM-DD)[ ‹trunc›]: body]`.
 * @param {{authorName:string,creationDate:(Date|string),resolved:boolean,content:string}} thread
 * @param {string|null} truncMark - truncation phrase; null for FULLY_INSIDE
 * @returns {string}
 */
function buildOpenToken(thread, truncMark) {
    const status = thread.resolved ? ` ${ANNOTATION_TOKENS.RESOLVED_MARK}` : '';
    const date = formatDate(thread.creationDate);
    const trunc = truncMark ? ` ‹${truncMark}›` : '';
    return `${ANNOTATION_TOKENS.OPEN_PREFIX}${status} — ${thread.authorName} (${date})${trunc}: ${thread.content}]`;
}

/**
 * Render replies as newline-prefixed lines. Extractor pre-sorts by creationDate.
 * @param {Array<{authorName:string,creationDate:(Date|string),content:string}>|undefined} replies
 * @returns {string} Empty string when no replies; else leading-newline + joined lines.
 */
function renderReplies(replies) {
    if (!replies || replies.length === 0) return '';
    const lines = replies.map((r) => (
        `${ANNOTATION_TOKENS.REPLY_PREFIX}${r.authorName} (${formatDate(r.creationDate)}): ${r.content}`
    ));
    return '\n' + lines.join('\n');
}

/**
 * Read the w:id attribute off a commentRange marker element (namespace-aware fallback).
 * @param {Element} el
 * @returns {string|null}
 */
function readMarkerWId(el) {
    return el.getAttributeNS(W_NS, 'id')
        || el.getAttribute('w:id')
        || el.getAttribute('id');
}

/**
 * Walk OOXML body in document order; emit flat segment array of
 * {kind:'start'|'end'|'text', value?, wId?}. Recurses into containers (w:ins,
 * w:del) so commentRange markers under tracked-change envelopes still emit
 * (research Pitfall 5).
 * @param {Element} bodyEl - w:body element from extractDocumentBody
 * @returns {Array<{kind:string,value?:string,wId?:string}>}
 */
function walkOoxmlSegments(bodyEl) {
    const segments = [];
    const walk = (node) => {
        for (const child of Array.from(node.childNodes || [])) {
            if (child.nodeType !== 1) continue;
            const name = child.localName;
            if (name === 'commentRangeStart') {
                segments.push({ kind: 'start', wId: readMarkerWId(child) });
            } else if (name === 'commentRangeEnd') {
                segments.push({ kind: 'end', wId: readMarkerWId(child) });
            } else if (name === 'r') {
                // w:t (visible) + w:delText (tracked deletion) — both visible from splicer POV.
                const visible = readRunText(child, /*useDelText=*/false);
                if (visible) segments.push({ kind: 'text', value: visible });
                const deleted = readRunText(child, /*useDelText=*/true);
                if (deleted) segments.push({ kind: 'text', value: deleted });
            } else {
                walk(child);
            }
        }
    };
    walk(bodyEl);
    return segments;
}

/**
 * Render-dispatch table keyed by MARKER_CASE → {open, close} token pair.
 * STYLE.md "Dispatch Over If/Else" — frozen lookup, not a conditional chain.
 */
const RENDER_DISPATCH = Object.freeze({
    [MARKER_CASE.FULLY_INSIDE]: (thread) => ({
        open: buildOpenToken(thread, /*truncMark=*/null),
        close: ANNOTATION_TOKENS.CLOSE_TOKEN,
    }),
    [MARKER_CASE.HEAD_OUTSIDE]: (thread) => ({
        open: buildOpenToken(thread, ANNOTATION_TOKENS.HEAD_TRUNC),
        close: ANNOTATION_TOKENS.CLOSE_TOKEN,
    }),
    [MARKER_CASE.TAIL_OUTSIDE]: (thread) => ({
        open: buildOpenToken(thread, /*truncMark=*/null),
        close: `[/COMMENT ‹${ANNOTATION_TOKENS.TAIL_TRUNC}›]`,
    }),
    [MARKER_CASE.BOTH_OUTSIDE]: (thread) => ({
        open: buildOpenToken(thread, ANNOTATION_TOKENS.BOTH_TRUNC),
        close: `[/COMMENT ‹${ANNOTATION_TOKENS.BOTH_TRUNC}›]`,
    }),
});

/**
 * Splice inline [COMMENT...] annotations into selection OOXML.
 *
 * Walks the body for w:commentRangeStart / w:commentRangeEnd markers, pairs
 * each w:id (first-appearance order) with a positional CommentThread
 * (FULLY_INSIDE / HEAD_OUTSIDE / TAIL_OUTSIDE), and emits open/close tokens at
 * marker positions. BOTH_OUTSIDE threads carry no markers in OOXML (range
 * starts before AND ends after selection), so the splicer wraps the entire
 * assembled selection. Replies render after the closing token (one per line,
 * prefixed "reply "). Pure: no Word.run, no I/O. O(n) via segment array + join.
 *
 * @param {string} ooxml - pkg:package-wrapped OOXML from selection.getOoxml()
 * @param {Array<{id:string,content:string,authorName:string,creationDate:(Date|string),resolved:boolean,locationRelation:string,replies?:Array<object>}>} commentsInDocOrder
 * @returns {string}
 */
export function formatSelectionWithComments(ooxml, commentsInDocOrder) {
    const parser = new DOMParser();
    const doc = parser.parseFromString(ooxml, 'application/xml');
    if (doc.getElementsByTagName('parsererror').length > 0) {
        const err = new Error('formatSelectionWithComments: OOXML parse error');
        console.error('[selection-with-comments]', err.message, {
            ooxmlPrefix: String(ooxml).slice(0, 200),
        });
        throw err;
    }

    const body = extractDocumentBody(doc);
    const segments = walkOoxmlSegments(body);
    const plainText = segments
        .filter((s) => s.kind === 'text')
        .map((s) => s.value)
        .join('');

    // Partition: BOTH_OUTSIDE wraps entire assembled output (no markers);
    // others pair positionally by w:id-first-appearance order.
    const bothOutsideThreads = [];
    const positionalThreads = [];
    for (const t of (commentsInDocOrder || [])) {
        const markerCase = RELATION_TO_CASE[t.locationRelation];
        if (!markerCase) {
            console.warn('[selection-with-comments] unknown locationRelation, skipping thread', {
                id: t.id, locationRelation: t.locationRelation,
            });
            continue;
        }
        if (markerCase === MARKER_CASE.BOTH_OUTSIDE) {
            bothOutsideThreads.push(t);
        } else {
            positionalThreads.push({ thread: t, markerCase });
        }
    }

    if (positionalThreads.length === 0 && bothOutsideThreads.length === 0) {
        return plainText;
    }

    // First-appearance order of w:ids in the segment stream is the pairing key.
    const wIdOrder = [];
    const seenWId = new Set();
    for (const seg of segments) {
        if ((seg.kind === 'start' || seg.kind === 'end') && seg.wId && !seenWId.has(seg.wId)) {
            seenWId.add(seg.wId);
            wIdOrder.push(seg.wId);
        }
    }
    if (wIdOrder.length !== positionalThreads.length) {
        console.warn('[selection-with-comments] w:id count vs positional thread count mismatch — truncating to minimum', {
            wIdCount: wIdOrder.length, threadCount: positionalThreads.length,
        });
    }
    const pairCount = Math.min(wIdOrder.length, positionalThreads.length);
    const openTokenForId = new Map();
    const closeTokenForId = new Map();
    const repliesForId = new Map();
    for (let i = 0; i < pairCount; i++) {
        const wId = wIdOrder[i];
        const { thread, markerCase } = positionalThreads[i];
        const tokens = RENDER_DISPATCH[markerCase](thread);
        openTokenForId.set(wId, tokens.open);
        closeTokenForId.set(wId, tokens.close);
        repliesForId.set(wId, renderReplies(thread.replies));
    }

    // Emission pass — segment array + .join (O(n), never += in a loop).
    const out = [];
    const openEmitted = new Set();
    const closeEmitted = new Set();

    // HEAD_OUTSIDE pre-pass: open at offset 0 (no commentRangeStart in selection).
    for (let i = 0; i < pairCount; i++) {
        const wId = wIdOrder[i];
        const { markerCase } = positionalThreads[i];
        if (markerCase === MARKER_CASE.HEAD_OUTSIDE && !openEmitted.has(wId)) {
            out.push(openTokenForId.get(wId));
            openEmitted.add(wId);
        }
    }

    for (const seg of segments) {
        if (seg.kind === 'text') {
            out.push(seg.value);
        } else if (seg.kind === 'start') {
            if (!openEmitted.has(seg.wId) && openTokenForId.has(seg.wId)) {
                out.push(openTokenForId.get(seg.wId));
                openEmitted.add(seg.wId);
            }
        } else if (seg.kind === 'end') {
            if (!closeEmitted.has(seg.wId) && closeTokenForId.has(seg.wId)) {
                out.push(closeTokenForId.get(seg.wId));
                const r = repliesForId.get(seg.wId);
                if (r) out.push(r);
                closeEmitted.add(seg.wId);
            }
        }
    }

    // TAIL_OUTSIDE post-pass: close at end (no commentRangeEnd in selection).
    for (let i = 0; i < pairCount; i++) {
        const wId = wIdOrder[i];
        const { markerCase } = positionalThreads[i];
        if (markerCase === MARKER_CASE.TAIL_OUTSIDE && !closeEmitted.has(wId)) {
            out.push(closeTokenForId.get(wId));
            const r = repliesForId.get(wId);
            if (r) out.push(r);
            closeEmitted.add(wId);
        }
    }

    let assembled = out.join('');

    // BOTH_OUTSIDE: each thread wraps the assembled selection (nested in input order).
    for (const t of bothOutsideThreads) {
        const tokens = RENDER_DISPATCH[MARKER_CASE.BOTH_OUTSIDE](t);
        const replies = renderReplies(t.replies);
        assembled = `${tokens.open}${assembled}${tokens.close}${replies}`;
    }

    return assembled;
}
