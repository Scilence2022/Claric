/**
 * Char-Level Diff Strategy (CJK-aware)
 *
 * office-word-diff's token map tokenizes with /(\w+|[^\w\s]+|\s+)/g, which
 * treats a run of CJK characters as ONE token — so a one-comma edit in a
 * Chinese sentence becomes a whole-sentence replacement redline. This module
 * applies character-level diffs instead: it walks diff-match-patch ops
 * left-to-right over the target range and applies each minimal insert/delete
 * with tracked changes, so "加了一个逗号" shows up as exactly that.
 *
 * Used by the reassembler for paragraphs that contain CJK text; Latin text
 * keeps the word-level token map (nicer redline granularity for English).
 *
 * Locating the op spans in the document happens in ONE batched search pass
 * when possible (a single context.sync for the whole range — this dominates
 * full-document apply time, where every paragraph pays a round-trip per op
 * otherwise). The batch maps each span to its match by occurrence index,
 * mirroring Word's greedy left-to-right Find semantics; the rare position
 * the greedy simulation cannot resolve (self-overlapping repeated text)
 * falls back to the original sequential cursor walk.
 *
 * @module char-diff
 */

import DiffMatchPatch from '../vendor/diff-match-patch.js';

/** Matches CJK ideographs, hiragana/katakana, and hangul. */
const CJK_RE = /[\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff\u3040-\u30ff\uac00-\ud7af]/;

/**
 * True when the text contains CJK characters (no usable word boundaries).
 *
 * @param {string} text
 * @returns {boolean}
 */
export function hasCjk(text) {
    return CJK_RE.test(text || '');
}

/**
 * Computes minimal char-level edit ops between two texts.
 * Pure function — exported for tests.
 *
 * @param {string} originalText
 * @param {string} newText
 * @returns {Array<[number, string]>} diff-match-patch ops ([[-1|0|1, text], ...])
 */
export function computeCharEdits(originalText, newText) {
    const dmp = new DiffMatchPatch();
    const diffs = dmp.diff_main(originalText || '', newText || '');
    dmp.diff_cleanupMerge(diffs);
    return /** @type {Array<[number, string]>} */ (diffs);
}

/** Safety cap: more ops than this means the texts diverged too much for
 *  reliable cursor tracking — callers should fall back to a coarser strategy. */
const MAX_OPS = 200;

/**
 * Computes which of Word's greedy, left-to-right, non-overlapping matches of
 * `token` corresponds to the occurrence at `position` in `text`, or null when
 * the greedy scan cannot land on that position (an earlier overlapping match
 * consumed it, e.g. locating "aa" at offset 1 of "aaaa"). Word's Find scans
 * from the scope start and resumes after each match's end, so the Nth match
 * a search returns is exactly the Nth token of this simulation — the index
 * used to map batched search results back to spans (the same trick the
 * token-map strategy uses for repeated tokens).
 *
 * Pure function — exported for tests.
 *
 * @param {string} text - The searched scope's pristine text
 * @param {number} position - Start offset of the span being located
 * @param {string} token - The span's text
 * @returns {number | null} Zero-based match index, or null when unresolvable
 */
export function _occurrenceIndex(text, position, token) {
    let count = 0;
    let pos = 0;
    while (pos <= position) {
        const idx = text.indexOf(token, pos);
        if (idx === -1 || idx > position) return null;
        if (idx === position) return count;
        count += 1;
        pos = idx + token.length;
    }
    return null;
}

/**
 * Builds the edit hunks and the batched-locate targets from the char ops,
 * without touching the document model. Locate targets are:
 *  - the tail (≤ SEARCH_PIECE_MAX chars) of every EQUAL run that anchors a
 *    pure insertion — only its end position matters for the anchor;
 *  - every DELETE span: whole when short, else its first + last search
 *    pieces (the spanning union covers the middle).
 *
 * Returns null when any target's occurrence cannot be resolved by the greedy
 * simulation; the caller then locates the whole range with the sequential
 * cursor walk, which matches Word's per-search semantics exactly.
 *
 * @private
 * @param {string} originalText
 * @param {Array<[number, string]>} ops
 * @returns {{hunks: Array<object>, targets: Array<object>} | null}
 */
function _planFromOps(originalText, ops) {
    /** Edit hunks in document order. */
    const hunks = [];
    /** Flat target list: { token, position, kind, occurrence, match }. */
    const targets = [];

    let lastEqual = null;   // {start, end} of the EQUAL run before the pending hunk
    let pendingDel = null;  // {start, end} spanning the pending contiguous deletes
    let pendingIns = '';
    let pos = 0;

    const addDeleteTargets = (hunk) => {
        const { start, end } = hunk.span;
        const pieces = sliceSearchPieces(originalText.slice(start, end));
        if (pieces.length === 1) {
            const target = { token: pieces[0], position: start, kind: 'delete', match: null };
            targets.push(target);
            hunk.delTargets = [target];
        } else {
            const last = pieces[pieces.length - 1];
            const firstTarget = { token: pieces[0], position: start, kind: 'delete', match: null };
            const lastTarget = { token: last, position: end - last.length, kind: 'delete', match: null };
            targets.push(firstTarget, lastTarget);
            hunk.delTargets = [firstTarget, lastTarget];
        }
    };

    const flushHunk = () => {
        if (!pendingDel && !pendingIns) return;
        if (!pendingDel) {
            // Pure insertion anchored after the last EQUAL run (or at the
            // range start when the text begins with an insertion). Only the
            // run's tail is located: the anchor position is the run's END,
            // and the tail match ends exactly there.
            const hunk = { type: 'insert', text: pendingIns, tail: null };
            hunks.push(hunk);
            if (lastEqual) {
                const pieces = sliceSearchPieces(originalText.slice(lastEqual.start, lastEqual.end));
                const tail = pieces[pieces.length - 1];
                const target = {
                    token: tail,
                    position: lastEqual.end - tail.length,
                    kind: 'equal',
                    match: null,
                };
                targets.push(target);
                hunk.tail = target;
            }
        } else {
            const hunk = {
                type: pendingIns ? 'replace' : 'delete',
                text: pendingIns,
                delTargets: [],
                span: pendingDel,
            };
            hunks.push(hunk);
            addDeleteTargets(hunk);
        }
        pendingDel = null;
        pendingIns = '';
    };

    for (const [op, text] of ops) {
        if (op === 0) {
            flushHunk();
            lastEqual = { start: pos, end: pos + text.length };
            pos += text.length;
        } else if (op === -1) {
            // Adjacent DELETE ops are contiguous (INSERTs consume no original
            // text), so one spanning {start, end} accumulates them all —
            // the same union the cursor walk builds from per-op matches.
            if (pendingDel) {
                pendingDel.end = pos + text.length;
            } else {
                pendingDel = { start: pos, end: pos + text.length };
            }
            pos += text.length;
        } else {
            pendingIns += text;
        }
    }
    flushHunk();

    // Resolve every target's occurrence index up front; a single
    // unresolvable position escalates the whole range to the cursor walk.
    for (const target of targets) {
        target.occurrence = _occurrenceIndex(originalText, target.position, target.token);
        if (target.occurrence === null) return null;
    }
    return { hunks, targets };
}

/**
 * Locates every target with ONE batched search pass (all searches queued, a
 * single sync — the whole point of the batched path), then assembles the edit
 * plan. Throws when a token cannot be matched (document drift), so callers
 * fall back to a coarser strategy.
 *
 * @private
 * @param {Word.RequestContext} context
 * @param {Word.Range} range
 * @param {{hunks: Array<object>, targets: Array<object>}} planned
 * @returns {Promise<Array<object>>} Edit plan
 */
async function _locatePlanBatched(context, range, planned) {
    const { hunks, targets } = planned;

    // One search per DISTINCT token; targets sharing a token differ only in
    // their occurrence index.
    const searchByToken = new Map();
    for (const target of targets) {
        if (!searchByToken.has(target.token)) {
            const matches = range.search(target.token, { matchCase: true, matchWholeWord: false });
            matches.load('items');
            searchByToken.set(target.token, matches);
        }
    }
    // SYNC: the single locate round-trip for the whole range.
    if (searchByToken.size > 0) {
        await context.sync();
    }

    for (const target of targets) {
        const items = searchByToken.get(target.token).items;
        if (!items || items.length <= target.occurrence) {
            throw new Error(`char-diff: ${target.kind} text not found in document: "${_preview(target.token)}"`);
        }
        target.match = items[target.occurrence];
    }

    return hunks.map((hunk) => {
        if (hunk.type === 'insert') {
            return hunk.tail
                ? { type: 'insert', anchor: hunk.tail.match, location: Word.InsertLocation.after, text: hunk.text }
                : { type: 'insert', anchor: range.getRange(Word.RangeLocation.start), location: Word.InsertLocation.before, text: hunk.text };
        }
        // Whole-span piece, or first + last pieces whose spanning union
        // covers the entire delete span (including the unsearched middle).
        const [first, last] = hunk.delTargets;
        const delRange = last ? first.match.expandTo(last.match) : first.match;
        return hunk.type === 'delete'
            ? { type: 'delete', range: delRange }
            : { type: 'replace', range: delRange, text: hunk.text };
    });
}

/**
 * Locates every op's span by walking a tail cursor with one search per piece
 * (one sync each). The original sequential path, kept as the fallback for
 * positions the batched occurrence mapping cannot resolve.
 *
 * Cursor mechanics: the walking cursor is the tail of the scope spanning
 * from the end of the last located match to the scope end, built with
 * getRange(End).expandTo(scopeEnd). Do NOT use
 * match.getRange(Word.RangeLocation.after): on search()-produced ranges
 * it yields a zero-width point whose .text is '', so every locate after
 * the first one fails.
 *
 * @private
 */
async function _buildPlanByCursorWalk(context, range, ops) {
    const scopeEnd = range.getRange(Word.RangeLocation.end);
    const plan = [];
    let cursor = range;
    let lastEqualMatch = null;
    let pendingDels = [];
    let pendingIns = '';

    const flushHunk = () => {
        if (pendingDels.length === 0 && !pendingIns) return;
        if (pendingDels.length === 0) {
            // Pure insertion: anchor after the last EQUAL match (or at the
            // range start when the text begins with an insertion).
            plan.push(lastEqualMatch
                ? { type: 'insert', anchor: lastEqualMatch, location: Word.InsertLocation.after, text: pendingIns }
                : { type: 'insert', anchor: range.getRange(Word.RangeLocation.start), location: Word.InsertLocation.before, text: pendingIns });
        } else if (!pendingIns) {
            plan.push({ type: 'delete', range: _unionRanges(pendingDels) });
        } else {
            plan.push({ type: 'replace', range: _unionRanges(pendingDels), text: pendingIns });
        }
        pendingDels = [];
        pendingIns = '';
    };

    for (const [op, text] of ops) {
        if (op === 0) {
            flushHunk();
            const match = await _locateAt(context, cursor, scopeEnd, text, 'equal');
            cursor = match.getRange(Word.RangeLocation.end).expandTo(scopeEnd);
            lastEqualMatch = match;
        } else if (op === -1) {
            const match = await _locateAt(context, cursor, scopeEnd, text, 'delete');
            pendingDels.push(match);
            cursor = match.getRange(Word.RangeLocation.end).expandTo(scopeEnd);
        } else {
            pendingIns += text;
        }
    }
    flushHunk();
    return plan;
}

/**
 * Applies a char-level diff to a range as tracked changes.
 *
 * Two phases, mirroring office-word-diff's token map architecture:
 *   Phase 1 (read-only): locate every op's span in the PRISTINE document
 *     text. Preferred path is one batched search pass (single sync) with
 *     occurrence-indexed mapping; the sequential cursor walk (one sync per
 *     piece) remains as the fallback for positions the greedy simulation
 *     cannot resolve. Interleaved delete+insert runs are merged into single
 *     "replace" ops over the contiguous deleted span — this keeps every plan
 *     entry anchored to one located range, which stays valid regardless of
 *     edit order.
 *   Phase 2 (edit): execute the plan in REVERSE document order with tracked
 *     changes on, so each minimal insert/delete/replace becomes its own
 *     revision.
 *
 * When the expected text can no longer be found in the scope, this throws so
 * callers can fall back to another strategy.
 *
 * @param {Word.RequestContext} context
 * @param {Word.Range} range - Target range (typically one paragraph)
 * @param {string} originalText - Expected current text of the range
 * @param {string} newText - Desired text
 * @param {function} [log] - Logging callback
 * @param {object} [options]
 * @param {boolean} [options.trackChanges=true] - When false the strategy does
 *   NOT touch the document's changeTrackingMode; the caller owns it.
 * @returns {Promise<{strategy: string, insertions: number, deletions: number, replacements: number}>}
 * @throws {Error} On cursor/document divergence or excessive op count
 */
export async function applyCharDiffStrategy(context, range, originalText, newText, log = () => {}, options = {}) {
    const ops = computeCharEdits(originalText, newText).filter(([, text]) => text.length > 0);
    if (ops.length > MAX_OPS) {
        throw new Error(`char-diff: ${ops.length} ops exceeds safety cap (${MAX_OPS})`);
    }

    // Phase 1: locate all spans against the pristine text (no edits yet).
    const planned = _planFromOps(originalText, ops);
    const plan = planned
        ? await _locatePlanBatched(context, range, planned)
        : await _buildPlanByCursorWalk(context, range, ops);

    // Phase 2: execute in reverse document order. When this strategy owns
    // tracking, force trackAll and ALWAYS restore off afterwards — even when
    // an edit throws mid-plan, so a failure never leaks tracking state.
    const trackChanges = options.trackChanges !== false;
    if (trackChanges && Word.ChangeTrackingMode) {
        context.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
    }

    let insertions = 0;
    let deletions = 0;
    let replacements = 0;
    try {
        for (const item of plan.slice().reverse()) {
            if (item.type === 'delete') {
                item.range.delete();
                deletions++;
            } else if (item.type === 'replace') {
                item.range.insertText(item.text, Word.InsertLocation.replace);
                replacements++;
            } else {
                item.anchor.insertText(item.text, item.location);
                insertions++;
            }
        }
        await context.sync();
    } finally {
        if (trackChanges && Word.ChangeTrackingMode) {
            context.document.changeTrackingMode = Word.ChangeTrackingMode.off;
            await context.sync();
        }
    }

    log(`Char-level diff applied (${insertions} insertions, ${deletions} deletions, ${replacements} replacements)`, 'info');
    return { strategy: 'char', insertions, deletions, replacements };
}

/**
 * Merges contiguous located ranges into one spanning range.
 * @private
 */
function _unionRanges(ranges) {
    let union = ranges[0];
    for (let i = 1; i < ranges.length; i++) {
        union = union.expandTo(ranges[i]);
    }
    return union;
}

/** Word's Find/Replace search string is limited to 255 characters, so runs
 *  longer than this are located piece by piece and unioned into one range. */
const SEARCH_PIECE_MAX = 200;

/**
 * Splits text into ≤ SEARCH_PIECE_MAX-char search pieces. A piece boundary
 * never falls between a surrogate pair: a lone surrogate half in a search
 * string cannot match and would break the locate step for emoji / astral
 * characters sitting exactly on the boundary.
 *
 * @param {string} text
 * @returns {string[]}
 */
export function sliceSearchPieces(text) {
    const pieces = [];
    let offset = 0;
    while (offset < text.length) {
        let end = Math.min(offset + SEARCH_PIECE_MAX, text.length);
        if (end < text.length) {
            const code = text.charCodeAt(end - 1);
            if (code >= 0xd800 && code <= 0xdbff) end -= 1;
        }
        pieces.push(text.slice(offset, end));
        offset = end;
    }
    return pieces;
}

/**
 * Locates expectedText starting at the cursor and returns the match range.
 *
 * The located range is walked in ≤ SEARCH_PIECE_MAX-char pieces because Word
 * rejects search strings longer than 255 chars. Each piece search starts from
 * the tail after the previous piece (built via getRange(End).expandTo), so
 * the first match is always the ordered next one; a genuine miss throws so
 * callers can fall back to a coarser strategy.
 * @private
 */
async function _locateAt(context, cursor, scopeEnd, expectedText, kind) {
    let union = null;
    let pieceCursor = cursor;
    for (const piece of sliceSearchPieces(expectedText)) {
        const matches = pieceCursor.search(piece, { matchCase: true, matchWholeWord: false });
        matches.load('items');
        await context.sync();
        if (!matches.items.length) {
            throw new Error(`char-diff: ${kind} text not found in document: "${_preview(piece)}"`);
        }
        const match = matches.items[0];
        union = union ? union.expandTo(match) : match;
        pieceCursor = match.getRange(Word.RangeLocation.end).expandTo(scopeEnd);
    }
    return union;
}

/** @private */
function _preview(text) {
    const t = (text || '').slice(0, 20);
    return text && text.length > 20 ? `${t}…` : t;
}
