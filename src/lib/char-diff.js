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
 * @module char-diff
 */

import DiffMatchPatch from 'diff-match-patch';

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
    return diffs;
}

/** Safety cap: more ops than this means the texts diverged too much for
 *  reliable cursor tracking — callers should fall back to a coarser strategy. */
const MAX_OPS = 200;

/**
 * Applies a char-level diff to a range as tracked changes.
 *
 * Two phases, mirroring office-word-diff's token map architecture:
 *   Phase 1 (read-only): locate every op's span in the PRISTINE document
 *     text, verifying the diff still matches the live content. Interleaved
 *     delete+insert runs are merged into single "replace" ops over the
 *     contiguous deleted span — this keeps every plan entry anchored to one
 *     located range (no zero-width cursors between ops), which stays valid
 *     regardless of edit order.
 *   Phase 2 (edit): execute the plan in REVERSE document order with tracked
 *     changes on, so each minimal insert/delete/replace becomes its own
 *     revision.
 *
 * Any divergence between the diff and the live document text throws, so
 * callers can fall back to another strategy.
 *
 * @param {Word.RequestContext} context
 * @param {Word.Range} range - Target range (typically one paragraph)
 * @param {string} originalText - Expected current text of the range
 * @param {string} newText - Desired text
 * @param {function} [log] - Logging callback
 * @returns {Promise<{strategy: string, insertions: number, deletions: number, replacements: number}>}
 * @throws {Error} On cursor/document divergence or excessive op count
 */
export async function applyCharDiffStrategy(context, range, originalText, newText, log = () => {}) {
    const ops = computeCharEdits(originalText, newText).filter(([, text]) => text.length > 0);
    if (ops.length > MAX_OPS) {
        throw new Error(`char-diff: ${ops.length} ops exceeds safety cap (${MAX_OPS})`);
    }

    // Phase 1: locate all spans against the pristine text (no edits yet).
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
            const match = await _locateAt(context, cursor, text, 'equal');
            cursor = match.getRange(Word.RangeLocation.after);
            lastEqualMatch = match;
        } else if (op === -1) {
            const match = await _locateAt(context, cursor, text, 'delete');
            pendingDels.push(match);
            cursor = match.getRange(Word.RangeLocation.after);
        } else {
            pendingIns += text;
        }
    }
    flushHunk();

    // Phase 2: execute in reverse document order as tracked changes.
    if (Word.ChangeTrackingMode) {
        context.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
    }

    let insertions = 0;
    let deletions = 0;
    let replacements = 0;
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

    if (Word.ChangeTrackingMode) {
        context.document.changeTrackingMode = Word.ChangeTrackingMode.off;
        await context.sync();
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
 * Verifies the cursor's remaining text starts with expectedText and returns
 * the located match range. Throws on divergence.
 *
 * The startsWith check covers the whole expectedText up front (pure JS, no
 * length limit); the located range is then walked in ≤ SEARCH_PIECE_MAX-char
 * pieces because Word rejects search strings longer than 255 chars. Each
 * piece is guaranteed to sit exactly at the piece cursor, so the first match
 * is always the right one.
 * @private
 */
async function _locateAt(context, cursor, expectedText, kind) {
    cursor.load('text');
    await context.sync();
    if (!cursor.text.startsWith(expectedText)) {
        throw new Error(
            `char-diff: expected ${kind} "${_preview(expectedText)}" at cursor, found "${_preview(cursor.text)}"`
        );
    }
    let union = null;
    let pieceCursor = cursor;
    for (let offset = 0; offset < expectedText.length; offset += SEARCH_PIECE_MAX) {
        const piece = expectedText.slice(offset, offset + SEARCH_PIECE_MAX);
        const matches = pieceCursor.search(piece, { matchCase: true, matchWholeWord: false });
        matches.load('items');
        await context.sync();
        if (!matches.items.length) {
            throw new Error(`char-diff: ${kind} text not found in document: "${_preview(piece)}"`);
        }
        const match = matches.items[0];
        union = union ? union.expandTo(match) : match;
        pieceCursor = match.getRange(Word.RangeLocation.after);
    }
    return union;
}

/** @private */
function _preview(text) {
    const t = (text || '').slice(0, 20);
    return text && text.length > 20 ? `${t}…` : t;
}
