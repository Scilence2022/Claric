/**
 * Diff View Module
 *
 * Renders an inline before/after text diff for the proposal card's change
 * list: deletions struck through, insertions highlighted. Display-only —
 * the source of truth for applying changes stays the word-diff layer; this
 * view just shows the user what the model changed.
 *
 * @module ui/diff-view
 */

import DiffMatchPatch from '../../lib/vendor/diff-match-patch.js';

const dmp = new DiffMatchPatch();

/**
 * Builds a DOM element showing the word/character-level differences between
 * two texts (<del> for removals, <ins> for additions, plain spans for
 * unchanged runs).
 *
 * @param {string} before - Original text
 * @param {string} after - Proposed text
 * @returns {HTMLElement}
 */
export function buildTextDiffElement(before, after) {
    const el = document.createElement('div');
    el.className = 'diff-view';
    const diffs = dmp.diff_main(String(before || ''), String(after || ''));
    dmp.diff_cleanupSemantic(diffs);
    for (const [op, text] of diffs) {
        if (!text) continue;
        const span = document.createElement(op === -1 ? 'del' : op === 1 ? 'ins' : 'span');
        span.textContent = text;
        el.appendChild(span);
    }
    return el;
}
