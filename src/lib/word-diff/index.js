/**
 * Word Diff Strategy Layer
 *
 * The facade for applying text diffs to Word ranges as tracked changes:
 *
 *   - token-map      word-level granularity for Latin text (vendored)
 *   - sentence-diff  sentence-level fallback / optional coarse mode (vendored)
 *   - block-replace  last-resort full-range replacement (vendored)
 *   - char-diff      character-level strategy for CJK text (project-original)
 *
 * The vendored strategies come from office-word-diff (Apache-2.0) — see
 * LICENSE and NOTICE in this directory.
 *
 * @module lib/word-diff
 */

import DiffMatchPatch from './diff-wordmode.js';

export { applyTokenMapStrategy } from './token-map.js';
export { applySentenceDiffStrategy, diff_sentenceMode } from './sentence-diff.js';
export { applyBlockReplaceStrategy } from './block-replace.js';
export { hasCjk, computeCharEdits, applyCharDiffStrategy, sliceSearchPieces } from './char-diff.js';

/**
 * Computes a word-level diff between two strings. No Office.js context
 * required.
 *
 * @param {string} text1 - Original text
 * @param {string} text2 - New text
 * @returns {Array<[number, string]>} diff-match-patch ops ([[-1|0|1, text], ...])
 */
export function computeDiff(text1, text2) {
    const dmp = new DiffMatchPatch();
    return dmp.diff_wordMode(text1, text2);
}
