/**
 * Paragraph Alignment Module
 *
 * Pairs an original paragraph list with an amended one. Shared by the
 * reassembler (which turns the pairing into tracked edits) and the amendment
 * guard (which uses it to measure how much of a chunk a model reply deletes
 * or rewrites before anything is staged).
 *
 * Pure JavaScript -- no Word API dependency.
 *
 * @module paragraph-alignment
 */

import { hasCjk } from './word-diff/char-diff.js';
import { normalizeWordText } from './word-revisions.js';

/** Minimum similarity for two paragraphs to count as the same, edited, paragraph. */
export const SIMILARITY_THRESHOLD = 0.4;

/**
 * A newline in the text protocol can be a native manual break, while blank
 * output lines are often model presentation. Preserve the known native
 * grouping instead of manufacturing paragraphs from either case.
 */
export function amendmentParagraphTexts(original, amendment) {
  const lines = normalizeWordText(amendment).split('\n');
  while (lines.length && !lines[0].trim()) lines.shift();
  while (lines.length && !lines.at(-1).trim()) lines.pop();
  const sourceLines = original.map((text) => normalizeWordText(text).split('\n'));
  const counts = sourceLines.map((lines) => lines.length);
  if (counts.every((count) => count === 1)) return lines.filter((line) => line.trim());
  const expected = counts.reduce((sum, count) => sum + count, 0);
  const nonEmpty = lines.filter((line) => line.trim());
  const flatSource = sourceLines.flat();
  if (nonEmpty.length !== flatSource.filter((line) => line.trim()).length) {
    throw new Error(`Cannot map ${nonEmpty.length} output content lines to ${original.length} native paragraphs containing ${expected} lines. Preserve every original manual line break and paragraph boundary.`);
  }
  let contentOffset = 0;
  const contentLines = flatSource.map((line) => line.trim() ? nonEmpty[contentOffset++] : line);
  let offset = 0;
  return counts.map((count) => {
    const text = contentLines.slice(offset, offset + count).join('\n');
    offset += count;
    return text;
  });
}

/**
 * Computes word-level similarity between two strings as a ratio (0-1).
 * Uses the number of shared words divided by the max word count.
 *
 * CJK text has no whitespace word boundaries — a whole paragraph is a single
 * "word", so word overlap is always 0 for an edited CJK paragraph and the
 * alignment would treat every polished paragraph as delete+insert (a
 * whole-paragraph redline). CJK input uses a character-bigram Dice
 * coefficient instead.
 *
 * @param {string} a - First string
 * @param {string} b - Second string
 * @returns {number} Similarity ratio from 0.0 to 1.0
 */
export function paragraphSimilarity(a, b) {
  const ta = a.trim();
  const tb = b.trim();
  if (ta === tb) return 1.0;
  if (!ta || !tb) return 0.0;

  if (hasCjk(ta) || hasCjk(tb)) {
    return _bigramSimilarity(ta, tb);
  }

  const wordsA = ta.split(/\s+/);
  const wordsB = tb.split(/\s+/);
  const setA = new Set(wordsA);
  let shared = 0;
  for (const w of wordsB) {
    if (setA.has(w)) shared++;
  }
  return shared / Math.max(wordsA.length, wordsB.length);
}

/**
 * Character-bigram Dice coefficient with multiset semantics:
 * 2 * |shared bigrams| / (|bigrams(a)| + |bigrams(b)|).
 *
 * @param {string} a - First string (already trimmed)
 * @param {string} b - Second string (already trimmed)
 * @returns {number} Similarity ratio from 0.0 to 1.0
 * @private
 */
function _bigramSimilarity(a, b) {
  if (a.length < 2 || b.length < 2) return a === b ? 1.0 : 0.0;
  const bigrams = new Map();
  for (let i = 0; i < a.length - 1; i++) {
    const bg = a.slice(i, i + 2);
    bigrams.set(bg, (bigrams.get(bg) || 0) + 1);
  }
  let shared = 0;
  for (let i = 0; i < b.length - 1; i++) {
    const bg = b.slice(i, i + 2);
    const count = bigrams.get(bg) || 0;
    if (count > 0) {
      shared++;
      bigrams.set(bg, count - 1);
    }
  }
  return (2 * shared) / (a.length - 1 + b.length - 1);
}

/**
 * Computes a paragraph-level alignment between original and amended paragraph arrays.
 *
 * Uses a two-phase approach:
 * 1. LCS on exact (trimmed) text matches to anchor unchanged paragraphs
 * 2. Greedy forward matching between LCS gaps with similarity threshold (>= 0.4)
 *    to capture paragraphs where the LLM made edits within the paragraph
 *
 * Returns an array of operations: 'keep' (aligned pair, may have text changes),
 * 'delete' (original only), 'insert' (amended only).
 *
 * @param {string[]} origParas - Original paragraph texts
 * @param {string[]} newParas - Amended paragraph texts
 * @returns {Array<{type: 'keep'|'delete'|'insert', origIdx?: number, newIdx?: number}>}
 */
export function alignParagraphs(origParas, newParas) {
  const m = origParas.length;
  const n = newParas.length;

  // Phase 1: LCS on exact trimmed text to find anchors
  const dp = Array.from({ length: m + 1 }, () => new Array(n + 1).fill(0));
  for (let i = 1; i <= m; i++) {
    for (let j = 1; j <= n; j++) {
      if (origParas[i - 1].trim() === newParas[j - 1].trim()) {
        dp[i][j] = dp[i - 1][j - 1] + 1;
      } else {
        dp[i][j] = Math.max(dp[i - 1][j], dp[i][j - 1]);
      }
    }
  }

  // Backtrack LCS to get exact-match anchors
  const anchors = []; // { origIdx, newIdx }
  let i = m;
  let j = n;
  while (i > 0 && j > 0) {
    if (origParas[i - 1].trim() === newParas[j - 1].trim()) {
      anchors.push({ origIdx: i - 1, newIdx: j - 1 });
      i--;
      j--;
    } else if (dp[i][j - 1] >= dp[i - 1][j]) {
      j--;
    } else {
      i--;
    }
  }
  anchors.reverse();

  // Phase 2: Fill gaps between anchors with similarity-based matching.
  // Process each gap (segment between consecutive anchors) independently.
  /** @type {Array<{type: 'keep'|'delete'|'insert', origIdx?: number, newIdx?: number}>} */
  const ops = [];

  // Add sentinel anchors at boundaries
  const allAnchors = [
    { origIdx: -1, newIdx: -1 },
    ...anchors,
    { origIdx: m, newIdx: n },
  ];

  for (let a = 0; a < allAnchors.length - 1; a++) {
    const prev = allAnchors[a];
    const next = allAnchors[a + 1];

    const origStart = prev.origIdx + 1;
    const origEnd = next.origIdx;
    const newStart = prev.newIdx + 1;
    const newEnd = next.newIdx;

    // Greedy forward matching within this gap
    let oi = origStart;
    let ni = newStart;

    while (oi < origEnd && ni < newEnd) {
      const sim = paragraphSimilarity(origParas[oi], newParas[ni]);
      if (sim >= SIMILARITY_THRESHOLD) {
        // Similar enough: treat as modified paragraph (keep with text replacement)
        ops.push({ type: 'keep', origIdx: oi, newIdx: ni });
        oi++;
        ni++;
      } else {
        // Not similar: check if the next new paragraph matches better
        // (handles case where a new paragraph was inserted before the current original)
        let foundBetterMatch = false;
        if (ni + 1 < newEnd) {
          const nextSim = paragraphSimilarity(origParas[oi], newParas[ni + 1]);
          if (nextSim >= SIMILARITY_THRESHOLD) {
            // Insert the unmatched new paragraph, then match
            ops.push({ type: 'insert', newIdx: ni });
            ni++;
            foundBetterMatch = true;
            continue;
          }
        }
        if (!foundBetterMatch && oi + 1 < origEnd) {
          const nextOrigSim = paragraphSimilarity(origParas[oi + 1], newParas[ni]);
          if (nextOrigSim >= SIMILARITY_THRESHOLD) {
            // Delete the unmatched original paragraph, then match
            ops.push({ type: 'delete', origIdx: oi });
            oi++;
            foundBetterMatch = true;
            continue;
          }
        }
        if (!foundBetterMatch) {
          // Neither lookahead helps: treat as delete + insert
          ops.push({ type: 'delete', origIdx: oi });
          ops.push({ type: 'insert', newIdx: ni });
          oi++;
          ni++;
        }
      }
    }

    // Remaining unmatched originals are deletions
    while (oi < origEnd) {
      ops.push({ type: 'delete', origIdx: oi });
      oi++;
    }

    // Remaining unmatched new paragraphs are insertions
    while (ni < newEnd) {
      ops.push({ type: 'insert', newIdx: ni });
      ni++;
    }

    // Emit the next anchor as a keep (unless it's the end sentinel)
    if (next.origIdx < m && next.newIdx < n) {
      ops.push({ type: 'keep', origIdx: next.origIdx, newIdx: next.newIdx });
    }
  }

  return ops;
}
