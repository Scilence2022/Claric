
/**
 * Reassembler Module
 *
 * Applies LLM chunk results back to the Word document as tracked changes and comments.
 * After the orchestrator collects LLM responses for all chunks, the reassembler maps
 * those responses back to specific paragraph ranges in the Word document.
 *
 * Key behaviors:
 * - Bookmarks chunk paragraph ranges before LLM processing
 * - Amendments applied in reverse chunk order to prevent range invalidation
 * - Paragraph-level diff strategy preserves OOXML formatting (styles, numbering, indentation)
 * - Comments inserted after all amendments, on bookmarked ranges
 * - Failed/cancelled chunks skipped gracefully
 * - Individual bookmark cleanup with error tolerance
 *
 * @module reassembler
 */

import { applyTokenMapStrategy, applySentenceDiffStrategy } from './word-diff/index.js';
import { hasCjk, applyCharDiffStrategy } from './word-diff/char-diff.js';

/**
 * Generates a unique hidden bookmark name for chunk range persistence.
 * Format: _wdp + lowercase hex timestamp + hex chunk index + 3 random alphanumeric chars.
 * Hidden (underscore prefix), alphanumeric + underscore only.
 *
 * @param {number} chunkIndex - Index of the chunk
 * @returns {string}
 * @private
 */
function _generateChunkBookmarkName(chunkIndex) {
  const timestamp = Date.now().toString(16);
  const idx = chunkIndex.toString(16);
  const random = Math.random().toString(36).slice(2, 5).replace(/[^a-z0-9]/g, 'a');
  return `_wdp${timestamp}${idx}${random}`;
}

/**
 * Sorts chunk results by endIndex descending (reverse document order).
 * This ensures amendments are applied from the end of the document forward,
 * preventing range invalidation when text lengths change.
 *
 * @param {Array} results - ChunkResult array
 * @returns {Array} Sorted copy of results
 * @private
 */
function _sortReverseDocumentOrder(results) {
  return [...results].sort((a, b) => {
    const endA = a.chunk ? a.chunk.endIndex : 0;
    const endB = b.chunk ? b.chunk.endIndex : 0;
    return endB - endA;
  });
}

/**
 * Yields to the event loop to prevent UI freeze during long operations.
 * @returns {Promise<void>}
 * @private
 */
function _yieldToEventLoop() {
  return new Promise((resolve) => setTimeout(resolve, 0));
}

/**
 * Normalizes text for comparison by standardizing line endings.
 * Office.js range.text uses \r for paragraph breaks; LLM output uses \n.
 *
 * @param {string} text - Text to normalize
 * @returns {string} Normalized text with \n line endings
 * @private
 */
function _normalizeLineEndings(text) {
  return text.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
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
 * @private
 */
function _similarity(a, b) {
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
 * @private
 */
function _alignParagraphs(origParas, newParas) {
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
  const SIMILARITY_THRESHOLD = 0.4;

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
      const sim = _similarity(origParas[oi], newParas[ni]);
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
          const nextSim = _similarity(origParas[oi], newParas[ni + 1]);
          if (nextSim >= SIMILARITY_THRESHOLD) {
            // Insert the unmatched new paragraph, then match
            ops.push({ type: 'insert', newIdx: ni });
            ni++;
            foundBetterMatch = true;
            continue;
          }
        }
        if (!foundBetterMatch && oi + 1 < origEnd) {
          const nextOrigSim = _similarity(origParas[oi + 1], newParas[ni]);
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

/**
 * Applies a paragraph-level amendment strategy that preserves document formatting.
 *
 * Instead of operating on the full range text (which loses paragraph structure),
 * this strategy:
 * 1. Loads individual paragraphs from the chunk's paragraph range
 * 2. Splits the amended text by newlines to get amended paragraphs
 * 3. Aligns original paragraphs with amended paragraphs using LCS
 * 4. Replaces text within matched paragraphs (preserving styles/numbering)
 * 5. Deletes removed paragraphs
 * 6. Inserts new paragraphs after their predecessor
 *
 * Blank (empty) paragraphs are excluded from the alignment and left
 * untouched — the newline-joined amendment text cannot represent them,
 * so aligning them would mark every blank line as LLM-deleted.
 *
 * Falls back to range-level diff strategies if paragraph-level operations fail.
 *
 * @param {Word.RequestContext} context - The Word request context
 * @param {Word.Range} range - The bookmarked chunk range
 * @param {string} amendedText - The LLM's amended text (newline-delimited paragraphs)
 * @param {boolean} trackChangesEnabled - Whether to enable tracked changes
 * @param {boolean} lineDiffEnabled - Whether to use sentence-diff vs token-map for fallback
 * @param {function} log - Logging callback
 * @returns {Promise<boolean>} True when changes were written; false when the
 *   amended text matched the original (nothing to do)
 * @private
 */
async function _applyParagraphLevelAmendment(context, range, amendedText, trackChangesEnabled, lineDiffEnabled, log) {
  // Get paragraphs within the range
  const rangeParagraphs = range.paragraphs;
  rangeParagraphs.load('items');
  await context.sync();

  const allParaItems = rangeParagraphs.items;
  if (allParaItems.length === 0) {
    throw new Error('No paragraphs found in range');
  }

  // Load text for each paragraph
  for (const para of allParaItems) {
    para.load('text');
  }
  await context.sync();

  // Blank spacer paragraphs never enter the alignment: the amendment text
  // (newline-joined paragraphs) cannot represent them, so aligning them
  // would mark every blank line as LLM-deleted. Leave them untouched.
  const paraItems = allParaItems.filter((p) => p.text && p.text.trim() !== '');
  if (paraItems.length === 0) {
    log('Paragraph-level: range contains only blank paragraphs, nothing to amend');
    return false;
  }

  // Table membership per paragraph. Insert/delete alignment ops against a
  // table paragraph would add an in-cell paragraph or delete cell content —
  // never a table row — so those ops are skipped in the loop below (in-cell
  // 'keep' text edits are safe and still tracked).
  const tableChecks = paraItems.map((p) => {
    const t = p.parentTableOrNullObject;
    t.load('isNullObject');
    return t;
  });
  await context.sync();
  const inTable = tableChecks.map((t) => !t.isNullObject);

  const origTexts = paraItems.map((p) => p.text);
  const amendedLines = _normalizeLineEndings(amendedText).split('\n');

  // Filter out trailing empty lines from amended text (LLM sometimes adds trailing newline)
  while (amendedLines.length > 0 && amendedLines[amendedLines.length - 1].trim() === '') {
    amendedLines.pop();
  }

  // Also filter leading empty lines (LLM preamble artifacts)
  while (amendedLines.length > 0 && amendedLines[0].trim() === '') {
    amendedLines.shift();
  }

  // Content validation: detect severely truncated or corrupted LLM output
  // (inspired by superdoc-redlines validateNewText pattern)
  const origTotalChars = origTexts.reduce((sum, t) => sum + t.length, 0);
  const amendedTotalChars = amendedLines.reduce((sum, t) => sum + t.length, 0);
  if (origTotalChars > 0 && amendedTotalChars < origTotalChars * 0.3) {
    log(`Paragraph-level: LLM output appears truncated (${amendedTotalChars} chars vs ${origTotalChars} original), skipping`, 'warning');
    throw new Error('LLM output appears truncated (< 30% of original length)');
  }

  log(`Paragraph-level: ${origTexts.length} original paras, ${amendedLines.length} amended paras`);

  // Quick check: if all paragraphs are identical, skip
  if (origTexts.length === amendedLines.length &&
      origTexts.every((t, i) => t.trim() === amendedLines[i].trim())) {
    log('Paragraph-level: no changes detected, skipping');
    return false;
  }

  // Align paragraphs
  const alignment = _alignParagraphs(origTexts, amendedLines);

  // Set tracked changes explicitly (on OR off) so the whole alignment loop —
  // including the per-paragraph diff strategies — runs under one mode.
  if (Word.ChangeTrackingMode) {
    context.document.changeTrackingMode = trackChangesEnabled
      ? Word.ChangeTrackingMode.trackAll
      : Word.ChangeTrackingMode.off;
  }

  // Process alignment operations in REVERSE order to prevent index invalidation.
  // We iterate from the end of the document upward.
  const reversedOps = [...alignment].reverse();

  for (const op of reversedOps) {
    if (op.type === 'keep') {
      // Text matched at paragraph level -- but there might be minor word-level edits.
      // Compare trimmed text; if different, apply word-level diff within the paragraph
      // to preserve run-level formatting (bold, italic, font, color).
      const origText = origTexts[op.origIdx];
      const newText = amendedLines[op.newIdx];

      if (origText.trim() !== newText.trim()) {
        const para = paraItems[op.origIdx];
        const paraRange = para.getRange('Content');
        paraRange.load('text');
        await context.sync();

        // Use word-level token map strategy scoped to single paragraph.
        // At paragraph scope, token map is much more reliable:
        // - no \r/\n mismatch (no paragraph breaks)
        // - smaller token count = fewer alignment errors
        // This preserves run-level formatting (w:rPr) while applying tracked changes.
        // CJK text has no word boundaries for the token map (a whole sentence
        // becomes one token), so it uses the char-level strategy instead.
        // The outer scope already owns the tracking mode (set above, restored
        // below), so the strategy must not clobber it mid-loop.
        try {
          const diffOptions = { trackChanges: false };
          if (hasCjk(paraRange.text) || hasCjk(newText)) {
            await applyCharDiffStrategy(context, paraRange, paraRange.text, newText.trim(), log, diffOptions);
          } else {
            await applyTokenMapStrategy(context, paraRange, paraRange.text, newText.trim(), log, diffOptions);
          }
        } catch (_diffErr) {
          // If word-level diff fails, fall back to full paragraph text replacement.
          // This loses run-level formatting but preserves paragraph-level properties.
          log(`Para ${op.origIdx}: word-level diff failed (${_diffErr.message}), using text replacement`, 'warning');
          paraRange.insertText(newText.trim(), Word.InsertLocation.replace);
          await context.sync();
        }
      }
    } else if (op.type === 'delete') {
      // Paragraph was removed by LLM -- delete it. Table paragraphs are
      // skipped: deleting cell content never deletes a row, it corrupts one.
      if (inTable[op.origIdx]) {
        log(`Para ${op.origIdx}: skipping delete — paragraph is inside a table`, 'warning');
        continue;
      }
      const para = paraItems[op.origIdx];
      para.delete();
    } else if (op.type === 'insert') {
      // New paragraph from LLM -- insert after the preceding original paragraph.
      // Find the last 'keep' or 'delete' op before this one that references an origIdx.
      const insertText = amendedLines[op.newIdx].trim();
      if (!insertText) continue; // Skip empty inserted lines

      // Find the anchor: the original paragraph immediately before this insertion point.
      // Walk backwards through alignment to find the most recent origIdx.
      let anchorOrigIdx = -1;
      const opIndex = alignment.indexOf(op);
      for (let k = opIndex - 1; k >= 0; k--) {
        if (alignment[k].origIdx !== undefined) {
          anchorOrigIdx = alignment[k].origIdx;
          break;
        }
      }

      if (anchorOrigIdx >= 0 && anchorOrigIdx < paraItems.length) {
        // Anchoring inside a table would insert an in-cell paragraph, not a
        // new row — skip instead of corrupting the table layout.
        if (inTable[anchorOrigIdx]) {
          log(`Skipping insert after para ${anchorOrigIdx} — anchor is inside a table`, 'warning');
          continue;
        }
        const anchorPara = paraItems[anchorOrigIdx];
        anchorPara.insertParagraph(insertText, Word.InsertLocation.after);
      } else if (paraItems.length > 0) {
        // Insert before the first paragraph
        paraItems[0].insertParagraph(insertText, Word.InsertLocation.before);
      }
    }
  }

  await context.sync();

  // Disable tracked changes
  if (Word.ChangeTrackingMode && trackChangesEnabled) {
    context.document.changeTrackingMode = Word.ChangeTrackingMode.off;
    await context.sync();
  }

  log('Paragraph-level amendment applied successfully');
  return true;
}

/**
 * Bookmarks each chunk's paragraph range before LLM processing.
 * Called once after parsing/chunking, before sending to orchestrator.
 * Bookmarks persist in the document and survive LLM processing time.
 *
 * @param {Array} chunks - DocumentChunk[] with startIndex/endIndex
 * @returns {Promise<Map<string, string>>} Map of chunkId -> bookmarkName
 */
export async function bookmarkChunkRanges(chunks) {
  const bookmarkMap = new Map();

  await Word.run(async (context) => {
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load('items');
    await context.sync();

    for (let i = 0; i < chunks.length; i++) {
      const chunk = chunks[i];
      const startPara = paragraphs.items[chunk.startIndex];
      const endPara = paragraphs.items[chunk.endIndex];

      const startRange = startPara.getRange('Start');
      const endRange = endPara.getRange('End');
      const fullRange = startRange.expandTo(endRange);

      const bookmarkName = _generateChunkBookmarkName(i);
      fullRange.insertBookmark(bookmarkName);

      bookmarkMap.set(chunk.id, bookmarkName);
    }

    await context.sync();
  });

  return bookmarkMap;
}

/**
 * Finds the stored original paragraph sequence as a contiguous window within
 * the range's current paragraph texts (trim-compared). A staged bookmark
 * range can absorb paragraphs inserted after staging (e.g. a title added by
 * another proposal card), shifting the original content within the bookmark.
 *
 * @param {string[]} currentTexts - Paragraph texts currently inside the range
 * @param {string[]} storedTexts - Paragraph texts captured at staging time
 * @returns {{start: number, end: number} | null} Window [start, end) within
 *   currentTexts matching the full stored sequence, or null when no
 *   contiguous match exists
 * @private
 */
function _findAnchorWindow(currentTexts, storedTexts) {
  const stored = storedTexts.map((t) => (t || '').trim());
  const current = currentTexts.map((t) => (t || '').trim());
  if (stored.length === 0 || stored.length > current.length) return null;

  for (let start = 0; start + stored.length <= current.length; start++) {
    let match = true;
    for (let k = 0; k < stored.length; k++) {
      if (current[start + k] !== stored[k]) {
        match = false;
        break;
      }
    }
    if (match) return { start, end: start + stored.length };
  }
  return null;
}

/**
 * Resolves the staged original paragraph texts for a chunk result.
 * Prefers the explicit chunkOriginals map (the retry path rebuilds chunks
 * without paragraphs); falls back to result.chunk.paragraphs.
 *
 * @param {Object} result - ChunkResult
 * @param {Map<string, string[]> | null} chunkOriginals - chunkId -> texts
 * @returns {string[] | null} Stored paragraph texts, or null when unavailable
 * @private
 */
function _storedParagraphTexts(result, chunkOriginals) {
  if (chunkOriginals && typeof chunkOriginals.has === 'function' && chunkOriginals.has(result.chunkId)) {
    const texts = chunkOriginals.get(result.chunkId);
    if (Array.isArray(texts) && texts.length > 0) return texts;
  }
  const chunk = result.chunk;
  if (chunk && Array.isArray(chunk.paragraphs) && chunk.paragraphs.length > 0) {
    return chunk.paragraphs.map((p) => p.text);
  }
  return null;
}

/**
 * Re-anchors a staged bookmark range to the stored original paragraphs.
 * When the bookmark absorbed paragraphs inserted after staging, narrows the
 * working range to just the window holding the original content, so the
 * amendment never deletes absorbed content (e.g. a title inserted by another
 * proposal card).
 *
 * Empty (blank) paragraphs are ignored: the parser skips them, so the stored
 * sequence never contains them, while range.paragraphs includes them.
 *
 * @param {Word.RequestContext} context
 * @param {Word.Range} range - The bookmarked chunk range
 * @param {string[]} storedTexts - Paragraph texts captured at staging time
 * @param {function} log
 * @returns {Promise<Word.Range | null>} The (possibly narrowed) range, or
 *   null when the stored sequence is no longer contiguously locatable
 * @private
 */
async function _reanchorChunkRange(context, range, storedTexts, log) {
  // Ranges without a paragraphs collection (test mocks, exotic bookmarks)
  // cannot be re-anchored; use them as-is.
  if (!range.paragraphs) return range;

  const rangeParagraphs = range.paragraphs;
  rangeParagraphs.load('items');
  await context.sync();

  const paraItems = rangeParagraphs.items;
  if (paraItems.length === 0) return range;

  for (const para of paraItems) {
    para.load('text');
  }
  await context.sync();

  // Compare on the non-empty paragraphs only: blank spacer paragraphs are
  // absent from the staged sequence but present in the live range.
  const nonEmptyIdx = [];
  const currentTexts = [];
  paraItems.forEach((para, i) => {
    if (para.text && para.text.trim() !== '') {
      nonEmptyIdx.push(i);
      currentTexts.push(para.text);
    }
  });

  const window = _findAnchorWindow(currentTexts, storedTexts);
  if (!window) return null;

  const startItemIdx = nonEmptyIdx[window.start];
  const endItemIdx = nonEmptyIdx[window.end - 1];
  if (startItemIdx === 0 && endItemIdx === paraItems.length - 1) {
    return range; // No drift: original content still spans the whole range
  }

  log(
    `Range drifted since staging; narrowed to paragraphs ${startItemIdx + 1}-${endItemIdx + 1} of ${paraItems.length}`,
    'info'
  );
  const startRange = paraItems[startItemIdx].getRange('Start');
  const endRange = paraItems[endItemIdx].getRange('End');
  const narrowed = startRange.expandTo(endRange);
  narrowed.load('text');
  await context.sync();
  return narrowed;
}

/**
 * Applies all chunk results to the document.
 * Amendments applied in reverse chunk order as tracked changes.
 * Uses paragraph-level strategy to preserve formatting; falls back to
 * range-level diff strategies if paragraph-level operations fail.
 * Comments inserted after all amendments, on bookmarked ranges.
 *
 * @param {Array} results - ChunkResult[]
 * @param {Map<string, string>} bookmarkMap - chunkId -> bookmarkName
 * @param {Object} options
 * @param {boolean} options.trackChangesEnabled
 * @param {boolean} options.lineDiffEnabled - use sentence-diff vs token-map for fallback
 * @param {function} options.log
 * @param {Map<string, string[]>} [options.chunkOriginals] - chunkId -> staged
 *   paragraph texts, used to re-anchor ranges that drifted since staging
 * @param {AbortSignal} [options.signal] - Cooperative pause: when aborted, the
 *   loop finishes the in-flight chunk then stops before the next, leaving the
 *   remaining chunks' bookmarks intact so the caller can resume later
 * @param {function(string, {applied: boolean, noChange?: boolean, error?: boolean, skipped?: boolean}): void} [options.onChunkApplied] -
 *   Called after each chunk's apply attempt with its chunkId + outcome, so the
 *   UI can mark individual items applied as they land
 * @returns {Promise<{amendmentsApplied: number, commentsInserted: number, noChangeCount: number,
 *   errors: string[], appliedChunkIds: string[], interrupted: boolean}>}
 */
export async function applyChunkResults(results, bookmarkMap, options) {
  const {
    trackChangesEnabled = true,
    lineDiffEnabled = false,
    log = () => {},
    chunkOriginals = null,
    signal = null,
    onChunkApplied = /** @type {function(string, object): void} */ (() => {}),
  } = options;

  let amendmentsApplied = 0;
  let commentsInserted = 0;
  let noChangeCount = 0;
  const errors = [];
  const appliedChunkIds = [];
  let interrupted = false;

  // Collect rejected/cancelled errors for reporting
  for (const result of results) {
    if (result.status === 'rejected' && result.error) {
      errors.push(`Chunk ${result.chunkId}: ${result.error}`);
    }
  }

  // Phase 1: Amendments in reverse document order
  const fulfilledWithAmendments = results
    .filter((r) => r.status === 'fulfilled' && r.amendment)
    .slice();

  const reverseSorted = _sortReverseDocumentOrder(fulfilledWithAmendments);

  for (const result of reverseSorted) {
    // Cooperative pause: the Stop button aborts the signal; we finish the
    // current chunk (already in-flight Word.run) and stop at the next
    // boundary, leaving the remaining chunks' bookmarks intact for resume.
    if (signal && signal.aborted) {
      interrupted = true;
      break;
    }

    const bookmarkName = bookmarkMap.get(result.chunkId);
    if (!bookmarkName) {
      errors.push(`Chunk ${result.chunkId}: no bookmark found`);
      appliedChunkIds.push(result.chunkId);
      onChunkApplied(result.chunkId, { applied: false, skipped: true });
      continue;
    }

    try {
      let chunkApplied = false;
      await Word.run(async (context) => {
        const range = context.document.getBookmarkRangeOrNullObject(bookmarkName);
        range.load('isNullObject,text');
        await context.sync();

        if (range.isNullObject) {
          errors.push(`Chunk ${result.chunkId}: bookmark range lost`);
          return;
        }

        // Re-anchor: the bookmark may have absorbed paragraphs inserted after
        // staging (e.g. a title from another proposal card). Narrow the
        // working range to the stored original paragraphs so the amendment
        // never deletes absorbed content.
        const storedTexts = _storedParagraphTexts(result, chunkOriginals);
        let workRange = range;
        if (storedTexts) {
          let anchored;
          try {
            anchored = await _reanchorChunkRange(context, range, storedTexts, log);
          } catch (anchorErr) {
            // A failed drift check means we cannot tell whether the range
            // absorbed new content; falling back to the raw bookmark range
            // would risk deleting absorbed paragraphs. Skip instead.
            errors.push(`Chunk ${result.chunkId}: re-anchor check failed (${anchorErr.message}); amendment skipped`);
            log(`Chunk ${result.chunkId}: re-anchor check failed (${anchorErr.message}), skipping to avoid deleting absorbed content`, 'warning');
            return;
          }
          if (anchored === null) {
            errors.push(`Chunk ${result.chunkId}: original content no longer matches the staged range (edited since staging?); amendment skipped`);
            log(`Chunk ${result.chunkId}: original content not found contiguously in staged range, skipping to avoid deleting absorbed content`, 'warning');
            return;
          }
          workRange = anchored;
        }

        // Try paragraph-level strategy first (preserves formatting)
        let applied;
        try {
          applied = await _applyParagraphLevelAmendment(
            context, workRange, result.amendment,
            trackChangesEnabled, lineDiffEnabled, log
          ) !== false;
        } catch (paraErr) {
          log(`Chunk ${result.chunkId}: paragraph-level strategy failed (${paraErr.message}), falling back to range-level`, 'warning');

          // Fallback to range-level diff strategies
          if (Word.ChangeTrackingMode) {
            context.document.changeTrackingMode = trackChangesEnabled
              ? Word.ChangeTrackingMode.trackAll
              : Word.ChangeTrackingMode.off;
          }

          // Normalize line endings for consistent diffing
          const originalText = _normalizeLineEndings(workRange.text);
          const normalizedAmendment = _normalizeLineEndings(result.amendment);

          const strategyOptions = { trackChanges: trackChangesEnabled };
          if (lineDiffEnabled) {
            await applySentenceDiffStrategy(context, workRange, originalText, normalizedAmendment, log, strategyOptions);
          } else if (hasCjk(originalText) || hasCjk(normalizedAmendment)) {
            // Same CJK rule as the paragraph-level path: a whole CJK run is a
            // single token to the word-level strategies, which would turn a
            // one-comma edit into a whole-range replacement redline.
            await applyCharDiffStrategy(context, workRange, originalText, normalizedAmendment, log, strategyOptions);
          } else {
            await applyTokenMapStrategy(context, workRange, originalText, normalizedAmendment, log, strategyOptions);
          }

          // Disable tracked changes after fallback (matching paragraph-level strategy behavior)
          if (Word.ChangeTrackingMode && trackChangesEnabled) {
            context.document.changeTrackingMode = Word.ChangeTrackingMode.off;
            await context.sync();
          }
          applied = true;
        }

        if (!applied) {
          noChangeCount++;
        } else {
          chunkApplied = true;
        }
      });

      if (chunkApplied) {
        amendmentsApplied++;
        log(`Chunk ${result.chunkId}: amendment applied`, 'info');
      } else {
        log(`Chunk ${result.chunkId}: no changes needed`, 'info');
      }
      appliedChunkIds.push(result.chunkId);
      onChunkApplied(result.chunkId, { applied: chunkApplied, noChange: !chunkApplied });
    } catch (err) {
      errors.push(`Chunk ${result.chunkId}: ${err.message || String(err)}`);
      log(`Chunk ${result.chunkId}: amendment failed -- ${err.message}`, 'error');
      appliedChunkIds.push(result.chunkId);
      onChunkApplied(result.chunkId, { applied: false, error: true });
    }

    // Yield to event loop between chunks to prevent UI freeze
    await _yieldToEventLoop();
  }

  // Phase 2: Comments in document order (after all amendments).
  // Skipped entirely when the apply was paused mid-way — the resume re-runs
  // this for the remaining chunks.
  if (!interrupted) {
    // Ensure tracked changes are off before inserting comments.
    // If any amendment fallback path left ChangeTrackingMode.trackAll enabled,
    // comment insertion on ranges containing tracked changes can fail with AccessDenied.
    if (fulfilledWithAmendments.length > 0) {
      try {
        await Word.run(async (context) => {
          if (Word.ChangeTrackingMode) {
            context.document.changeTrackingMode = Word.ChangeTrackingMode.off;
            await context.sync();
          }
        });
      } catch (_err) {
        // Best-effort -- continue with comment insertion even if this fails
      }
    }

    const fulfilledWithComments = results
      .filter((r) => r.status === 'fulfilled' && r.comment);

    for (const result of fulfilledWithComments) {
      const bookmarkName = bookmarkMap.get(result.chunkId);
      if (!bookmarkName) continue;

      try {
        await Word.run(async (context) => {
          const range = context.document.getBookmarkRangeOrNullObject(bookmarkName);
          range.load('isNullObject,text');
          await context.sync();

          if (range.isNullObject) {
            errors.push(`Chunk ${result.chunkId}: bookmark range lost for comment`);
            return;
          }

          range.insertComment(result.comment);
          await context.sync();
        });

        commentsInserted++;
        log(`Chunk ${result.chunkId}: comment inserted`, 'info');
      } catch (err) {
        errors.push(`Chunk ${result.chunkId}: comment failed -- ${err.message || String(err)}`);
        log(`Chunk ${result.chunkId}: comment failed -- ${err.message}`, 'error');
      }

      // Yield to event loop between comments to prevent UI freeze and
      // avoid overwhelming the Word document model with rapid-fire Word.run() calls
      await _yieldToEventLoop();
    }
  }

  return { amendmentsApplied, commentsInserted, noChangeCount, errors, appliedChunkIds, interrupted };
}

/**
 * Removes all chunk bookmarks from the document.
 * Tolerates individual bookmark deletion failures.
 *
 * @param {Map<string, string>} bookmarkMap - chunkId -> bookmarkName
 * @returns {Promise<void>}
 */
export async function cleanupBookmarks(bookmarkMap) {
  await Word.run(async (context) => {
    for (const bookmarkName of bookmarkMap.values()) {
      try {
        context.document.deleteBookmark(bookmarkName);
      } catch (_err) {
        // Tolerate individual bookmark deletion failures
      }
    }
    await context.sync();
  });
}

// Export internals for testing
export { _normalizeLineEndings, _alignParagraphs, _findAnchorWindow };
