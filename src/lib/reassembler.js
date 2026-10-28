
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
import { alignParagraphs, amendmentParagraphTexts } from './paragraph-alignment.js';
import { queueRevisionRead, resolveRevisionRead, RevisionSafetyError, verifyMutationText, hostSupportsTrackedChanges,
  normalizeWordText, wordErrorDetail } from './word-revisions.js';
import { validateWhitespaceCleanup } from './whitespace-cleanup.js';
import { tryRevisionDiff } from './word-diff/revision-diff-loader.js';

/**
 * Thrown when an LLM amendment fails the content-length sanity check (the
 * < 30% of original guard). Distinct from a strategy failure on purpose:
 * falling back to a range-level strategy with the same truncated text would
 * still write the truncated content — only more crudely — so the strategy
 * dispatcher escalates this to the chunk-level error path and leaves the
 * document untouched instead of downgrading the apply strategy.
 */
export class TruncatedOutputError extends Error {
  constructor(message) {
    super(message);
    this.name = 'TruncatedOutputError';
  }
}

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
  return normalizeWordText(text);
}

// Paragraph pairing is shared with the amendment guards (paragraph-alignment.js).
const _alignParagraphs = alignParagraphs;
const NATIVE_READ_BATCH_SIZE = 24;
const _paragraphTargetText = (text, whitespaceOnly) => whitespaceOnly ? text : text.replace(/^[ \u00a0]+|[ \u00a0]+$/g, '');

/** Full paragraph OOXML packages must stay bounded during apply as during parse. */
async function _readParagraphStates(context, paraItems) {
  const inTable = [];
  const revisionStates = [];
  for (let start = 0; start < paraItems.length; start += NATIVE_READ_BATCH_SIZE) {
    const batch = paraItems.slice(start, start + NATIVE_READ_BATCH_SIZE);
    const reads = batch.map((para) => {
      para.load('text');
      return queueRevisionRead(para, { paragraph: true, singleParagraphContent: true });
    });
    const tables = batch.map((para) => {
      const table = para.parentTableOrNullObject;
      table.load('isNullObject');
      return table;
    });
    await context.sync();
    inTable.push(...tables.map((table) => !table.isNullObject));
    revisionStates.push(...batch.map((para, index) => resolveRevisionRead(para, reads[index], { paragraph: true, singleParagraphContent: true })));
  }
  return { inTable, revisionStates };
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
 * @param {boolean} lineDiffEnabled - When true, per-paragraph edits use the
 *   sentence-diff strategy (Settings: Force Line Diff / Sentence Mode) and
 *   the range-level fallback does too; the default uses token-map for
 *   spaced scripts and char-diff for CJK
 * @param {function} log - Logging callback
 * @param {{paraItems: Word.Paragraph[], inTable: boolean[], revisionStates?: any[]}} [preloaded] -
 *   Paragraph reads already batched by the re-anchor pass (items with loaded
 *   texts + table membership). When provided, this function performs no read
 *   syncs of its own before the alignment. Omit for ranges without a
 *   paragraphs collection (the reads then run here, and a range without
 *   paragraphs throws into the caller's range-level fallback).
 * @returns {Promise<boolean>} True when changes were written; false when the
 *   amended text matched the original (nothing to do)
 * @private
 */
async function _applyParagraphLevelAmendment(context, range, amendedText, trackChangesEnabled, lineDiffEnabled, log, preloaded = null,
  { whitespaceOnly = false, expectedTexts = null,
    progress = { report: () => {}, appliedParagraphs: 0, mutationAttempted: false, pendingVerification: false, pendingStructuralWrites: false,
      stage: 'paragraph read' } } = {}) {
  let allParaItems;
  let inTable;
  let revisionStates;
  if (preloaded) {
    allParaItems = preloaded.paraItems;
    inTable = preloaded.inTable;
    revisionStates = preloaded.revisionStates;
  } else {
    // Get paragraphs within the range
    const rangeParagraphs = range.paragraphs;
    rangeParagraphs.load('items');
    await context.sync();

    allParaItems = rangeParagraphs.items;
    if (allParaItems.length === 0) {
      throw new Error('No paragraphs found in range');
    }

    // Paragraph texts and table membership in one batched read. Table
    // membership gates the insert/delete ops below: against a table
    // paragraph they would add an in-cell paragraph or delete cell content —
    // never a table row (in-cell 'keep' text edits are safe and still
    // tracked).
    ({ inTable, revisionStates } = await _readParagraphStates(context, allParaItems));
  }

  // Blank spacer paragraphs never enter the alignment: the amendment text
  // (newline-joined paragraphs) cannot represent them, so aligning them
  // would mark every blank line as LLM-deleted. Leave them untouched.
  const paraItems = [];
  const paraInTable = [];
  const states = [];
  allParaItems.forEach((p, i) => {
    const state = revisionStates?.[i] || resolveRevisionRead(p, null, { paragraph: true, singleParagraphContent: true });
    if (state.text.trim() !== '') {
      paraItems.push(p);
      paraInTable.push(inTable[i]);
      states.push(state);
    }
  });
  if (paraItems.length === 0) {
    log('Paragraph-level: range contains only blank paragraphs, nothing to amend');
    return false;
  }
  inTable = paraInTable;

  const origTexts = states.map((s) => s.text);
  let amendedLines;
  try { amendedLines = amendmentParagraphTexts(origTexts, amendedText); }
  catch (error) { throw new RevisionSafetyError(error.message); }

  if (whitespaceOnly) {
    if (expectedTexts && (expectedTexts.length !== origTexts.length
        || expectedTexts.some((text, index) => _normalizeLineEndings(text) !== _normalizeLineEndings(origTexts[index])))) {
      throw new Error('The whitespace cleanup source changed since staging. Draft a new proposal.');
    }
    const check = validateWhitespaceCleanup(origTexts.join('\n'), amendedText);
    if (!check.valid) throw new Error(check.reason);
  }

  // Filter out trailing empty lines from amended text (LLM sometimes adds trailing newline)
  while (!whitespaceOnly && amendedLines.length > 0 && amendedLines[amendedLines.length - 1].trim() === '') {
    amendedLines.pop();
  }

  // Also filter leading empty lines (LLM preamble artifacts)
  while (!whitespaceOnly && amendedLines.length > 0 && amendedLines[0].trim() === '') {
    amendedLines.shift();
  }

  // Content validation: detect severely truncated or corrupted LLM output
  // (inspired by superdoc-redlines validateNewText pattern)
  const origTotalChars = origTexts.reduce((sum, t) => sum + t.length, 0);
  const amendedTotalChars = amendedLines.reduce((sum, t) => sum + t.length, 0);
  if (origTotalChars > 0 && amendedTotalChars < origTotalChars * 0.3) {
    log(`Paragraph-level: LLM output appears truncated (${amendedTotalChars} chars vs ${origTotalChars} original), skipping`, 'warning');
    throw new TruncatedOutputError('LLM output appears truncated (< 30% of original length)');
  }

  log(`Paragraph-level: ${origTexts.length} original paras, ${amendedLines.length} amended paras`);

  // Quick check: if all paragraphs are identical, skip
  const comparableText = (text) => whitespaceOnly ? _normalizeLineEndings(text) : text.trim();
  if (origTexts.length === amendedLines.length &&
      origTexts.every((t, i) => comparableText(t) === comparableText(amendedLines[i]))) {
    log('Paragraph-level: no changes detected, skipping');
    return false;
  }

  // Align paragraphs
  // With unchanged paragraph count, revise pending text in place even when
  // a substantial rewrite falls below the fuzzy alignment threshold. A
  // delete+insert interpretation would otherwise touch its revision marks.
  /** @type {Array<{type: 'keep'|'delete'|'insert', origIdx?: number, newIdx?: number}>} */
  const alignment = origTexts.length === amendedLines.length && (whitespaceOnly || states.some((s) => s.hasRevisions || s.protectedSpans?.length))
    ? origTexts.map((_text, index) => ({ type: 'keep', origIdx: index, newIdx: index }))
    : _alignParagraphs(origTexts, amendedLines);
  if (alignment.some((op) => op.type === 'delete' && states[op.origIdx].hasRevisions)) {
    throw new RevisionSafetyError('Deleting a paragraph with unresolved revisions requires a fresh in-place text edit.');
  }
  if (alignment.some((op) => op.type === 'delete' && states[op.origIdx].protectedSpans?.length)) {
    throw new RevisionSafetyError('Cannot delete a paragraph containing native fields through a prose amendment. Preserve its field results.');
  }
  let structuralAnchor = 0;
  for (const op of alignment) {
    if (op.type === 'delete' && inTable[op.origIdx]) {
      throw new RevisionSafetyError('Cannot delete a table paragraph through a prose amendment. Preserve its paragraph structure.');
    }
    if (op.type === 'insert' && amendedLines[op.newIdx].trim() && inTable[structuralAnchor]) {
      throw new RevisionSafetyError('Cannot insert a prose paragraph inside a table. Preserve its paragraph structure.');
    }
    if (op.origIdx !== undefined) structuralAnchor = op.origIdx;
  }

  // Resolve changed content in bounded read batches before the first write.
  // Each write still revalidates its native text and target positions.
  /** @type {Map<number, Word.Range>} */
  const changedParaRanges = new Map();
  const contentReads = new Map();
  const contentStates = new Map();
  const changedOps = alignment.filter((op) => op.type === 'keep'
    && comparableText(origTexts[op.origIdx]) !== comparableText(amendedLines[op.newIdx]));
  progress.stage = 'paragraph preflight';
  progress.report?.();
  for (let start = 0; start < changedOps.length; start += NATIVE_READ_BATCH_SIZE) {
    const batch = changedOps.slice(start, start + NATIVE_READ_BATCH_SIZE);
    for (const op of batch) {
      const paraRange = paraItems[op.origIdx].getRange('Content');
      paraRange.load('text');
      changedParaRanges.set(op.origIdx, paraRange);
      contentReads.set(op.origIdx, queueRevisionRead(paraRange, { paragraph: true, singleParagraphContent: true }));
    }
    await context.sync();
    for (const op of batch) contentStates.set(op.origIdx,
      resolveRevisionRead(changedParaRanges.get(op.origIdx), contentReads.get(op.origIdx), { paragraph: true, singleParagraphContent: true }));
  }
  if (changedParaRanges.size > 0) {
    // Validate every native content range before the first write. Paragraph
    // exports and content-range exports can differ at their end boundaries;
    // a mismatch must not leave the earlier reverse-order edits half applied.
    for (const [index, read] of contentReads) {
      if (!read) continue;
      const state = contentStates.get(index);
      if (state.text !== origTexts[index]) {
        throw new RevisionSafetyError(`Paragraph ${index + 1}: native content differs from the staged baseline. Draft a new proposal.`);
      }
      if (state.protectedStructure && !state.fieldOnlyProtection) {
        throw new RevisionSafetyError(`Paragraph ${index + 1}: structural revisions or protected objects cannot be safely edited in place.`);
      }
      const op = alignment.find((op) => op.type === 'keep' && op.origIdx === index);
      const after = _paragraphTargetText(amendedLines[op.newIdx], whitespaceOnly);
      await tryRevisionDiff(context, changedParaRanges.get(index), origTexts[index], after, log,
        { trackChanges: false, paragraph: true, preflightOnly: true, nativeTargets: true });
    }
  }

  // Set tracked changes explicitly (on OR off) so the whole alignment loop —
  // including the per-paragraph diff strategies — runs under one mode.
  if (trackChangesEnabled && !hostSupportsTrackedChanges()) {
    throw new RevisionSafetyError('This Word host cannot enable tracked changes.');
  }
  const previousTrackingMode = context.document.changeTrackingMode ?? Word.ChangeTrackingMode?.off;
  if (Word.ChangeTrackingMode) {
    context.document.changeTrackingMode = trackChangesEnabled
      ? Word.ChangeTrackingMode.trackAll
      : Word.ChangeTrackingMode.off;
  }

  // Process alignment operations in REVERSE order to prevent index invalidation.
  // We iterate from the end of the document upward.
  const reversedOps = [...alignment].reverse();

  try {
    // Confirm the requested mode before queuing structural writes too;
    // text strategies perform their own reads, but insert/delete may not.
    if (Word.ChangeTrackingMode) {
      try { await context.sync(); }
      catch (error) { throw new RevisionSafetyError(`Could not set Word's revision mode: ${error.message}`); }
    }
    for (const op of reversedOps) {
      if (op.type === 'keep') {
        // Text matched at paragraph level -- but there might be minor word-level edits.
        // Compare trimmed text; if different, apply word-level diff within the paragraph
        // to preserve run-level formatting (bold, italic, font, color).
        const origText = origTexts[op.origIdx];
        const newText = amendedLines[op.newIdx];
        const targetText = _paragraphTargetText(newText, whitespaceOnly);

        if (comparableText(origText) !== comparableText(newText)) {
          progress.stage = `paragraph ${op.origIdx + 1}: native text edit`;
          progress.report?.();
          const paraRange = changedParaRanges.get(op.origIdx);

          // Use word-level token map strategy scoped to single paragraph.
          // At paragraph scope, token map is much more reliable:
          // - no \r/\n mismatch (no paragraph breaks)
          // - smaller token count = fewer alignment errors
          // This preserves run-level formatting (w:rPr) while applying tracked changes.
          // CJK text has no word boundaries for the token map (a whole sentence
          // becomes one token), so it uses the char-level strategy instead.
          // The Settings toggle (Force Line Diff / Sentence Mode) forces the
          // sentence-level strategy — the same ordering as the selection-scope
          // apply path.
          // The outer scope already owns the tracking mode (set above, restored
          // below), so the strategy must not clobber it mid-loop.
          try {
            const diffOptions = { trackChanges: false, paragraph: true, nativeTargets: true, verificationParagraph: paraItems[op.origIdx] };
            if (whitespaceOnly) {
              await applyCharDiffStrategy(context, paraRange, origText, targetText, log, diffOptions);
            } else if (lineDiffEnabled) {
              await applySentenceDiffStrategy(context, paraRange, origText, targetText, log, diffOptions);
            } else if (hasCjk(origText) || hasCjk(newText)) {
              await applyCharDiffStrategy(context, paraRange, origText, targetText, log, diffOptions);
            } else {
              await applyTokenMapStrategy(context, paraRange, origText, targetText, log, diffOptions);
            }
          } catch (_diffErr) {
            if (_diffErr.mutationAttempted) progress.mutationAttempted = true;
            if (whitespaceOnly || _diffErr instanceof RevisionSafetyError || _diffErr.mutationAttempted
                || states[op.origIdx].hasRevisions) throw _diffErr;
            // If word-level diff fails, fall back to full paragraph text replacement.
            // This loses run-level formatting but preserves paragraph-level properties.
            log(`Para ${op.origIdx}: word-level diff failed (${_diffErr.message}), using text replacement`, 'warning');
            progress.mutationAttempted = true;
            progress.pendingVerification = true;
            paraRange.insertText(targetText, Word.InsertLocation.replace);
            await context.sync();
          }
          progress.mutationAttempted = true;
          progress.pendingVerification = true;
          progress.stage = `paragraph ${op.origIdx + 1}: read-back`;
          progress.report?.();
          // Preflight proved the prose against native ranges and protected
          // these inline objects. Preserve that evidence for the outer check;
          // raw reviewed text alone may use a different equation serializer.
          await verifyMutationText(context, paraRange, targetText,
            { paragraph: true, verificationParagraph: paraItems[op.origIdx],
              verifiedInlineFingerprint: contentStates.get(op.origIdx)?.inlineFingerprint });
          progress.pendingVerification = false;
          progress.appliedParagraphs++;
          progress.report?.();
        }
      } else if (op.type === 'delete') {
        progress.stage = `paragraph ${op.origIdx + 1}: structural delete`;
        progress.report?.();
        // Structural table operations were refused before the first write.
        const para = paraItems[op.origIdx];
        progress.mutationAttempted = true;
        progress.pendingStructuralWrites = true;
        para.delete();
      } else if (op.type === 'insert') {
        progress.stage = `output paragraph ${op.newIdx + 1}: structural insert`;
        progress.report?.();
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
          const anchorPara = paraItems[anchorOrigIdx];
          progress.mutationAttempted = true;
          progress.pendingStructuralWrites = true;
          anchorPara.insertParagraph(insertText, Word.InsertLocation.after);
        } else if (paraItems.length > 0) {
          // Insert before the first paragraph
          progress.mutationAttempted = true;
          progress.pendingStructuralWrites = true;
          paraItems[0].insertParagraph(insertText, Word.InsertLocation.before);
        }
      }
    }

    await context.sync();
    progress.pendingStructuralWrites = false;
  } finally {
    // A revision-safe failure must not leak the caller's tracking mode.
    if (Word.ChangeTrackingMode) {
      context.document.changeTrackingMode = previousTrackingMode;
      await context.sync();
    }
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

    // Boundary validation: the chunk indexes come from a PREVIOUS Word.run
    // (parseDocument). Any paragraph add/delete in between — a concurrent
    // card apply, a comment-queue flush, the user typing — silently shifts
    // them, and amendments would land on the wrong paragraphs. Verify the
    // boundary paragraphs still hold the chunk's expected texts before
    // bookmarking anything.
    const boundaryIndexes = new Set();
    for (const chunk of chunks) {
      boundaryIndexes.add(chunk.startIndex);
      boundaryIndexes.add(chunk.endIndex);
    }
    const boundaryParagraphs = new Map();
    for (const idx of boundaryIndexes) {
      const para = paragraphs.items[idx];
      if (!para) {
        throw new Error(
          `Document changed while staging: paragraph index ${idx} no longer exists ` +
          `(document has ${paragraphs.items.length} paragraphs). Re-run the instruction.`
        );
      }
      para.load('text');
      boundaryParagraphs.set(idx, para);
    }
    const boundaryStates = new Map();
    const boundaryEntries = [...boundaryParagraphs];
    for (let start = 0; start < boundaryEntries.length; start += NATIVE_READ_BATCH_SIZE) {
      const batch = boundaryEntries.slice(start, start + NATIVE_READ_BATCH_SIZE);
      const reads = batch.map(([, para]) => queueRevisionRead(para, { paragraph: true, singleParagraphContent: true }));
      await context.sync();
      batch.forEach(([index, para], i) => boundaryStates.set(index, resolveRevisionRead(para, reads[i], { paragraph: true, singleParagraphContent: true })));
    }

    for (let i = 0; i < chunks.length; i++) {
      const chunk = chunks[i];
      const startPara = paragraphs.items[chunk.startIndex];
      const endPara = paragraphs.items[chunk.endIndex];

      if (Array.isArray(chunk.paragraphs) && chunk.paragraphs.length > 0) {
        const expectedStart = (chunk.paragraphs[0].text || '').trim();
        const expectedEnd = (chunk.paragraphs[chunk.paragraphs.length - 1].text || '').trim();
        const actualStart = (boundaryStates.get(chunk.startIndex).text || '').trim();
        const actualEnd = (boundaryStates.get(chunk.endIndex).text || '').trim();
        if (actualStart !== expectedStart || actualEnd !== expectedEnd) {
          throw new Error(
            'Document changed while staging: chunk boundary paragraphs no longer match ' +
            `(expected "${expectedStart.slice(0, 40)}…", found "${actualStart.slice(0, 40)}…"). ` +
            'Re-run the instruction.'
          );
        }
      }

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
 * Deletes leftover hidden chunk bookmarks (`_wdp*`) from a previous
 * session. Chunk bookmarks are normally removed by the in-memory apply/
 * discard closures of a staged run; when the taskpane reloads mid-run those
 * closures are gone and the bookmarks would leak into the document forever
 * (restored proposals render as static cards that can never clean up).
 * Runs once at startup, before any new run can create its own bookmarks.
 *
 * @param {function} [log] - Logging callback (message, type)
 * @returns {Promise<number>} Number of orphaned bookmarks removed
 */
export async function reapOrphanChunkBookmarks(log = () => {}) {
  let removed = 0;
  try {
    await Word.run(async (context) => {
      const wholeDoc = context.document.body.getRange('Whole');
      const namesResult = wholeDoc.getBookmarks(true /* includeHidden */);
      await context.sync();
      const names = namesResult.value || [];
      const orphans = names.filter((name) => typeof name === 'string' && name.startsWith('_wdp'));
      for (const name of orphans) {
        context.document.deleteBookmark(name);
      }
      await context.sync();
      removed = orphans.length;
    });
    if (removed > 0) {
      log(`Cleaned up ${removed} leftover bookmark(s) from a previous session.`, 'info');
    }
  } catch (err) {
    // Best-effort hygiene: hosts without Range.getBookmarks or with a
    // protected document just skip reaping — never block startup.
    log(`Orphan bookmark cleanup skipped: ${err.message}`, 'warning');
  }
  return removed;
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

/** Read a failed section's current source for an explicit recovery action.
 * The bookmark remains the scope; never search the whole document for a
 * replacement. Changed paragraph boundaries require a fresh document proposal.
 */
export async function readChunkSnapshot(chunk, bookmarkName) {
  if (!bookmarkName) throw new RevisionSafetyError('The section bookmark is missing. Draft a fresh document proposal.');
  return Word.run(async (context) => {
    const range = context.document.getBookmarkRangeOrNullObject(bookmarkName);
    range.load('isNullObject');
    await context.sync();
    if (range.isNullObject || !range.paragraphs) {
      throw new RevisionSafetyError('The section can no longer be located. Draft a fresh document proposal.');
    }
    range.paragraphs.load('items');
    await context.sync();
    const { revisionStates, inTable } = await _readParagraphStates(context, range.paragraphs.items);
    const visible = revisionStates.map((state, index) => ({ state, index })).filter(({ state }) => state.text.trim());
    if (!visible.length || visible.length !== chunk.paragraphs?.length) {
      throw new RevisionSafetyError('The section paragraph boundaries changed. Draft a fresh document proposal to establish its scope.');
    }
    const paragraphs = visible.map(({ state, index }, i) => ({
      ...chunk.paragraphs[i], text: state.text, revisionFingerprint: state.fingerprint,
      hasRevisions: state.hasRevisions, protectedSpans: state.protectedSpans || [],
      tokenEstimate: Math.ceil(state.text.length / 4), inTable: inTable[index],
    }));
    return { ...chunk, paragraphs, tokenCount: Math.ceil(paragraphs.reduce((n, p) => n + p.text.length, 0) / 4) };
  });
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
 * The paragraph reads this performs (items, texts, table membership) are
 * returned alongside the range as `preloaded`, so the paragraph-level
 * amendment can reuse them instead of re-reading the same data (the
 * re-anchor + read phase costs two syncs for the whole chunk instead of
 * five).
 *
 * @param {Word.RequestContext} context
 * @param {Word.Range} range - The bookmarked chunk range
 * @param {string[]} storedTexts - Paragraph texts captured at staging time
 * @param {function} log
 * @returns {Promise<{range: Word.Range, preloaded: {paraItems: Word.Paragraph[], inTable: boolean[], revisionStates?: any[]} | null} | null>}
 *   The (possibly narrowed) range with its preloaded paragraph reads, or
 *   null when the stored sequence is no longer contiguously locatable.
 *   preloaded is null when the range exposes no paragraphs collection
 *   (test mocks, exotic bookmarks) or has no paragraphs at all.
 * @private
 */
async function _reanchorChunkRange(context, range, storedTexts, log, fingerprints = []) {
  // Ranges without a paragraphs collection (test mocks, exotic bookmarks)
  // cannot be re-anchored; use them as-is.
  if (!range.paragraphs) return { range, preloaded: null };

  const rangeParagraphs = range.paragraphs;
  rangeParagraphs.load('items');
  await context.sync();

  const paraItems = rangeParagraphs.items;
  if (paraItems.length === 0) return { range, preloaded: null };

  // Paragraph texts (for the drift window match) and table membership (for
  // the amendment's table guards) in a single batched read.
  const { inTable, revisionStates } = await _readParagraphStates(context, paraItems);

  // Compare on the non-empty paragraphs only: blank spacer paragraphs are
  // absent from the staged sequence but present in the live range.
  const nonEmptyIdx = [];
  const currentTexts = [];
  paraItems.forEach((para, i) => {
    const text = revisionStates[i].text;
    if (text.trim() !== '') {
      nonEmptyIdx.push(i);
      currentTexts.push(text);
    }
  });

  const window = _findAnchorWindow(currentTexts, storedTexts);
  if (!window) return null;
  if (fingerprints.some((fingerprint, i) => fingerprint
      && revisionStates[nonEmptyIdx[window.start + i]]?.fingerprint !== fingerprint)) return null;

  const startItemIdx = nonEmptyIdx[window.start];
  const endItemIdx = nonEmptyIdx[window.end - 1];
  if (startItemIdx === 0 && endItemIdx === paraItems.length - 1) {
    return { range, preloaded: { paraItems, inTable, revisionStates } }; // No drift: original content still spans the whole range
  }

  log(
    `Range drifted since staging; narrowed to paragraphs ${startItemIdx + 1}-${endItemIdx + 1} of ${paraItems.length}`,
    'info'
  );
  const startRange = paraItems[startItemIdx].getRange('Start');
  const endRange = paraItems[endItemIdx].getRange('End');
  const narrowed = startRange.expandTo(endRange);
  // The narrowed range's own paragraph list and its text (for the range-level
  // fallback) in one batched read, then the items' texts + tables in another.
  narrowed.load('text');
  if (!narrowed.paragraphs) {
    // No paragraph collection on the narrowed range (exotic mocks): the
    // paragraph-level amendment reloads for itself and its failure keeps
    // flowing into the range-level fallback.
    await context.sync();
    return { range: narrowed, preloaded: null };
  }
  narrowed.paragraphs.load('items');
  await context.sync();

  const narrowedItems = narrowed.paragraphs.items;
  const narrowedState = await _readParagraphStates(context, narrowedItems);

  return {
    range: narrowed,
    preloaded: {
      paraItems: narrowedItems,
      ...narrowedState,
    },
  };
}

/**
 * Number of paragraphs in a chunk's working range, or null when the range
 * exposes no paragraph collection (exotic bookmarks, test mocks).
 *
 * @private
 */
async function _rangeParagraphCount(context, range, preloaded) {
  if (preloaded) return preloaded.paraItems.length;
  if (!range.paragraphs) return null;
  const paragraphs = range.paragraphs;
  paragraphs.load('items');
  await context.sync();
  return paragraphs.items.length;
}

/** @private */
function _multiParagraphFallbackError(paraErr) {
  return new RevisionSafetyError(`paragraph-level edit failed (${paraErr.message}); a multi-paragraph section `
    + 'cannot fall back to a whole-range rewrite safely. No text was written; draft a new proposal.');
}

/**
 * Applies all chunk results to the document.
 * Amendments applied in reverse chunk order as tracked changes.
 * Uses paragraph-level strategy to preserve formatting; a single-paragraph
 * range falls back to range-level diff strategies if paragraph-level
 * operations fail before writing. Multi-paragraph ranges fail closed.
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
 * @param {Map<string, string[]>} [options.chunkRevisionFingerprints] - revision/formatting baselines by chunk
 * @param {AbortSignal} [options.signal] - Cooperative pause: when aborted, the
 *   loop finishes the in-flight chunk then stops before the next, leaving the
 *   remaining chunks' bookmarks intact so the caller can resume later
 * @param {function(string, {applied: boolean, noChange?: boolean, error?: boolean, skipped?: boolean, partial?: boolean, uncertain?: boolean}): void} [options.onChunkApplied] -
 *   Called after each chunk's apply attempt with its chunkId + outcome, so the
 *   UI can mark individual items applied as they land
 * @param {function(object): void} [options.onProgress] - Current section and paragraph verification stage
 * @returns {Promise<{amendmentsApplied: number, commentsInserted: number, noChangeCount: number,
 *   errors: string[], appliedChunkIds: string[], attemptedChunkIds: string[], failedChunkIds: string[],
 *   partialChunkIds: string[], uncertainChunkIds: string[], appliedParagraphs: number, interrupted: boolean,
 *   halted: boolean, skippedChunkIds: string[]}>}
 */
export async function applyChunkResults(results, bookmarkMap, options) {
  const {
    trackChangesEnabled = true,
    lineDiffEnabled = false,
    log = () => {},
    chunkOriginals = null,
    chunkRevisionFingerprints = null,
    signal = null,
    onChunkApplied = /** @type {function(string, object): void} */ (() => {}),
    onProgress = /** @type {function(object): void} */ (() => {}),
  } = options;

  let amendmentsApplied = 0;
  let commentsInserted = 0;
  let noChangeCount = 0;
  const errors = [];
  const appliedChunkIds = [];
  const attemptedChunkIds = [];
  const failedChunkIds = [];
  const partialChunkIds = [];
  const uncertainChunkIds = [];
  const completedChunkIds = new Set();
  let appliedParagraphs = 0;
  let interrupted = false;
  const halted = false; // Kept for callers consuming older application results.
  const skippedChunkIds = [];

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
    attemptedChunkIds.push(result.chunkId);
    if (!bookmarkName) {
      errors.push(`Chunk ${result.chunkId}: no bookmark found`);
      failedChunkIds.push(result.chunkId);
      onChunkApplied(result.chunkId, { applied: false, skipped: true });
      continue;
    }

    const progress = { report: () => {
      // Progress display failures must never interrupt Word mutations.
      try { onProgress({ chunkId: result.chunkId, stage: progress.stage, appliedParagraphs: progress.appliedParagraphs }); } catch (_err) { /* UI only */ }
    }, appliedParagraphs: 0, mutationAttempted: false, pendingVerification: false, pendingStructuralWrites: false,
      stage: 'bookmark and source read' };
    progress.report();
    try {
      let chunkApplied = false;
      await Word.run(async (context) => {
        const range = context.document.getBookmarkRangeOrNullObject(bookmarkName);
        range.load('isNullObject,text');
        if (Word.ChangeTrackingMode && typeof context.document.load === 'function') context.document.load('changeTrackingMode');
        await context.sync();

        if (range.isNullObject) {
          throw new RevisionSafetyError('bookmark range lost');
        }

        // Re-anchor: the bookmark may have absorbed paragraphs inserted after
        // staging (e.g. a title from another proposal card). Narrow the
        // working range to the stored original paragraphs so the amendment
        // never deletes absorbed content.
        const storedTexts = _storedParagraphTexts(result, chunkOriginals);
        if (uncertainChunkIds.length && !storedTexts) {
          throw new RevisionSafetyError('Cannot independently validate this section after an unverified write: staged source is missing.');
        }
        let workRange = range;
        let preloaded = null;
        if (storedTexts) {
          let anchored;
          try {
            const fingerprints = chunkRevisionFingerprints?.get(result.chunkId)
              || result.chunk?.paragraphs?.map((p) => p.revisionFingerprint) || [];
            anchored = await _reanchorChunkRange(context, range, storedTexts, log, fingerprints);
          } catch (anchorErr) {
            // A failed drift check means we cannot tell whether the range
            // absorbed new content; falling back to the raw bookmark range
            // would risk deleting absorbed paragraphs. Skip instead.
            log(`Chunk ${result.chunkId}: re-anchor check failed (${anchorErr.message}), skipping to avoid deleting absorbed content`, 'warning');
            throw new RevisionSafetyError(`re-anchor check failed (${anchorErr.message}); amendment skipped`);
          }
          if (anchored === null) {
            log(`Chunk ${result.chunkId}: original content not found contiguously in staged range, skipping to avoid deleting absorbed content`, 'warning');
            throw new RevisionSafetyError('original content no longer matches the staged range (edited since staging?); amendment skipped');
          }
          workRange = anchored.range;
          preloaded = anchored.preloaded;
          if (uncertainChunkIds.length && !preloaded) {
            throw new RevisionSafetyError('Cannot independently validate this section after an unverified write: paragraph evidence is unavailable.');
          }
        }

        // Try paragraph-level strategy first (preserves formatting)
        let applied;
        try {
          applied = await _applyParagraphLevelAmendment(
            context, workRange, result.amendment,
            trackChangesEnabled, lineDiffEnabled, log, preloaded,
            { whitespaceOnly: !!result.whitespaceOnly, expectedTexts: storedTexts, progress }
          ) !== false;
        } catch (paraErr) {
          if (progress.mutationAttempted || paraErr.mutationAttempted || result.whitespaceOnly
              || paraErr instanceof TruncatedOutputError || paraErr instanceof RevisionSafetyError) {
            // The amendment text itself is truncated. A range-level strategy
            // would just write the same truncated text more crudely (and
            // without per-paragraph formatting protection) — escalate to the
            // chunk-level error path so the chunk is reported as failed and
            // the document is left untouched.
            throw paraErr;
          }
          // A range-level strategy diffs the range's whole text. Across
          // several paragraphs that text includes blank spacer paragraphs the
          // amendment never contained (the parser skips them), so the diff
          // would delete them and still pass read-back, which compares
          // against the amendment. Only a single-paragraph range may fall
          // back; anything larger fails closed like every other unsafe path.
          const paragraphCount = await _rangeParagraphCount(context, workRange, preloaded);
          if (paragraphCount > 1) {
            throw _multiParagraphFallbackError(paraErr);
          }
          log(`Chunk ${result.chunkId}: paragraph-level strategy failed (${paraErr.message}), falling back to range-level`, 'warning');

          // Fallback to range-level diff strategies
          if (trackChangesEnabled && !hostSupportsTrackedChanges()) {
            throw new RevisionSafetyError('This Word host cannot enable tracked changes.');
          }
          const previousMode = context.document.changeTrackingMode ?? Word.ChangeTrackingMode?.off;
          try {
            if (Word.ChangeTrackingMode) {
              context.document.changeTrackingMode = trackChangesEnabled
                ? Word.ChangeTrackingMode.trackAll
                : Word.ChangeTrackingMode.off;
              try { await context.sync(); }
              catch (error) { throw new RevisionSafetyError(`Could not set Word's revision mode: ${error.message}`); }
            }

            // Normalize line endings for consistent diffing
            const revisionRead = queueRevisionRead(workRange);
            if (revisionRead) await context.sync();
            const originalText = _normalizeLineEndings(resolveRevisionRead(workRange, revisionRead).text);
            // Without a paragraph collection, a paragraph break in the text is
            // the only evidence of a multi-paragraph range.
            if (!paragraphCount && originalText.includes('\n')) {
              throw _multiParagraphFallbackError(paraErr);
            }
            const normalizedAmendment = _normalizeLineEndings(result.amendment);

            const strategyOptions = { trackChanges: false };
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
            progress.mutationAttempted = true;
            applied = true;
          } finally {
            if (Word.ChangeTrackingMode) {
              context.document.changeTrackingMode = previousMode;
              await context.sync();
            }
          }
        }

        if (!applied) {
          noChangeCount++;
        } else {
          chunkApplied = true;
        }
      });

      if (chunkApplied) {
        amendmentsApplied++;
        appliedChunkIds.push(result.chunkId);
        log(`Chunk ${result.chunkId}: amendment applied`, 'info');
      } else {
        log(`Chunk ${result.chunkId}: no changes needed`, 'info');
      }
      completedChunkIds.add(result.chunkId);
      onChunkApplied(result.chunkId, { applied: chunkApplied, noChange: !chunkApplied });
    } catch (err) {
      const partial = progress.appliedParagraphs > 0;
      const uncertain = !!(err.mutationAttempted || progress.pendingVerification || progress.pendingStructuralWrites
        || (progress.mutationAttempted && !partial));
      failedChunkIds.push(result.chunkId);
      if (partial) partialChunkIds.push(result.chunkId);
      if (uncertain) uncertainChunkIds.push(result.chunkId);
      const detail = wordErrorDetail(err);
      errors.push(`Chunk ${result.chunkId}: ${detail} (${progress.stage})`);
      log(`Chunk ${result.chunkId}: amendment failed -- ${detail} (${progress.stage})`
        + (partial ? `; ${progress.appliedParagraphs} paragraph edit(s) already applied` : '')
        + (uncertain ? '; writes could not be verified; inspect Word before drafting fresh edits' : ''), 'error');
      onChunkApplied(result.chunkId, { applied: false, error: true, partial, uncertain });
      if (uncertain) log(`Chunk ${result.chunkId}: left for review; other sections will be checked independently before application.`, 'warning');
    }
    appliedParagraphs += progress.appliedParagraphs;

    // Yield to event loop between chunks to prevent UI freeze
    await _yieldToEventLoop();
  }

  // Phase 2: Comments in document order (after all amendments).
  // Failed sections retain their bookmarks for recovery, but must not receive
  // comments. After an uncertain write, only independently verified sections
  // are eligible; comment-only ranges have not passed that validation.
  if (!interrupted) {
    const allCommentResults = results
      .filter((r) => r.status === 'fulfilled' && r.comment)
      .filter((r) => bookmarkMap.get(r.chunkId) && !failedChunkIds.includes(r.chunkId))
      .filter((r) => !uncertainChunkIds.length || completedChunkIds.has(r.chunkId));
    let commentTrackingMode;
    let restoreCommentTracking = false;
    // Ensure tracked changes are off before inserting comments.
    // If any amendment fallback path left ChangeTrackingMode.trackAll enabled,
    // comment insertion on ranges containing tracked changes can fail with AccessDenied.
    if (fulfilledWithAmendments.length > 0 && allCommentResults.length > 0) {
      try {
        await Word.run(async (context) => {
          if (Word.ChangeTrackingMode) {
            if (typeof context.document.load === 'function') {
              context.document.load('changeTrackingMode');
              await context.sync();
            }
            commentTrackingMode = context.document.changeTrackingMode ?? Word.ChangeTrackingMode.off;
            restoreCommentTracking = true;
            context.document.changeTrackingMode = Word.ChangeTrackingMode.off;
            await context.sync();
          }
        });
      } catch (_err) {
        // Best-effort -- continue with comment insertion even if this fails
      }
    }

    // One Word.run for the whole comment pass: bookmark lookups batched into
    // a single sync, then one insert+sync per comment (the per-comment sync
    // keeps the old per-comment-run error isolation). A failed insert leaves
    // the run's remaining queue unreliable, so the loop escapes and re-drives
    // the remainder in a fresh run; the queue is guaranteed to shrink every
    // iteration.
    let commentQueue = allCommentResults;
    while (commentQueue.length > 0) {
      let consumed = 0;
      try {
        await Word.run(async (context) => {
          const pending = commentQueue.map((result) => {
            const range = context.document.getBookmarkRangeOrNullObject(bookmarkMap.get(result.chunkId));
            range.load('isNullObject');
            return { result, range };
          });
          await context.sync();

          for (let i = 0; i < pending.length; i++) {
            const { result, range } = pending[i];
            if (range.isNullObject) {
              errors.push(`Chunk ${result.chunkId}: bookmark range lost for comment`);
              consumed = i + 1;
              continue;
            }
            try {
              range.insertComment(result.comment);
              await context.sync();
              commentsInserted++;
              log(`Chunk ${result.chunkId}: comment inserted`, 'info');
              consumed = i + 1;
            } catch (err) {
              errors.push(`Chunk ${result.chunkId}: comment failed -- ${err.message || String(err)}`);
              log(`Chunk ${result.chunkId}: comment failed -- ${err.message}`, 'error');
              consumed = i + 1;
              throw err; // poisoned context: escape, retry the rest in a fresh run
            }

            // Yield to event loop between comments to prevent UI freeze and
            // pace the rapid-fire host writes
            await _yieldToEventLoop();
          }
        });
      } catch (err) {
        // Run-level failure (lookup sync, host error): fail the first queue
        // entry so the loop always makes progress, matching the old
        // per-comment-run behavior for the rest.
        if (consumed === 0) {
          const dropped = commentQueue[0];
          errors.push(`Chunk ${dropped.chunkId}: comment failed -- ${err.message || String(err)}`);
          log(`Chunk ${dropped.chunkId}: comment failed -- ${err.message}`, 'error');
          consumed = 1;
        }
      }
      commentQueue = commentQueue.slice(consumed);
    }
    if (restoreCommentTracking) {
      try {
        await Word.run(async (context) => {
          context.document.changeTrackingMode = commentTrackingMode;
          await context.sync();
        });
      } catch (error) {
        errors.push(`Could not restore Word's revision mode after comments: ${error.message}`);
      }
    }
  }

  return { amendmentsApplied, commentsInserted, noChangeCount, errors, appliedChunkIds, attemptedChunkIds,
    failedChunkIds, partialChunkIds, uncertainChunkIds, appliedParagraphs, interrupted, halted, skippedChunkIds };
}

/**
 * Removes chunk bookmarks from the document.
 * Tolerates individual bookmark deletion failures.
 *
 * @param {Map<string, string>} bookmarkMap - chunkId -> bookmarkName
 * @param {object} [options]
 * @param {Set<string>} [options.keep] - Bookmark names to leave in place.
 *   The apply path keeps FAILED chunks' bookmarks: the retry link re-drives
 *   exactly those chunks and needs their staged ranges to still exist.
 * @returns {Promise<void>}
 */
export async function cleanupBookmarks(bookmarkMap, { keep = null } = {}) {
  await Word.run(async (context) => {
    for (const bookmarkName of bookmarkMap.values()) {
      if (keep && keep.has(bookmarkName)) continue;
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
