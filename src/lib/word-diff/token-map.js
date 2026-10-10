/**
 * Token Map Strategy for word-level diff application
 *
 * This strategy maps individual words/tokens 1:1 to preserve formatting
 * and apply granular tracked changes.
 *
 * Vendored from office-word-diff (Apache-2.0) — see LICENSE and NOTICE in
 * this directory.
 *
 * Local modifications:
 *   - Occurrence-aware token mapping: upstream searched every fine token
 *     with coarseRange.search() and always took the FIRST match, so a token
 *     repeated within one coarse range (e.g. "." in "a.b.c") mapped onto the
 *     same range twice. We now take the Nth match for the Nth occurrence.
 *   - Only the tokens the edits actually touch are located. The token
 *     sequence is recomputed client-side with the exact regex diff_wordMode
 *     tokenizes with, and each needed token's match is picked by occurrence
 *     index — simulating Word's greedy substring scan (so embedded matches
 *     like the "the" inside "other" never shift the mapping) — with no
 *     coarse-range reads (range.getTextRanges) and no search per document
 *     token (hundreds of host-side searches per paragraph in the upstream
 *     design). Deleted tokens plus one anchor per insertion are located in
 *     ONE batched search pass.
 *   - Accepts options.trackChanges (default true); when false the strategy
 *     does not touch the document's changeTrackingMode (caller owns it).
 *   - Read-only mapping failures try another locator without resetting the
 *     native runs. Failures after writes stop without another fallback.
 *   - Typed activity-log messages instead of DEBUG/emoji lines.
 *
 * @module lib/word-diff/token-map
 */

import DiffMatchPatch from './diff-wordmode.js';
import { tryRevisionDiff } from './revision-diff.js';
import { applySentenceDiffStrategy } from './sentence-diff.js';
import { _occurrenceIndex, applyCharDiffStrategy } from './char-diff.js';
import { RevisionSafetyError, MutationSafetyError, verifyMutationText, enableTrackedWrites } from '../word-revisions.js';

/** Word/punctuation/whitespace tokenization — MUST match the regex
 *  diff_wordMode tokenizes with (diff-wordmode.js), since the diff walk
 *  consumes token counts against this sequence. */
const TOKEN_RE = /(\w+|[^\w\s]+|\s+)/g;

/**
 * Applies the "Refined Token Map" strategy to update a range with new text.
 *
 * This strategy attempts to map words 1:1 to preserve formatting and track changes granularly.
 * If it fails (e.g., due to complex structural changes), it falls back to the Sentence Diff strategy.
 *
 * @param {Word.RequestContext} context - The Word request context
 * @param {Word.Range} range - The target range to update
 * @param {string} originalText - The original text of the range (for fallback)
 * @param {string} newText - The new text to apply
 * @param {function} log - Callback for logging messages
 * @param {object} [options]
 * @param {boolean} [options.trackChanges=true] - When false the strategy does
 *   NOT touch the document's changeTrackingMode; the caller owns it.
 * @returns {Promise<{strategy: string, insertions: number, deletions: number}>}
 * @throws {Error} If all strategies fail
 */
export async function applyTokenMapStrategy(context, range, originalText, newText, log, options = {}) {
    const revised = await tryRevisionDiff(context, range, originalText, newText, log, options);
    if (revised) return revised;
    const trackChanges = options.trackChanges !== false;
    if (trackChanges && !Word.ChangeTrackingMode) throw new RevisionSafetyError('This Word host cannot enable tracked changes.');
    log('Running word-level diff...', 'info');

    let insertions = 0;
    let deletions = 0;
    // True only when THIS strategy enabled tracking (and must restore it).
    let trackingEnabled = false;
    let mutationAttempted = false;
    const insertedRanges = [];

    try {
        // Run diff_wordMode
        // (prototype extension from diff-wordmode.js — outside the vendored lib's type surface)
        const dmp = /** @type {any} */ (new DiffMatchPatch());
        const diffs = dmp.diff_wordMode(originalText, newText);

        // The diff's chunks are substrings of originalText, so re-tokenizing
        // originalText with the same regex yields exactly the token sequence
        // the diff walk steps through — entirely client-side.
        const tokenTexts = originalText.match(TOKEN_RE) || [];

        // --- Pass 1: Identify Deletions (token indices) ---
        const deleteIndices = [];
        let tokenIndex = 0;

        for (const [op, chunk] of diffs) {
            const chunkTokens = chunk.match(TOKEN_RE) || [];
            if (op === 0) { // EQUAL
                tokenIndex += chunkTokens.length;
            } else if (op === -1) { // DELETE
                for (let i = 0; i < chunkTokens.length; i++) {
                    if (tokenIndex < tokenTexts.length) {
                        deleteIndices.push(tokenIndex);
                        tokenIndex++;
                    }
                }
            }
        }

        // --- Pass 2: Identify Insertions (anchor token indices) ---
        const deletedSet = new Set(deleteIndices);
        const survivors = [];
        for (let i = 0; i < tokenTexts.length; i++) {
            if (!deletedSet.has(i)) survivors.push({ text: tokenTexts[i], index: i });
        }

        /** @type {Array<{anchorIndex: number, text: string}>} anchorIndex
         *  -1 = insert at the very start of the range. */
        const insertOps = [];
        let currentSurvivorIdx = 0;
        let lastAnchorIndex = -1;

        for (const [op, chunk] of diffs) {
            if (op === 0) { // EQUAL
                let textToConsume = chunk;
                while (textToConsume.length > 0 && currentSurvivorIdx < survivors.length) {
                    const token = survivors[currentSurvivorIdx];

                    if (textToConsume.startsWith(token.text)) {
                        textToConsume = textToConsume.slice(token.text.length);
                        lastAnchorIndex = token.index;
                        currentSurvivorIdx++;
                    } else {
                        log(`Token ${token.index + 1}: diff alignment mismatch.`, 'warning');
                        throw new Error('Map lookup failed: Token mismatch.');
                    }
                }
            } else if (op === 1) { // INSERT
                insertOps.push({ anchorIndex: lastAnchorIndex, text: chunk });
            }
        }

        // --- Locate only the tokens the edits touch (one batched sync) ---
        // Word's search (matchWholeWord: false) returns SUBSTRING matches, so
        // the k-th match of "the" includes the one inside "other". The
        // occurrence index must therefore come from simulating that greedy,
        // left-to-right, non-overlapping substring scan up to each token's
        // own position — counting earlier identical TOKENS would map the
        // second standalone "the" onto the "the" inside "other" and delete
        // mid-word. An unresolvable position (rare self-overlap) throws into
        // the sentence-diff fallback below.
        const neededIndices = new Set(deleteIndices);
        for (const op of insertOps) {
            if (op.anchorIndex >= 0) neededIndices.add(op.anchorIndex);
        }

        const tokenStarts = new Array(tokenTexts.length);
        TOKEN_RE.lastIndex = 0;
        for (let i = 0, match; (match = TOKEN_RE.exec(originalText)) !== null; i++) {
            tokenStarts[i] = match.index;
        }

        const occurrenceByIndex = new Map();
        for (const index of neededIndices) {
            const occurrence = _occurrenceIndex(originalText, tokenStarts[index], tokenTexts[index]);
            if (occurrence === null) {
                throw new Error(`Token ${index + 1}: overlapping occurrences prevent mapping.`);
            }
            occurrenceByIndex.set(index, occurrence);
        }

        const searchByToken = new Map();
        for (const index of neededIndices) {
            const text = tokenTexts[index];
            if (!searchByToken.has(text)) {
                const matches = range.search(text, { matchCase: true });
                matches.load('items');
                searchByToken.set(text, matches);
            }
        }

        // SYNC: execute the searches (the single locate round-trip)
        if (searchByToken.size > 0) {
            await context.sync();
        }

        /** @type {Map<number, Word.Range>} token index -> located range */
        const rangeByIndex = new Map();
        for (const index of neededIndices) {
            const text = tokenTexts[index];
            const occurrence = occurrenceByIndex.get(index);
            const matches = searchByToken.get(text);
            if (!matches.items || matches.items.length <= occurrence) {
                throw new Error(`Token mapping failed at offset ${tokenStarts[index]} (occurrence ${occurrence + 1}).`);
            }
            rangeByIndex.set(index, matches.items[occurrence]);
        }

        // --- Execution Phase ---
        if (trackChanges && Word.ChangeTrackingMode) {
            await enableTrackedWrites(context);
            trackingEnabled = true;
        }

        // Apply Deletes (reverse document order). Adjacent tokens (consecutive
        // indices) are coalesced into one spanning delete: a single Word
        // revision + undo entry per contiguous run instead of one per token.
        // Per-token deletions inflate Word's session-scoped undo/revision
        // bookkeeping, which survives add-in close and is a known source of
        // post-apply editing lag (cleared only by closing the document).
        deleteIndices.sort((a, b) => a - b);
        const deleteRuns = [];
        for (const index of deleteIndices) {
            const run = deleteRuns[deleteRuns.length - 1];
            if (run && index === run.last + 1) {
                run.last = index;
            } else {
                deleteRuns.push({ first: index, last: index });
            }
        }
        for (let i = deleteRuns.length - 1; i >= 0; i--) {
            const { first, last } = deleteRuns[i];
            const firstRange = rangeByIndex.get(first);
            mutationAttempted = true;
            (last === first ? firstRange : firstRange.expandTo(rangeByIndex.get(last))).delete();
        }
        deletions = deleteIndices.length;

        // Apply Inserts
        insertOps.forEach((op) => {
            mutationAttempted = true;
            if (op.anchorIndex >= 0) {
                insertedRanges.push(rangeByIndex.get(op.anchorIndex).insertText(op.text, Word.InsertLocation.after));
            } else {
                // Insert at start of range
                insertedRanges.push(range.getRange(Word.RangeLocation.start).insertText(op.text, Word.InsertLocation.before));
            }
        });
        insertions = insertOps.length;

        // Commit all edits
        await context.sync();
        await verifyMutationText(context, range, newText, options, insertedRanges);
        log(`Word-level diff applied (${insertions} insertions, ${deletions} deletions)`, 'info');

        return { strategy: 'token', insertions, deletions };
    } catch (e) {
        if (mutationAttempted || e.mutationAttempted) throw e instanceof MutationSafetyError ? e : new MutationSafetyError(e);
        if (e instanceof RevisionSafetyError) throw e;
        log(`Word-level mapping failed before writing (${e.message}); trying character diff`, 'warning');

        // A mapping failure is read-only. Keep the native runs intact and
        // try another locator; an old-baseline reset would lose formatting
        // and, under caller-owned tracking, introduce fresh revisions.
        if (trackingEnabled) {
            try {
                context.document.changeTrackingMode = Word.ChangeTrackingMode.off;
                await context.sync();
                trackingEnabled = false;
            } catch (_resetErr) {
                // Best-effort; sentence diff re-asserts tracking itself.
            }
        }
        try {
            return await applyCharDiffStrategy(context, range, originalText, newText, log, options);
        } catch (charError) {
            if (charError instanceof RevisionSafetyError || charError.mutationAttempted) throw charError;
            log(`Character mapping failed before writing (${charError.message}); trying sentence diff`, 'warning');
            return await applySentenceDiffStrategy(context, range, originalText, newText, log, options);
        }
    } finally {
        if (trackingEnabled) {
            try {
                context.document.changeTrackingMode = Word.ChangeTrackingMode.off;
                await context.sync();
            } catch (_restoreErr) {
                // Best-effort restore; never mask the primary outcome.
            }
        }
    }
}
