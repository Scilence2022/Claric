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
 *   - Accepts options.trackChanges (default true); when false the strategy
 *     does not touch the document's changeTrackingMode (caller owns it).
 *   - The fallback reset runs with tracking OFF so it does not show up as a
 *     spurious whole-range revision pair.
 *   - Typed activity-log messages instead of DEBUG/emoji lines.
 *
 * @module lib/word-diff/token-map
 */

import DiffMatchPatch from './diff-wordmode.js';
import { applySentenceDiffStrategy } from './sentence-diff.js';

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
    const trackChanges = options.trackChanges !== false;
    log('Running word-level diff...', 'info');

    let insertions = 0;
    let deletions = 0;
    // True only when THIS strategy enabled tracking (and must restore it).
    let trackingEnabled = false;

    try {
        // Run diff_wordMode
        const dmp = new DiffMatchPatch();
        const diffs = dmp.diff_wordMode(originalText, newText);

        // --- Build Refined Token Map (Batched) ---

        // 1. Get Coarse Ranges
        const coarseRanges = range.getTextRanges([' '], false);
        coarseRanges.load('items/text');
        await context.sync();

        const fineTokens = [];
        const dmpRegex = /(\w+|[^\w\s]+|\s+)/g;
        const searchProxies = [];

        // 2. Queue all searches. Regex tokens partition each coarse text, so
        // the k-th regex occurrence of a token corresponds to the k-th search
        // match in document order — record the occurrence index per coarse
        // range so repeated tokens do not all map onto the first match.
        for (let i = 0; i < coarseRanges.items.length; i++) {
            const coarseRange = coarseRanges.items[i];
            const coarseText = coarseRange.text;
            const seen = Object.create(null);
            let match;
            dmpRegex.lastIndex = 0;

            while ((match = dmpRegex.exec(coarseText)) !== null) {
                const tokenText = match[0];
                if (tokenText.length === 0) continue;

                const occurrence = seen[tokenText] || 0;
                seen[tokenText] = occurrence + 1;

                // Queue search
                const searchResults = coarseRange.search(tokenText, { matchCase: true });
                searchResults.load('items');
                searchProxies.push({
                    text: tokenText,
                    occurrence,
                    results: searchResults,
                    coarseText: coarseText,
                });
            }
        }

        // SYNC: Execute all searches
        await context.sync();

        // 3. Process results
        for (const proxy of searchProxies) {
            if (proxy.results.items.length > proxy.occurrence) {
                fineTokens.push({
                    text: proxy.text,
                    range: proxy.results.items[proxy.occurrence],
                });
            } else {
                log(`Could not map token "${proxy.text}" (occurrence ${proxy.occurrence + 1}) inside "${proxy.coarseText}"`, 'warning');
                throw new Error(`Token mapping failed for "${proxy.text}"`);
            }
        }

        fineTokens.forEach((t, i) => t.index = i);

        // --- Pass 1: Identify Deletions ---
        const deleteTargets = [];
        let tokenIndex = 0;

        for (const [op, chunk] of diffs) {
            if (op === 0) { // EQUAL
                const chunkTokens = chunk.match(/(\w+|[^\w\s]+|\s+)/g) || [];
                tokenIndex += chunkTokens.length;
            } else if (op === -1) { // DELETE
                const chunkTokens = chunk.match(/(\w+|[^\w\s]+|\s+)/g) || [];
                const count = chunkTokens.length;
                for (let i = 0; i < count; i++) {
                    if (tokenIndex < fineTokens.length) {
                        deleteTargets.push(fineTokens[tokenIndex]);
                        tokenIndex++;
                    }
                }
            }
        }

        // --- Pass 2: Identify Insertions ---
        const deletedIndices = new Set(deleteTargets.map((t) => t.index));
        const tokensAfterDeletes = fineTokens.filter((t) => !deletedIndices.has(t.index));

        const insertOps = [];
        let currentTokenIdx = 0;
        let lastAnchorRange = null;

        for (const [op, chunk] of diffs) {
            if (op === 0) { // EQUAL
                let textToConsume = chunk;
                while (textToConsume.length > 0 && currentTokenIdx < tokensAfterDeletes.length) {
                    const token = tokensAfterDeletes[currentTokenIdx];
                    const tokenText = token.text;

                    if (textToConsume.startsWith(tokenText)) {
                        textToConsume = textToConsume.slice(tokenText.length);
                        lastAnchorRange = token.range;
                        currentTokenIdx++;
                    } else {
                        log(`Token mismatch: expected "${textToConsume.slice(0, 10)}..." but found "${tokenText}"`, 'warning');
                        throw new Error('Map lookup failed: Token mismatch.');
                    }
                }
            } else if (op === 1) { // INSERT
                if (lastAnchorRange) {
                    insertOps.push({
                        anchor: lastAnchorRange,
                        location: Word.InsertLocation.after,
                        text: chunk,
                    });
                } else {
                    // Insert at start of range
                    insertOps.push({
                        anchor: range.getRange(Word.RangeLocation.start),
                        location: Word.InsertLocation.before,
                        text: chunk,
                    });
                }
            }
        }

        // --- Execution Phase ---
        if (trackChanges && Word.ChangeTrackingMode) {
            context.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
            await context.sync();
            trackingEnabled = true;
        }

        // Apply Deletes (Reverse order)
        deleteTargets.sort((a, b) => b.index - a.index);
        deleteTargets.forEach((token) => token.range.delete());
        deletions = deleteTargets.length;

        // Apply Inserts
        insertOps.forEach((op) => op.anchor.insertText(op.text, op.location));
        insertions = insertOps.length;

        // Commit all edits
        await context.sync();
        log(`Word-level diff applied (${insertions} insertions, ${deletions} deletions)`, 'info');

        return { strategy: 'token', insertions, deletions };
    } catch (e) {
        log(`Word-level strategy failed (${e.message}); falling back to sentence diff`, 'warning');

        // Reset the range with tracking OFF so the reset itself does not show
        // up as a spurious whole-range revision pair, then sentence-diff
        // (which manages tracking itself via the same options).
        if (trackingEnabled) {
            try {
                context.document.changeTrackingMode = Word.ChangeTrackingMode.off;
                await context.sync();
                trackingEnabled = false;
            } catch (_resetErr) {
                // Best-effort; sentence diff re-asserts tracking itself.
            }
        }
        range.insertText(originalText, Word.InsertLocation.replace);
        await context.sync();

        return applySentenceDiffStrategy(context, range, originalText, newText, log, options);
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
