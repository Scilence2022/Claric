/**
 * Token Map Strategy for word-level diff application
 *
 * This strategy maps individual words/tokens 1:1 to preserve formatting
 * and apply granular tracked changes.
 *
 * Vendored from office-word-diff (Apache-2.0) — see LICENSE and NOTICE in
 * this directory.
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
 * @returns {Promise<{strategy: string, insertions: number, deletions: number}>}
 * @throws {Error} If all strategies fail
 */
export async function applyTokenMapStrategy(context, range, originalText, newText, log) {
    log("DEBUG: Running Token Map Strategy...");

    let insertions = 0;
    let deletions = 0;

    try {
        // Run diff_wordMode
        const dmp = new DiffMatchPatch();
        const diffs = dmp.diff_wordMode(originalText, newText);

        // --- Build Refined Token Map (Batched) ---
        log("DEBUG: Building Refined Token Map (Batched)...");

        // 1. Get Coarse Ranges
        const coarseRanges = range.getTextRanges([" "], false);
        coarseRanges.load("items/text");
        await context.sync();

        const fineTokens = [];
        const dmpRegex = /(\w+|[^\w\s]+|\s+)/g;
        const searchProxies = [];

        // 2. Queue all searches
        for (let i = 0; i < coarseRanges.items.length; i++) {
            const coarseRange = coarseRanges.items[i];
            const coarseText = coarseRange.text;
            let match;
            dmpRegex.lastIndex = 0;

            while ((match = dmpRegex.exec(coarseText)) !== null) {
                const tokenText = match[0];
                if (tokenText.length === 0) continue;

                // Queue search
                const searchResults = coarseRange.search(tokenText, { matchCase: true });
                searchResults.load("items");
                searchProxies.push({
                    text: tokenText,
                    results: searchResults,
                    coarseText: coarseText
                });
            }
        }

        // SYNC: Execute all searches
        await context.sync();

        // 3. Process results
        for (const proxy of searchProxies) {
            if (proxy.results.items.length > 0) {
                fineTokens.push({
                    text: proxy.text,
                    range: proxy.results.items[0]
                });
            } else {
                log(`⚠️ Could not map fine token "${proxy.text}" inside "${proxy.coarseText}"`);
                throw new Error(`Token mapping failed for "${proxy.text}"`);
            }
        }

        fineTokens.forEach((t, i) => t.index = i);
        log(`DEBUG: Refined Token Map built with ${fineTokens.length} entries.`);

        // --- Pass 1: Identify Deletions ---
        log("DEBUG: Pass 1 - Collecting delete targets");
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
        log("DEBUG: Pass 2 - Collecting insert operations");
        const deletedIndices = new Set(deleteTargets.map(t => t.index));
        const tokensAfterDeletes = fineTokens.filter(t => !deletedIndices.has(t.index));

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
                        log(`⚠️ Sync warning: Expected "${textToConsume.slice(0, 10)}..." but found token "${tokenText}"`);
                        throw new Error("Map lookup failed: Token mismatch.");
                    }
                }
            } else if (op === 1) { // INSERT
                if (lastAnchorRange) {
                    insertOps.push({
                        anchor: lastAnchorRange,
                        location: Word.InsertLocation.after,
                        text: chunk
                    });
                } else {
                    // Insert at start of range
                    insertOps.push({
                        anchor: range.getRange(Word.RangeLocation.start),
                        location: Word.InsertLocation.before,
                        text: chunk
                    });
                }
            }
        }

        // --- Execution Phase (Atomic-ish) ---
        log("DEBUG: Executing queued operations...");

        // 1. Enable Track Changes
        if (Word.ChangeTrackingMode) {
            context.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
        }

        // 2. Apply Deletes (Reverse order)
        deleteTargets.sort((a, b) => b.index - a.index);
        deleteTargets.forEach(token => token.range.delete());
        deletions = deleteTargets.length;

        // 3. Apply Inserts
        insertOps.forEach(op => op.anchor.insertText(op.text, op.location));
        insertions = insertOps.length;

        // 4. Disable Track Changes
        if (Word.ChangeTrackingMode) {
            context.document.changeTrackingMode = Word.ChangeTrackingMode.off;
        }

        // SYNC: Commit all edits
        await context.sync();
        log("✅ SUCCESS: Word-level diff applied.");

        return { strategy: 'token', insertions, deletions };

    } catch (e) {
        log(`❌ Word-level strategy failed: ${e.message}`);
        log("⚠️ Initiating Clean Fallback to Sentence Diff Strategy...");

        // Fallback Logic:
        // 1. Reset the range to the original text to ensure a clean state
        range.insertText(originalText, Word.InsertLocation.replace);
        await context.sync();
        log("DEBUG: Range reset to original text for fallback.");

        return await applySentenceDiffStrategy(context, range, originalText, newText, log);
    }
}
