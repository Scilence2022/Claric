/**
 * Block Replace Strategy - Last resort fallback
 *
 * This strategy deletes the entire range content and inserts new text.
 * Used when token and sentence strategies fail.
 *
 * Vendored from office-word-diff (Apache-2.0) — see LICENSE and NOTICE in
 * this directory.
 *
 * @module lib/word-diff/block-replace
 */

/**
 * Applies the "Block Replace" strategy.
 * Deletes the entire range and inserts new text as tracked changes.
 * This is the final fallback when more granular strategies fail.
 *
 * @param {Word.RequestContext} context - The Word request context
 * @param {Word.Range} range - The target range to update
 * @param {string} newText - The new text to apply
 * @param {function} log - Callback for logging messages
 * @returns {Promise<{strategy: string, insertions: number, deletions: number}>}
 */
export async function applyBlockReplaceStrategy(context, range, newText, log) {
    log("DEBUG: Running Block Replace Strategy (Final Fallback)...");

    try {
        // Enable Track Changes
        let trackingEnabled = false;
        try {
            if (Word.ChangeTrackingMode) {
                context.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
                await context.sync();
                trackingEnabled = true;
                log("DEBUG: Track changes enabled");
            }
        } catch (e) {
            log(`⚠️ Could not enable track changes: ${e.message}`);
        }

        // Get the content range
        const contentRange = range.getRange(Word.RangeLocation.content);

        // Delete the content (this will be tracked)
        contentRange.delete();

        // Insert new text after the deleted range
        // Using 'after' ensures it appears as a replacement in track changes
        contentRange.insertText(newText, Word.InsertLocation.after);

        await context.sync();
        log("✅ Block replacement applied successfully.");

        // Disable Track Changes
        if (trackingEnabled) {
            context.document.changeTrackingMode = Word.ChangeTrackingMode.off;
            await context.sync();
        }

        return {
            strategy: 'block',
            insertions: 1,
            deletions: 1
        };

    } catch (e) {
        log(`❌ Block replace strategy failed: ${e.message}`);
        throw new Error(`All diff strategies failed. Final error: ${e.message}`);
    }
}
