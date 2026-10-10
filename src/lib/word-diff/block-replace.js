/**
 * Block Replace Strategy - Last resort fallback
 *
 * This strategy deletes the entire range content and inserts new text.
 * Used when token and sentence strategies fail.
 *
 * Vendored from office-word-diff (Apache-2.0) — see LICENSE and NOTICE in
 * this directory.
 *
 * Local modifications:
 *   - Accepts options.trackChanges (default true); when false the strategy
 *     does not touch the document's changeTrackingMode (caller owns it).
 *   - Tracking state is restored in a finally block, so a mid-edit failure
 *     never leaves trackAll enabled on the document.
 *   - Typed activity-log messages instead of DEBUG/emoji lines.
 *
 * @module lib/word-diff/block-replace
 */
import { RevisionSafetyError, MutationSafetyError, verifyMutationText, loadMutationVerifier, enableTrackedWrites } from '../word-revisions.js';

/**
 * Applies the "Block Replace" strategy.
 * Deletes the entire range and inserts new text as tracked changes.
 * This is the final fallback when more granular strategies fail.
 *
 * @param {Word.RequestContext} context - The Word request context
 * @param {Word.Range} range - The target range to update
 * @param {string} newText - The new text to apply
 * @param {function} log - Callback for logging messages
 * @param {object} [options]
 * @param {boolean} [options.trackChanges=true] - When false the strategy does
 *   NOT touch the document's changeTrackingMode; the caller owns it.
 * @returns {Promise<{strategy: string, insertions: number, deletions: number}>}
 */
export async function applyBlockReplaceStrategy(context, range, newText, log, options = {}) {
    await loadMutationVerifier();
    const trackChanges = options.trackChanges !== false;
    if (trackChanges && !Word.ChangeTrackingMode) throw new RevisionSafetyError('This Word host cannot enable tracked changes.');
    log('Running block replace (final fallback)...', 'info');

    // True only when THIS strategy enabled tracking (and must restore it).
    let trackingEnabled = false;
    let mutationAttempted = false;

    try {
        if (trackChanges && Word.ChangeTrackingMode) {
            await enableTrackedWrites(context);
            trackingEnabled = true;
        }

        // Get the content range
        const contentRange = range.getRange(Word.RangeLocation.content);

        // Delete the content (tracked when tracking is on)
        mutationAttempted = true;
        contentRange.delete();

        // Insert new text after the deleted range
        // Using 'after' ensures it appears as a replacement in track changes
        const inserted = contentRange.insertText(newText, Word.InsertLocation.after);

        await context.sync();
        await verifyMutationText(context, range, newText, options, [inserted]);
        log('Block replacement applied.', 'info');

        return {
            strategy: 'block',
            insertions: 1,
            deletions: 1,
        };
    } catch (e) {
        log(`Block replace strategy failed: ${e.message}`, 'error');
        if (mutationAttempted || e.mutationAttempted) throw e instanceof MutationSafetyError ? e : new MutationSafetyError(e);
        if (e instanceof RevisionSafetyError) throw e;
        throw new Error(`All diff strategies failed. Final error: ${e.message}`);
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
