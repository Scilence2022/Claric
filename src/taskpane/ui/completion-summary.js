/** Derive a short, truthful task outcome from live or saved message data. */
export function buildCompletionSummary({ proposals, completionSummary = '', status = '', error = null } = {}) {
    const cards = Array.isArray(proposals) ? proposals.filter((p) => p && typeof p === 'object') : [];
    const completed = typeof completionSummary === 'string' ? completionSummary.trim() : '';
    if (!cards.length) {
        if (error) return `The task ended with an error: ${error}`;
        if (/^Cancel(?:led|ed)\b/i.test(status)) {
            return /^Document review cancelled\b/.test(completed) ? completed
                : 'The task was cancelled. Review the document for any changes already applied.';
        }
        // Older document checks saved only this terminal status. Recover their
        // summary too, without presenting arbitrary progress text as an outcome.
        return completed || (/^Done: \d+ amendment\(s\), \d+ comment\(s\) across \d+ section\(s\)\./.test(status) ? status : '');
    }

    const counts = new Map();
    for (const card of cards) {
        const state = card.state || 'pending';
        counts.set(state, (counts.get(state) || 0) + 1);
    }
    const changes = cards.reduce((sum, card) => sum + (Array.isArray(card.items) ? card.items.length : 0), 0);
    const size = cards.length === 1 && cards[0].countsText ? ` (${cards[0].countsText})` : '';
    const scope = changes ? ` with ${changes} proposed change${changes === 1 ? '' : 's'}` : '';
    let outcome;
    if (error) outcome = 'The task ended with an error. Review any available proposals before continuing.';
    else if (/^Cancelled\b/i.test(status)) outcome = 'The task was cancelled. Review any available proposals and any changes already applied.';
    else if (counts.size === 1 && counts.has('pending')) {
        outcome = 'Ready for review. The proposed changes have not been applied.';
    } else if (counts.size === 1 && counts.has('applied')) {
        // A card may have applied only the checked items; never report every
        // originally proposed item as applied.
        outcome = 'The selected changes were applied.';
    } else if (counts.size === 1 && counts.has('rejected')) outcome = 'The proposals were rejected.';
    else {
        const labels = { pending: 'awaiting review', applied: 'applied', rejected: 'rejected',
            warning: 'requiring review', error: 'with an application error', paused: 'paused during application' };
        outcome = [...counts].map(([state, count]) =>
            `${count} ${Object.hasOwn(labels, state) ? labels[state] : 'requiring review'}`).join('; ') + '.';
        if (counts.has('warning') || counts.has('error') || counts.has('paused')) {
            outcome += ' Inspect the proposal details and the document.';
        }
    }
    const detail = cards.map((card) => typeof card.summary === 'string' ? card.summary.trim() : '')
        .filter(Boolean).join(' ').replace(/\s+/g, ' ');
    const brief = detail.length > 280 ? `${detail.slice(0, 279)}…` : detail;
    return (completed ? `${completed}\n\n` : '')
        + `Prepared ${cards.length} proposal${cards.length === 1 ? '' : 's'}${scope}${size}. ${outcome}`
        + (brief ? ` Proposed work: ${brief}` : '');
}

/** Summarize document checks from actual processing and Word application results. */
export function buildDocumentCompletionSummary({ results = [], chunks = [], applicationResult = {}, cancelled = false } = {}) {
    const reviewed = results.filter((result) => result.status === 'fulfilled').length;
    const failed = results.filter((result) => result.status === 'rejected').length;
    const cancelledCount = results.filter((result) => result.status === 'cancelled').length;
    const errors = applicationResult.errors || [];
    const incomplete = failed > 0 || errors.length > 0 || reviewed < chunks.length
        || applicationResult.halted || applicationResult.partialChunkIds?.length > 0
        || applicationResult.uncertainChunkIds?.length > 0 || applicationResult.failedChunkIds?.length > 0;
    const stopped = cancelled || cancelledCount > 0 || applicationResult.interrupted;
    const outcome = stopped ? 'cancelled' : incomplete ? 'incomplete' : 'completed';
    const parts = [`Document review ${outcome}. Reviewed ${reviewed} of ${chunks.length} section(s).`,
        `Inserted ${applicationResult.commentsInserted || 0} comment(s) in Word; applied ${applicationResult.amendmentsApplied || 0} amendment(s).`];
    const evidenceReviews = results.filter((result) => result.status === 'fulfilled' && result.review);
    if (evidenceReviews.length) {
        const confirmed = evidenceReviews.reduce((sum, result) => sum + result.review.confirmedCount, 0);
        const withheld = evidenceReviews.reduce((sum, result) => sum + result.review.withheldCount, 0);
        parts.push(`Evidence checks confirmed ${confirmed} issue(s).`);
        if (withheld) parts.push(`${withheld} candidate issue(s) were withheld as unsupported or uncertain; they are not correction requests.`);
    }
    if (failed) parts.push(`${failed} section(s) failed during processing.`);
    if (errors.length || applicationResult.failedChunkIds?.length) parts.push('Some results could not be applied. See the activity log for details.');
    if (stopped) parts.push('Any comments and changes already applied remain in the document.');
    else if (incomplete) parts.push('The review did not finish successfully for every section.');
    else if (!chunks.length) parts.push('The document has no sections to review.');
    if (applicationResult.commentsInserted > 0) parts.push('Review the comments in the document.');
    // A comment is a section-level review note, not a count of distinct issues.
    // Keep a small preview and leave the full notes in Word and model activity.
    const notes = results.filter((result) => result.status === 'fulfilled' && typeof result.comment === 'string')
        .map((result) => result.comment.trim()).filter((note) => note && !/^NO ISSUES FOUND\.?$/i.test(note));
    const detail = notes.slice(0, 3).join(' ').replace(/\s+/g, ' ');
    if (detail) parts.push(`Review notes: ${detail.length > 280 ? `${detail.slice(0, 279)}…` : detail}`);
    return parts.join(' ');
}
