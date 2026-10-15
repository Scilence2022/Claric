/** Derive a short, truthful task outcome from the proposal states, live or saved. */
export function buildCompletionSummary({ proposals, status = '', error = null } = {}) {
    const cards = Array.isArray(proposals) ? proposals.filter((p) => p && typeof p === 'object') : [];
    if (!cards.length) return '';

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
    return `Prepared ${cards.length} proposal${cards.length === 1 ? '' : 's'}${scope}${size}. ${outcome}`
        + (brief ? ` Proposed work: ${brief}` : '');
}
