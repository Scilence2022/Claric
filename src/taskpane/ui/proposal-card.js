/**
 * Proposal Card Module
 *
 * Builds the staged edit-proposal card shown in chat for selection-scope
 * amendment turns. The LLM's proposed rewrite is NOT applied until the user
 * clicks "Apply as tracked changes"; "Reject" dismisses the card.
 *
 * @module ui/proposal-card
 */

/**
 * Creates a proposal card.
 *
 * @param {object} args
 * @param {string} args.title - Card heading
 * @param {number} args.beforeChars - Character count of the original selection
 * @param {number} args.afterChars - Character count of the proposed rewrite
 * @param {string} [args.comment] - Optional comment text (merged mode)
 * @param {function()} args.onApply - Async apply callback
 * @param {function()} [args.onReject] - Reject callback
 * @returns {{ el: HTMLElement, markApplied: function(), markRejected: function(), markError: function(string) }}
 */
export function createProposalCard({ title, beforeChars, afterChars, comment, onApply, onReject }) {
    const el = document.createElement('div');
    el.className = 'proposal-card';

    const head = document.createElement('div');
    head.className = 'proposal-card-head';
    const titleEl = document.createElement('span');
    titleEl.className = 'proposal-card-title';
    titleEl.textContent = title;
    const counts = document.createElement('span');
    counts.className = 'proposal-card-counts';
    counts.textContent = `${beforeChars} → ${afterChars} chars`;
    head.appendChild(titleEl);
    head.appendChild(counts);
    el.appendChild(head);

    if (comment) {
        const commentEl = document.createElement('div');
        commentEl.className = 'proposal-card-comment';
        commentEl.textContent = `Comment: ${comment}`;
        el.appendChild(commentEl);
    }

    const status = document.createElement('div');
    status.className = 'proposal-card-status';
    status.style.display = 'none';

    const actions = document.createElement('div');
    actions.className = 'proposal-card-actions';

    const applyBtn = document.createElement('button');
    applyBtn.type = 'button';
    applyBtn.className = 'btn btn-primary btn-compact';
    applyBtn.textContent = 'Apply as tracked changes';

    const rejectBtn = document.createElement('button');
    rejectBtn.type = 'button';
    rejectBtn.className = 'btn btn-secondary btn-compact';
    rejectBtn.textContent = 'Reject';

    actions.appendChild(applyBtn);
    actions.appendChild(rejectBtn);
    el.appendChild(actions);
    el.appendChild(status);

    /** Disables both buttons and shows a terminal status line. */
    function settle(text, className) {
        applyBtn.disabled = true;
        rejectBtn.disabled = true;
        el.classList.add(className);
        status.style.display = '';
        status.textContent = text;
    }

    applyBtn.addEventListener('click', async () => {
        applyBtn.disabled = true;
        rejectBtn.disabled = true;
        await onApply();
    });

    rejectBtn.addEventListener('click', () => {
        settle('Rejected — no changes were made.', 'proposal-rejected');
        if (onReject) onReject();
    });

    return {
        el,
        /** Terminal state after a successful apply. */
        markApplied() {
            settle('Applied as tracked changes.', 'proposal-applied');
        },
        /** Terminal state for a rejected proposal (idempotent). */
        markRejected() {
            settle('Rejected — no changes were made.', 'proposal-rejected');
        },
        /** Re-enables Apply after a failed attempt so the user can retry. */
        markError(message) {
            applyBtn.disabled = false;
            rejectBtn.disabled = false;
            status.style.display = '';
            status.textContent = `Apply failed: ${message}`;
            el.classList.add('proposal-error');
        },
    };
}
