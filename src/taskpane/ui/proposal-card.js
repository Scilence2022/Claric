/**
 * Proposal Card Module
 *
 * Builds the staged edit-proposal card shown in chat for amendment turns
 * (selection-scope edits and gated document-scope runs). The LLM's proposed
 * rewrite is NOT applied until the user clicks "Apply as tracked changes";
 * "Reject" dismisses the card.
 *
 * When `items` are given, the card carries an expandable change list — one
 * checkbox row per change, with an optional inline diff (before/after) and
 * an optional locate-in-document link. Apply then passes the SELECTED item
 * ids to onApply, so the user applies only the changes they checked.
 *
 * @module ui/proposal-card
 */

import { buildTextDiffElement } from './diff-view.js';

/**
 * Creates a proposal card.
 *
 * @param {object} args
 * @param {string} args.title - Card heading
 * @param {number} [args.beforeChars] - Character count of the original selection
 * @param {number} [args.afterChars] - Character count of the proposed rewrite
 * @param {string} [args.countsText] - Overrides the "before → after chars" line
 *   (for non-text proposals such as formatting ops)
 * @param {string} [args.previewSrc] - Optional image data-URL preview shown
 *   under the heading (illustration proposals)
 * @param {Array<object>} [args.items] - Optional change list. Each item:
 *   { id: string|number, label: string, before?: string, after?: string,
 *     searchText?: string }. before+after render an inline diff; searchText
 *   plus onLocate adds a locate-in-document link.
 * @param {function(string)} [args.onLocate] - Locate-in-document callback
 *   (receives the item's searchText)
 * @param {string} [args.comment] - Optional comment text (merged mode)
 * @param {function(Array<string|number>|undefined)} args.onApply - Async
 *   apply callback; receives the selected item ids when items are given
 * @param {function()} [args.onReject] - Reject callback
 * @returns {{ el: HTMLElement, markApplied: function(), markRejected: function(), markWarning: function(string), markError: function(string) }}
 */
export function createProposalCard({ title, beforeChars, afterChars, countsText, previewSrc, items, onLocate, comment, onApply, onReject }) {
    const el = document.createElement('div');
    el.className = 'proposal-card';

    const head = document.createElement('div');
    head.className = 'proposal-card-head';
    const titleEl = document.createElement('span');
    titleEl.className = 'proposal-card-title';
    titleEl.textContent = title;
    const counts = document.createElement('span');
    counts.className = 'proposal-card-counts';
    counts.textContent = countsText || `${beforeChars} → ${afterChars} chars`;
    head.appendChild(titleEl);
    head.appendChild(counts);
    el.appendChild(head);

    if (previewSrc) {
        const img = document.createElement('img');
        img.className = 'proposal-card-preview';
        img.alt = 'Proposal preview';
        img.src = previewSrc;
        el.appendChild(img);
    }

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

    // Expandable per-change list: checkbox + label (+ locate link + inline
    // diff). Apply forwards the checked ids; unchecking everything disables
    // Apply. The applied count is remembered for the terminal status line.
    const changeBoxes = [];
    let lastAppliedCount = 0;
    if (items && items.length) {
        const details = document.createElement('details');
        details.className = 'proposal-card-changes';
        const summary = document.createElement('summary');
        summary.textContent = `${items.length} change(s)`;
        details.appendChild(summary);

        for (const item of items) {
            const row = document.createElement('div');
            row.className = 'proposal-card-change';

            const rowHead = document.createElement('label');
            rowHead.className = 'proposal-card-change-head';
            const box = document.createElement('input');
            box.type = 'checkbox';
            box.checked = true;
            changeBoxes.push(box);
            const text = document.createElement('span');
            text.className = 'proposal-card-change-label';
            text.textContent = item.label;
            rowHead.appendChild(box);
            rowHead.appendChild(text);
            if (item.searchText && onLocate) {
                const locate = document.createElement('button');
                locate.type = 'button';
                locate.className = 'proposal-card-locate';
                locate.title = 'Show in document';
                locate.textContent = '§';
                locate.addEventListener('click', () => onLocate(item.searchText));
                rowHead.appendChild(locate);
            }
            row.appendChild(rowHead);

            if (item.before !== undefined && item.after !== undefined) {
                row.appendChild(buildTextDiffElement(item.before, item.after));
            }
            details.appendChild(row);
        }

        changeBoxes.forEach((box) => box.addEventListener('change', () => {
            applyBtn.disabled = changeBoxes.every((b) => !b.checked);
        }));
        el.appendChild(details);
    }

    /** Ids of the currently checked change items (undefined without items). */
    function selectedIds() {
        if (!changeBoxes.length) return undefined;
        return items.filter((_, i) => changeBoxes[i].checked).map((item) => item.id);
    }

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
        lastAppliedCount = changeBoxes.filter((b) => b.checked).length;
        applyBtn.disabled = true;
        rejectBtn.disabled = true;
        await onApply(selectedIds());
    });

    rejectBtn.addEventListener('click', () => {
        settle('Rejected — no changes were made.', 'proposal-rejected');
        if (onReject) onReject();
    });

    return {
        el,
        /** Terminal state after a successful apply. */
        markApplied() {
            settle(
                changeBoxes.length
                    ? `Applied ${lastAppliedCount} of ${changeBoxes.length} change(s) as tracked changes.`
                    : 'Applied as tracked changes.',
                'proposal-applied'
            );
        },
        /** Terminal state for a rejected proposal (idempotent). */
        markRejected() {
            settle('Rejected — no changes were made.', 'proposal-rejected');
        },
        /** Terminal state when apply finished but nothing (or only part)
         *  landed — "Applied" would be a lie, so surface the reason. */
        markWarning(message) {
            settle(message, 'proposal-warning');
        },
        /** Re-enables Apply after a failed attempt so the user can retry. */
        markError(message) {
            applyBtn.disabled = changeBoxes.length > 0 && changeBoxes.every((b) => !b.checked);
            rejectBtn.disabled = false;
            status.style.display = '';
            status.textContent = `Apply failed: ${message}`;
            el.classList.add('proposal-error');
        },
    };
}
