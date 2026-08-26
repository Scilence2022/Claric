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

const TABLE_PREVIEW_MAX_ROWS = 30;
const TABLE_PREVIEW_MAX_COLUMNS = 20;
const TABLE_PREVIEW_MAX_CELL_CHARS = 400;
const TABLE_PREVIEW_MAX_META_CHARS = 80;

function _isPlainObject(value) {
    if (!value || typeof value !== 'object') return false;
    const prototype = Object.getPrototypeOf(value);
    return prototype === Object.prototype || prototype === null;
}

function _boundedText(value, maxLength) {
    if (typeof value !== 'string') return null;
    return value.slice(0, maxLength);
}

/**
 * Normalizes the serializable table-preview contract for both live and saved
 * cards. The result contains no references to caller-owned rows or cells.
 *
 * @param {object} preview
 * @returns {object|null}
 */
export function sanitizeTablePreview(preview) {
    if (!_isPlainObject(preview) || !Array.isArray(preview.rows)) return null;

    let wasTruncated = preview.truncated === true;
    const sourceRows = preview.rows;
    if (sourceRows.length > TABLE_PREVIEW_MAX_ROWS) wasTruncated = true;

    const rows = [];
    for (let rowIndex = 0; rowIndex < Math.min(sourceRows.length, TABLE_PREVIEW_MAX_ROWS); rowIndex++) {
        const sourceRow = sourceRows[rowIndex];
        if (!Array.isArray(sourceRow)) {
            wasTruncated = true;
            rows.push([]);
            continue;
        }
        if (sourceRow.length > TABLE_PREVIEW_MAX_COLUMNS) wasTruncated = true;
        const row = [];
        for (let colIndex = 0; colIndex < Math.min(sourceRow.length, TABLE_PREVIEW_MAX_COLUMNS); colIndex++) {
            const cell = sourceRow[colIndex];
            if (cell === null || cell === undefined) {
                row.push('');
            } else if (typeof cell === 'string') {
                if (cell.length > TABLE_PREVIEW_MAX_CELL_CHARS) wasTruncated = true;
                row.push(cell.slice(0, TABLE_PREVIEW_MAX_CELL_CHARS));
            } else {
                wasTruncated = true;
                row.push(typeof cell === 'number' || typeof cell === 'boolean' ? String(cell) : '');
            }
        }
        rows.push(row);
    }

    const columnCount = rows.reduce((max, row) => Math.max(max, row.length), 0);
    for (const row of rows) {
        while (row.length < columnCount) row.push('');
    }

    let headerRowCount = 0;
    if (typeof preview.headerRowCount === 'number' && Number.isFinite(preview.headerRowCount)) {
        headerRowCount = Math.floor(preview.headerRowCount);
        if (headerRowCount < 0 || headerRowCount !== preview.headerRowCount) wasTruncated = true;
        headerRowCount = Math.max(0, headerRowCount);
    } else if (preview.headerRowCount !== undefined && preview.headerRowCount !== null) {
        wasTruncated = true;
    }
    if (headerRowCount > rows.length) {
        headerRowCount = rows.length;
        wasTruncated = true;
    }

    const style = _boundedText(preview.style, TABLE_PREVIEW_MAX_META_CHARS);
    const position = _boundedText(preview.position, TABLE_PREVIEW_MAX_META_CHARS);
    if (preview.style !== undefined && style === null) wasTruncated = true;
    if (preview.position !== undefined && position === null) wasTruncated = true;
    if (typeof preview.style === 'string' && preview.style.length > TABLE_PREVIEW_MAX_META_CHARS) wasTruncated = true;
    if (typeof preview.position === 'string' && preview.position.length > TABLE_PREVIEW_MAX_META_CHARS) wasTruncated = true;

    const normalized = { rows, headerRowCount };
    if (style !== null) normalized.style = style;
    if (position !== null) normalized.position = position;
    if (wasTruncated) normalized.truncated = true;
    return normalized;
}

/**
 * Builds a compact, read-only table preview using DOM nodes only.
 *
 * @param {object} preview
 * @returns {HTMLElement|null}
 */
export function renderTablePreview(preview) {
    const normalized = sanitizeTablePreview(preview);
    if (!normalized) return null;

    const rowCount = normalized.rows.length;
    const columnCount = normalized.rows.reduce((max, row) => Math.max(max, row.length), 0);
    const wrapper = document.createElement('div');
    wrapper.className = 'proposal-card-table-preview';

    const metadata = document.createElement('div');
    metadata.className = 'proposal-card-table-meta';
    const metadataParts = [
        `Dimensions: ${rowCount} × ${columnCount} (rows × columns)`,
        normalized.headerRowCount
            ? `Header: ${normalized.headerRowCount} ${normalized.headerRowCount === 1 ? 'row' : 'rows'}`
            : 'Header: none',
    ];
    if (normalized.style) metadataParts.push(`Style: ${normalized.style}`);
    if (normalized.position) metadataParts.push(`Position: ${normalized.position}`);
    if (normalized.truncated) metadataParts.push('Preview truncated');
    metadata.textContent = metadataParts.join(' · ');
    wrapper.appendChild(metadata);

    const scroll = document.createElement('div');
    scroll.className = 'proposal-card-table-scroll';
    scroll.setAttribute('role', 'region');
    scroll.setAttribute('aria-label', 'Table proposal preview');
    scroll.setAttribute('tabindex', '0');

    const table = document.createElement('table');
    table.className = 'proposal-card-table';
    const caption = document.createElement('caption');
    caption.className = 'proposal-card-table-caption';
    caption.textContent = 'Table proposal preview';
    table.appendChild(caption);

    const thead = document.createElement('thead');
    const tbody = document.createElement('tbody');
    normalized.rows.forEach((row, rowIndex) => {
        const tr = document.createElement('tr');
        const cellTag = rowIndex < normalized.headerRowCount ? 'th' : 'td';
        row.forEach((cell) => {
            const cellEl = document.createElement(cellTag);
            cellEl.textContent = cell;
            if (cellTag === 'th') cellEl.setAttribute('scope', 'col');
            if (!cell) cellEl.setAttribute('aria-label', 'Empty cell');
            tr.appendChild(cellEl);
        });
        (rowIndex < normalized.headerRowCount ? thead : tbody).appendChild(tr);
    });
    if (thead.childElementCount) table.appendChild(thead);
    if (tbody.childElementCount) table.appendChild(tbody);
    scroll.appendChild(table);
    wrapper.appendChild(scroll);

    if (rowCount === 0 || columnCount === 0) {
        const empty = document.createElement('div');
        empty.className = 'proposal-card-table-empty';
        empty.textContent = 'No table cells in preview.';
        wrapper.appendChild(empty);
    }
    return wrapper;
}

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
 * @param {object} [args.tablePreview] - Optional read-only table preview:
 *   { rows: string[][], headerRowCount?: number, style?: string,
 *     position?: string, truncated?: boolean }
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
export function createProposalCard({ title, beforeChars, afterChars, countsText, previewSrc, tablePreview, items, onLocate, comment, onApply, onReject }) {
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

    const tablePreviewEl = renderTablePreview(tablePreview);
    if (tablePreviewEl) el.appendChild(tablePreviewEl);

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

    const api = {
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

    // Reject must go through the public markRejected(): chat-view wraps that
    // method to sync the proposal's history metadata; calling settle()
    // directly would leave saved sessions stuck at "pending".
    rejectBtn.addEventListener('click', () => {
        api.markRejected();
        if (onReject) onReject();
    });

    return api;
}
