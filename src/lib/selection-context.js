/**
 * Selection Context Module
 *
 * Pure formatters that turn structured selection content (Word tables) into
 * LLM-ready markdown, so row/column structure survives the trip into the
 * prompt — selection.text flattens cell text into one string and loses it.
 *
 * Counterpart to selection-with-comments.js: no Office.js, no network,
 * hermetic-testable. The Word-side reader lives in
 * taskpane/word-actions.js readSelectionTableContext.
 *
 * @module selection-context
 */

/**
 * Default token budget for ONE rendered table (estimateTokenCount
 * semantics: 1 token ≈ 4 chars). Oversized tables truncate data rows
 * first — never the header — so column structure survives truncation.
 */
const DEFAULT_TABLE_MAX_TOKENS = 2000;

/**
 * Rendered row for the markdown separator line under the header.
 * Frozen — STYLE.md "Enums for Fixed Values".
 */
const SEPARATOR_CELL = '---';

/**
 * Makes one cell's text safe for a markdown table row: cell-internal
 * newlines (multi-paragraph cells) become <br> (markdown tables cannot
 * hold raw newlines), and pipes are escaped so cell content can never
 * forge column boundaries. CR/LF variants are normalized first.
 *
 * @param {string} text - Raw cell text from Word.Table.values
 * @returns {string}
 */
function sanitizeCellText(text) {
    return String(text || '')
        .replace(/\r\n?/g, '\n')
        .replace(/\n/g, '<br>')
        .replace(/\|/g, '\\|')
        .trim();
}

/**
 * Renders one markdown table line from already-sanitized cells.
 *
 * @param {string[]} cells
 * @returns {string}
 */
function renderRow(cells) {
    return `| ${cells.join(' | ')} |`;
}

/**
 * Formats a Word table's `values` matrix as LLM-ready markdown.
 *
 * The matrix comes straight from Word.Table.values (0-based, full table).
 * Ragged rows are padded to the widest row. The first row renders as the
 * markdown header (markdown tables require one) under a marker line that
 * states the real dimensions, so the model cannot mistake it for a
 * semantic header row; `note` carries selection metadata (e.g. the covered
 * region).
 *
 * Budget enforcement order: data rows drop from the end (header always
 * stays), then columns drop from the right, and a final hard clip guards
 * the pathological single-giant-cell case. Every truncation appends an
 * explicit note — a silently clipped table would corrupt the model's view
 * of the selection.
 *
 * @param {string[][]} values - Full-table matrix from Word.Table.values
 * @param {object} [options]
 * @param {number} [options.maxTokens=DEFAULT_TABLE_MAX_TOKENS] - Render budget
 * @param {string} [options.note] - Selection metadata appended to the marker line
 * @returns {string} '' when values holds no usable rows
 */
export function formatTableMarkdown(values, { maxTokens = DEFAULT_TABLE_MAX_TOKENS, note = '' } = {}) {
    const rows = (Array.isArray(values) ? values : []).filter((r) => Array.isArray(r));
    if (rows.length === 0) return '';

    const colCount = rows.reduce((max, r) => Math.max(max, r.length), 0);
    if (colCount === 0) return '';

    // Sanitize + pad to a rectangle: merged-cell tables can read ragged.
    const grid = rows.map((r) => {
        const cells = r.map(sanitizeCellText);
        while (cells.length < colCount) cells.push('');
        return cells;
    });

    const marker = `[TABLE ${grid.length} rows × ${colCount} cols${note ? ` — ${note}` : ''}]`;
    const maxChars = maxTokens * 4; // estimateTokenCount is ceil(len / 4)

    const render = (nCols, nData) => {
        const lines = [
            marker,
            renderRow(grid[0].slice(0, nCols)),
            renderRow(Array(nCols).fill(SEPARATOR_CELL)),
        ];
        for (let i = 1; i <= nData; i++) lines.push(renderRow(grid[i].slice(0, nCols)));
        return lines.join('\n');
    };

    // Column pass first (rare, header alone busting the budget): drop from
    // the right, keep at least one column.
    let nCols = colCount;
    while (nCols > 1 && render(nCols, 0).length > maxChars) {
        nCols--;
    }

    // Row pass: prefix sums over rendered line lengths keep this O(n)
    // instead of re-rendering the whole table per candidate row count.
    const headerLen = render(nCols, 0).length;
    const dataLines = grid.slice(1).map((r) => renderRow(r.slice(0, nCols)));
    let used = headerLen;
    let nData = 0;
    while (nData < dataLines.length && used + 1 + dataLines[nData].length <= maxChars) {
        used += 1 + dataLines[nData].length; // +1 for the joining newline
        nData++;
    }

    let text = render(nCols, nData);

    // Last-resort guard: one giant cell can still overflow at 1 column.
    let clipped = false;
    if (text.length > maxChars) {
        text = text.slice(0, maxChars);
        clipped = true;
    }

    const truncationNotes = [];
    if (nData < dataLines.length) {
        truncationNotes.push(`[... ${dataLines.length - nData} more row(s) truncated]`);
    }
    if (nCols < colCount) {
        truncationNotes.push(`[... ${colCount - nCols} more column(s) truncated]`);
    }
    if (clipped) {
        truncationNotes.push('[... truncated]');
    }
    if (truncationNotes.length > 0) {
        text += '\n' + truncationNotes.join(' ');
    }

    return text;
}

/**
 * Formats a mixed paragraph+table selection as document-ordered context.
 *
 * Non-table paragraphs keep their text (one per line, blanks dropped —
 * mirroring readMixedTableSelection); each table renders via
 * formatTableMarkdown as its own blank-line-separated block. Each table
 * gets the full per-table budget: mixed selections rarely hold many
 * tables, and the surrounding document context dwarfs them anyway.
 *
 * @param {Array<{type: 'paragraph', text: string}
 *   | {type: 'table', values: string[][], note?: string}>} parts -
 *   Document-ordered selection content
 * @param {object} [options]
 * @param {number} [options.maxTokens=DEFAULT_TABLE_MAX_TOKENS] - Per-table budget
 * @returns {string}
 */
export function formatMixedContext(parts, { maxTokens = DEFAULT_TABLE_MAX_TOKENS } = {}) {
    const blocks = [];
    let paragraphLines = [];

    const flushParagraphs = () => {
        if (paragraphLines.length > 0) {
            blocks.push(paragraphLines.join('\n'));
            paragraphLines = [];
        }
    };

    for (const part of (Array.isArray(parts) ? parts : [])) {
        if (part.type === 'paragraph') {
            const text = (part.text || '').trim();
            if (text) paragraphLines.push(text);
        } else if (part.type === 'table') {
            flushParagraphs();
            blocks.push(formatTableMarkdown(part.values, { maxTokens, note: part.note }));
        }
    }
    flushParagraphs();

    return blocks.filter(Boolean).join('\n\n');
}
