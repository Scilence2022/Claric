/**
 * Table Tool Model
 *
 * L2 of the tool-calling stack: the draft model table tools operate on.
 * Seeded from ONE region (readSelectionTableRegion — selection path) or a
 * REGIONS ARRAY (readDocumentTableRegions — document-scope table_management).
 * NEVER touches Word — ops accumulate as a transaction, translated at the end
 * into the existing coordinate patch shape (table-patch.js), so the proposal
 * card and applySelectionAmendment are reused unchanged.
 *
 * Two op families: content/structure (set_cell, insert_row, delete_row,
 * merge_cells) and styling (set_table_style, set_borders, set_cell_format,
 * set_font, set_header_row, set_layout, set_column_widths — validated by
 * lib/table-style.js). Style ops with a `region` are coordinate-bound and
 * apply before row structure changes; table-level style ops apply after.
 *
 * Multi-table discipline: every tool accepts a `tableIndex` (1-based,
 * defaults to 1 — single-table callers stay source-compatible). All emitted
 * patch elements carry `tableIndex` so the apply side can anchor each op to
 * the right table. Row ops from different tables do not shift each other.
 *
 * Coordinate discipline (mirrors the apply path): cell coordinates always
 * refer to ORIGINAL row numbering within their own table during the loop —
 * row ops are queued, never shift cell coordinates mid-loop. Insert values
 * must therefore be final. This keeps the translated patch expressible in
 * the existing descending-row-op apply order.
 *
 * Pure module — no Office.js, no network. Hermetic-testable.
 *
 * @module table-model
 */

import { planRowOpOrder } from './table-patch.js';
import { defineTool } from './tool-registry.js';
import {
    TABLE_STYLE_LIMITS,
    BUILT_IN_TABLE_STYLES,
    HORIZONTAL_ALIGNMENTS,
    VERTICAL_ALIGNMENTS,
    expandBorderSet,
    normalizeAlignment,
    normalizeColor,
    normalizeFontPayload,
    normalizeStyleTarget,
} from './table-style.js';

/**
 * Tool specs for the table draft model. The descriptions are the contract
 * the model sees — constraints live here, enforcement lives in the model.
 * Multi-table coverage: tools addressing cells/rows/columns accept a
 * `tableIndex` (1-based, from the get_state listing) — omit it when the
 * session covers a single table.
 */
export const TABLE_TOOL_SPECS = Object.freeze([
    defineTool({
        name: 'get_state',
        description: 'Re-read the covered table region(s): the coordinate grid(s) — merged slots are marked read-only, the current table style (borders, shading, alignments, header rows, font), the pending operations recorded so far, and whether row operations are allowed. When the session covers several tables, each grid is listed under its table index (e.g. "table 1 (3x2):").',
        argsExample: {},
    }),
    defineTool({
        name: 'set_cell',
        description: 'Set the FULL new text of one cell. "tableIndex" (optional, default 1) selects the table. "row"/"col" are 1-based, from the grid listing, and must be inside the covered region. Slots marked "(merged — read-only)" cannot be set. Rows with a pending delete cannot be set. Setting the current text is a no-op.',
        argsExample: { tableIndex: 1, row: 2, col: 1, text: 'new text' },
    }),
    defineTool({
        name: 'insert_row',
        description: 'Insert one new row next to an existing row. "tableIndex" (optional, default 1) selects the table; "values" must contain the FINAL text of every cell in the new row (exactly that table\'s colCount entries, plain text). Only allowed when row operations are enabled (uniform table, full-width region).',
        argsExample: { tableIndex: 1, position: 'after', row: 3, values: ['cell A', 'cell B'] },
    }),
    defineTool({
        name: 'delete_row',
        description: 'Delete one existing row by its 1-based row number. "tableIndex" (optional, default 1) selects the table. Only allowed when row operations are enabled (uniform table, full-width region). The table must keep at least one row.',
        argsExample: { tableIndex: 1, row: 4 },
    }),
    defineTool({
        name: 'merge_cells',
        description: 'Merge a rectangular block of cells into one cell. "tableIndex" (optional, default 1) selects the table; "row"/"col" is the top-left cell (1-based, from the grid); "rows"/"cols" is the block size (the block must span 2+ cells). All cells in the block must be editable (not already merged away, not in a pending-delete row). The top-left cell becomes the anchor. Only one merge can be staged per table, and staging it disables further row operations on that table.',
        argsExample: { tableIndex: 1, row: 1, col: 1, rows: 2, cols: 2 },
    }),
    defineTool({
        name: 'set_table_style',
        description: 'Set the overall LOOK of one table (the whole table; "tableIndex" optional, default 1): a built-in Word table style plus its banding/emphasis options. "style" must be a built-in style name: TableGrid, TableGridLight, PlainTable1..PlainTable5, or GridTable/ListTable families (GridTable1Light, GridTable2, GridTable3, GridTable4, GridTable5Dark, GridTable6Colorful, GridTable7Colorful, ListTable1Light, ListTable2, ..., ListTable7Colorful), each optionally suffixed _Accent1.._Accent6. Optional boolean flags tune the style: bandedRows, bandedColumns, firstColumn, lastColumn, totalRow. To remove or draw specific borders use set_borders instead.',
        argsExample: { tableIndex: 1, style: 'GridTable4_Accent1', bandedRows: true, firstColumn: true },
    }),
    defineTool({
        name: 'set_borders',
        description: 'Set borders for ONE table (whole table), or one row\'s borders when "row" is given. "tableIndex" (optional, default 1) selects the table. "borders" keys: top, bottom, left, right, insideH, insideV, plus shorthands all / outside / inside. Each value is a border type string ("none", "single", "double", "dotted", "dashed", "dotDashed", "triple", "wave", ...) or {type, color?, width?} with color "#RRGGBB"/"auto"/a color name and width in points. Academic three-line table: set_borders({borders:{top:{type:"single",width:1.5}, bottom:{type:"single",width:1.5}, inside:"none"}}) then set_borders({row:1, borders:{bottom:{type:"single",width:0.75}}}).',
        argsExample: { tableIndex: 1, borders: { top: { type: 'single', width: 1.5 }, bottom: { type: 'single', width: 1.5 }, inside: 'none' } },
    }),
    defineTool({
        name: 'set_cell_format',
        description: 'Set cell shading (background color) and content alignment. "tableIndex" (optional, default 1) selects the table. Target: {row, col} one cell, {row, rows} a row band, {col, cols} a column band, {row, col, rows, cols} a block, or omit target for the whole table. Payload keys: shadingColor ("#RRGGBB", "auto", or a color name), horizontalAlignment (left | centered | right | justified), verticalAlignment (top | center | bottom). Coordinates covered by a merge resolve to the merged cell. Cannot target rows with a pending delete.',
        argsExample: { tableIndex: 1, row: 1, shadingColor: '#DEEBF7', horizontalAlignment: 'centered' },
    }),
    defineTool({
        name: 'set_font',
        description: 'Set font properties. "tableIndex" (optional, default 1) selects the table. Same targeting as set_cell_format ({row?, col?, rows?, cols?}; omit for the whole table). "font": {bold?, italic?, underline? ("none"|"single"|"double"|...), size? (points), name?, color?}. Example header emphasis: {row: 1, font: {bold: true}}.',
        argsExample: { tableIndex: 1, row: 1, font: { bold: true, size: 11 } },
    }),
    defineTool({
        name: 'set_header_row',
        description: 'Mark the first N rows of one table as the header row(s) — Word repeats them on every page — and optionally style them. "tableIndex" (optional, default 1) selects the table. args: {rows? (default 1), font?, shadingColor?}. Only allowed when the covered region includes that table\'s row 1.',
        argsExample: { tableIndex: 1, rows: 1, font: { bold: true }, shadingColor: '#DEEBF7' },
    }),
    defineTool({
        name: 'set_layout',
        description: 'Table-level layout for ONE table ("tableIndex" optional, default 1): alignment of the table against the page column ("left" | "centered" | "right"), widthPt (total width in points), autoFitWindow (true = stretch columns to the window), distributeColumns (true = equal column widths), cellPaddingPt (uniform cell padding on all sides, 0–100pt).',
        argsExample: { tableIndex: 1, alignment: 'centered', autoFitWindow: true },
    }),
    defineTool({
        name: 'set_column_widths',
        description: 'Set per-column widths of ONE table in points: {widthsPt: [...]} with exactly that table\'s colCount entries (5–1584 each). "tableIndex" (optional, default 1) selects the table. Only allowed on uniform tables (no merged cells) — merged tables reject it.',
        argsExample: { tableIndex: 1, widthsPt: [120, 80, 200] },
    }),
]);

/**
 * Creates one table draft (the per-table op accumulator).
 *
 * @param {object} region - Table region: { rowCount, colCount, values,
 *   bounds, merged?, shadowKeys?, style? }
 * @param {number} tableIndex - 1-based index of this table in the session
 * @returns {object} Draft with getState/mutators/toPatch sections
 * @private
 */
function createTableDraft(region, tableIndex) {
    const rowCount = region.rowCount;
    const colCount = region.colCount;
    const values = region.values || [];
    const bounds = region.bounds || { startRow: 1, endRow: rowCount, startCol: 1, endCol: colCount };
    const merged = !!region.merged;
    const shadowKeys = region.shadowKeys instanceof Set ? region.shadowKeys : new Set();
    const allowRowOps = !merged && bounds.startCol === 1 && bounds.endCol === colCount;
    const styleSnapshot = region.style && typeof region.style === 'object' ? region.style : null;

    /** @type {Map<string, {row: number, col: number, text: string}>} */
    const cellEdits = new Map();
    /** @type {Array<{op: string, row: number, values?: string[]}>} */
    const rowOps = [];
    /** @type {Array<{op: string, startRow: number, startCol: number, endRow: number, endCol: number}>} */
    const mergeOps = [];
    /** @type {Array<Record<string, *>>} */
    const styleOps = [];
    /** @type {Set<string>} */
    const mergedAway = new Set();

    const deletedRows = () => new Set(rowOps.filter((o) => o.op === 'delete').map((o) => o.row));
    const rowOpsBlocked = () => mergeOps.length > 0;

    const err = (error) => ({ ok: false, error });

    function _lookup(key) {
        const [r, c] = key.split(',').map(Number);
        return cellEdits.has(key) ? cellEdits.get(key).text : (values[r - 1] && values[r - 1][c - 1]) || '';
    }

    function grid() {
        const lines = [];
        for (let r = bounds.startRow; r <= bounds.endRow; r++) {
            for (let c = bounds.startCol; c <= bounds.endCol; c++) {
                const key = `${r},${c}`;
                const current = _lookup(key);
                lines.push(`[R${r}C${c}]${shadowKeys.has(key) || mergedAway.has(key) ? ' (merged — read-only)' : ''} ${current}`);
            }
        }
        return lines.join('\n');
    }

    function pendingOps() {
        return [
            ...[...cellEdits.values()].map((e) => ({ tool: 'set_cell', tableIndex, ...e })),
            ...rowOps.map((o) => ({ ...o, tableIndex })),
            ...mergeOps.map((m) => ({ tool: 'merge_cells', tableIndex, ...m })),
            ...styleOps.map((s) => {
                const copy = { ...s, tableIndex };
                // Region-bound style ops keep their region (already validated).
                return copy;
            }),
        ];
    }

    return {
        tableIndex,
        rowCount,
        colCount,
        values,
        bounds,
        merged,
        allowRowOps,
        shadowKeys,
        grid,
        pendingOps,
        styleSnapshot,
        cellEdits,
        rowOps,
        mergeOps,
        styleOps,
        mergedAway,
        deletedRows,
        rowOpsBlocked,
        err,
    };
}

/** @private */
function _styleBudgetError(draft) {
    if (draft.styleOps.length >= TABLE_STYLE_LIMITS.MAX_STYLE_OPS) {
        return `Style op limit reached (${TABLE_STYLE_LIMITS.MAX_STYLE_OPS}) — finish with what is staged.`;
    }
    return null;
}

/** True when the region intersects any pending-delete row of the draft. @private */
function _intersectsDeletedRow(draft, region) {
    if (!region) return false;
    const deleted = draft.deletedRows();
    for (let r = region.startRow; r <= region.endRow; r++) {
        if (deleted.has(r)) return true;
    }
    return false;
}

/**
 * Creates the table draft model. Accepts a single region (selection path)
 * or an array of regions (document-scope table_management path).
 *
 * @param {object|object[]} input - One region or an array of regions:
 *   { rowCount, colCount, values (string[][]), bounds {startRow,endRow,startCol,endCol},
 *     merged?: boolean, shadowKeys?: Set<string> ("row,col" 1-based),
 *     style?: object|null (advisory style snapshot for get_state) }
 * @returns {{getState: Function, setCell: Function, insertRow: Function,
 *   deleteRow: Function, mergeCells: Function, setTableStyle: Function,
 *   setBorders: Function, setCellFormat: Function, setFont: Function,
 *   setHeaderRow: Function, setLayout: Function, setColumnWidths: Function,
 *   toTablePatch: Function, opCount: number}}
 */
export function createTableModel(input) {
    const regions = Array.isArray(input) ? input : [input || {}];
    const drafts = regions.map((region, i) => createTableDraft(region, i + 1));

    const _ok = (result) => ({ ok: true, result });

    /** Resolves args.tableIndex against the covered tables. @private */
    function _draft(args) {
        const tableIndex = args && args.tableIndex !== undefined && args.tableIndex !== null
            ? Number(args.tableIndex)
            : 1;
        const draft = drafts.find((d) => d.tableIndex === tableIndex);
        if (!draft) {
            return {
                error: `Unknown "tableIndex" ${tableIndex} — this session covers table(s) ${drafts.map((d) => d.tableIndex).join(', ')}.`,
            };
        }
        return { draft };
    }

    /** True when the session covers more than one table. @private */
    const multiTable = drafts.length > 1;

    function getState() {
        const tables = drafts.map((draft) => ({
            tableIndex: draft.tableIndex,
            rowCount: draft.rowCount,
            colCount: draft.colCount,
            coveredRegion: `R${draft.bounds.startRow}C${draft.bounds.startCol}–R${draft.bounds.endRow}C${draft.bounds.endCol}`,
            mergedTable: draft.merged,
            rowOpsAllowed: draft.allowRowOps,
            grid: draft.grid(),
            style: _describeStyleSnapshot(draft),
        }));
        const pendingOps = drafts.flatMap((draft) => draft.pendingOps());
        // Single-table sessions surface legacy-shaped pending ops (no
        // tableIndex) so the existing review/UI code stays compatible.
        const shownPendingOps = drafts.length === 1
            ? pendingOps.map((op) => {
                const { tableIndex: _ti, ...rest } = op;
                return rest;
            })
            : pendingOps;
        const first = drafts[0];
        return _ok({
            tableCount: drafts.length,
            rowCount: first.rowCount,
            colCount: first.colCount,
            coveredRegion: `R${first.bounds.startRow}C${first.bounds.startCol}–R${first.bounds.endRow}C${first.bounds.endCol}`,
            mergedTable: first.merged,
            rowOpsAllowed: first.allowRowOps,
            grid: multiTable
                ? tables.map((t) => `table ${t.tableIndex} (${t.rowCount}x${t.colCount}):\n${t.grid}`).join('\n\n')
                : first.grid(),
            style: _describeStyleSnapshot(first),
            tables,
            pendingOps: shownPendingOps,
        });
    }

    /** @private */
    function _describeStyleSnapshot(draft) {
        const s = draft.styleSnapshot;
        if (!s) return '(style snapshot unavailable on this host)';
        const lines = [];
        lines.push(`style: ${s.styleBuiltIn || s.style || 'unknown'}${s.headerRowCount ? `, header rows: ${s.headerRowCount}` : ''}`);
        lines.push(`table alignment: ${s.alignment || 'unknown'}; cell alignment: ${s.horizontalAlignment || '?'} / ${s.verticalAlignment || '?'} (h/v)`);
        lines.push(`shading: ${s.shadingColor || 'none'}; font: ${s.font ? `${s.font.name || '?'} ${s.font.size || '?'}pt${s.font.bold ? ' bold' : ''}` : 'unknown'}`);
        if (s.borders) {
            lines.push(`borders: ${Object.entries(s.borders)
                .map(([loc, b]) => `${loc}=${b && b.type !== 'none' ? `${b.type}${b.width ? ` ${b.width}pt` : ''}` : 'none'}`)
                .join(', ')}`);
        }
        return lines.join('\n');
    }

    function setCell(row, col, text, _tableIndex) {
        const resolved = _draft({ tableIndex: _tableIndex });
        if (resolved.error) return { ok: false, error: resolved.error };
        const draft = resolved.draft;
        const { rowCount, colCount, bounds } = draft;
        if (!Number.isInteger(row) || !Number.isInteger(col)
            || row < 1 || row > rowCount || col < 1 || col > colCount) {
            return { ok: false, error: `Coordinates out of bounds (row=${row}, col=${col}); table ${draft.tableIndex} is ${rowCount}x${colCount}.` };
        }
        if (row < bounds.startRow || row > bounds.endRow || col < bounds.startCol || col > bounds.endCol) {
            return { ok: false, error: `Cell R${row}C${col} is outside the covered region R${bounds.startRow}C${bounds.startCol}–R${bounds.endRow}C${bounds.endCol}.` };
        }
        if (draft.shadowKeys.has(`${row},${col}`)) {
            return { ok: false, error: `Cell R${row}C${col} is covered by a merged cell — it is read-only. Edit the merge anchor slot instead.` };
        }
        if (draft.mergedAway.has(`${row},${col}`)) {
            return { ok: false, error: `Cell R${row}C${col} is set to be merged away — it is read-only. Set the merge anchor (top-left of the block) instead.` };
        }
        if (typeof text !== 'string') {
            return { ok: false, error: '"text" must be a string.' };
        }
        if (draft.deletedRows().has(row)) {
            return { ok: false, error: `Row ${row} has a pending delete — setting cells in it is not allowed.` };
        }
        const key = `${row},${col}`;
        const original = (draft.values[row - 1] && draft.values[row - 1][col - 1]) || '';
        if (text.trim() === original.trim() && !draft.cellEdits.has(key)) {
            return _ok({ note: `R${row}C${col} already holds that text — no change.` });
        }
        draft.cellEdits.set(key, { row, col, text });
        return _ok({ set: `R${row}C${col}` });
    }

    function insertRow(args = {}) {
        const resolved = _draft(args);
        if (resolved.error) return { ok: false, error: resolved.error };
        const draft = resolved.draft;
        const { rowCount, colCount, bounds } = draft;
        if (draft.merged) {
            return { ok: false, error: 'Row operations are not allowed: the table contains merged cells.' };
        }
        if (draft.rowOpsBlocked()) {
            return { ok: false, error: 'Row operations are not allowed while a cell merge is pending — apply the merge or clear it first.' };
        }
        if (!draft.allowRowOps) {
            return { ok: false, error: 'Row operations require a full-width selection covering every column.' };
        }
        if (args.position !== 'after' && args.position !== 'before') {
            return { ok: false, error: '"position" must be "after" or "before".' };
        }
        if (!Number.isInteger(args.row) || args.row < 1 || args.row > rowCount) {
            return { ok: false, error: `Anchor row out of bounds (row=${args.row}); table ${draft.tableIndex} has ${rowCount} rows.` };
        }
        if (args.row < bounds.startRow || args.row > bounds.endRow) {
            return { ok: false, error: `Anchor row ${args.row} is outside the covered region.` };
        }
        if (!Array.isArray(args.values) || args.values.length !== colCount) {
            return { ok: false, error: `"values" must be an array of exactly ${colCount} strings (one per column).` };
        }
        if (!args.values.every((v) => typeof v === 'string')) {
            return { ok: false, error: 'Every entry in "values" must be a string.' };
        }
        const op = args.position === 'after' ? 'insertAfter' : 'insertBefore';
        draft.rowOps.push({ op, row: args.row, values: [...args.values] });
        return _ok({ inserted: `${op} row ${args.row}` });
    }

    function deleteRow(row, _tableIndex) {
        const resolved = _draft({ tableIndex: _tableIndex });
        if (resolved.error) return { ok: false, error: resolved.error };
        const draft = resolved.draft;
        const { rowCount, bounds } = draft;
        if (draft.merged) {
            return { ok: false, error: 'Row operations are not allowed: the table contains merged cells.' };
        }
        if (draft.rowOpsBlocked()) {
            return { ok: false, error: 'Row operations are not allowed while a cell merge is pending — apply the merge or clear it first.' };
        }
        if (!draft.allowRowOps) {
            return { ok: false, error: 'Row operations require a full-width selection covering every column.' };
        }
        if (!Number.isInteger(row) || row < 1 || row > rowCount) {
            return { ok: false, error: `Row out of bounds (row=${row}); table ${draft.tableIndex} has ${rowCount} rows.` };
        }
        if (row < bounds.startRow || row > bounds.endRow) {
            return { ok: false, error: `Row ${row} is outside the covered region.` };
        }
        if (draft.deletedRows().size + 1 >= rowCount) {
            return { ok: false, error: 'Deleting this row would remove every row of the table — not allowed.' };
        }
        if (draft.deletedRows().has(row)) {
            return _ok({ note: `Row ${row} already has a pending delete.` });
        }
        draft.rowOps.push({ op: 'delete', row });
        return _ok({ deleted: `row ${row}` });
    }

    /**
     * Stages a rectangular cell merge within one table.
     *
     * @param {object} args - {tableIndex?, row, col, rows, cols}
     * @returns {{ok: boolean, result?: *, error?: string}}
     */
    function mergeCells(args = {}) {
        const resolved = _draft(args);
        if (resolved.error) return { ok: false, error: resolved.error };
        const draft = resolved.draft;
        const { row: row_, col: col_, rows, cols } = args;
        if (!Number.isInteger(row_) || !Number.isInteger(col_) || !Number.isInteger(rows) || !Number.isInteger(cols)) {
            return { ok: false, error: '"row", "col", "rows", "cols" must be integers.' };
        }
        const row = row_;
        const col = col_;
        if (rows < 1 || cols < 1) {
            return { ok: false, error: '"rows" and "cols" must be ≥ 1.' };
        }
        const endRow = row + rows - 1;
        const endCol = col + cols - 1;
        if (row < 1 || endRow > draft.rowCount || col < 1 || endCol > draft.colCount) {
            return { ok: false, error: `Merge region R${row}C${col}–R${endRow}C${endCol} is outside the ${draft.rowCount}x${draft.colCount} table.` };
        }
        if (row < draft.bounds.startRow || endRow > draft.bounds.endRow || col < draft.bounds.startCol || endCol > draft.bounds.endCol) {
            return { ok: false, error: `Merge region R${row}C${col}–R${endRow}C${endCol} is outside the covered region R${draft.bounds.startRow}C${draft.bounds.startCol}–R${draft.bounds.endRow}C${draft.bounds.endCol}.` };
        }
        if (rows * cols < 2) {
            return { ok: false, error: 'Merging a single cell is a no-op — the block must span 2+ cells.' };
        }
        if (draft.rowOpsBlocked()) {
            return { ok: false, error: 'Only one cell merge can be staged at a time.' };
        }
        for (let r = row; r <= endRow; r++) {
            for (let c = col; c <= endCol; c++) {
                const key = `${r},${c}`;
                if (draft.shadowKeys.has(key)) {
                    return { ok: false, error: `Cell R${r}C${c} is covered by an existing merged cell — cannot merge.` };
                }
                if (draft.mergedAway.has(key)) {
                    return { ok: false, error: `Cell R${r}C${c} is already part of a staged merge.` };
                }
                if (draft.deletedRows().has(r)) {
                    return { ok: false, error: `Row ${r} has a pending delete — cannot merge.` };
                }
                if (r !== row || c !== col) draft.mergedAway.add(key);
            }
        }
        draft.mergeOps.push({ op: 'merge', startRow: row, startCol: col, endRow, endCol });
        return _ok({ merged: `R${row}C${col}–R${endRow}C${endCol} (anchor R${row}C${col})` });
    }

    /**
     * Stages the whole-table look: built-in style plus banding/emphasis flags.
     *
     * @param {{tableIndex?: number, style?: string, bandedRows?: boolean}} args
     */
    function setTableStyle(args = {}) {
        const resolved = _draft(args);
        if (resolved.error) return { ok: false, error: resolved.error };
        const draft = resolved.draft;
        const budgetError = _styleBudgetError(draft);
        if (budgetError) return { ok: false, error: budgetError };
        /** @type {Record<string, *>} */
        const op = { type: 'tableStyle', tool: 'set_table_style', tableIndex: draft.tableIndex };
        if (args.style !== undefined && args.style !== null) {
            if (typeof args.style !== 'string') return { ok: false, error: '"style" must be a built-in table style name string.' };
            const canonical = BUILT_IN_TABLE_STYLES.find((name) =>
                name.replace(/[\s_-]+/g, '').toLowerCase() === args.style.replace(/[\s_-]+/g, '').toLowerCase());
            if (!canonical) {
                return { ok: false, error: `Unknown built-in table style "${args.style}". Use e.g. TableGrid, TableGridLight, PlainTable1..5, or a GridTable/ListTable family (optionally _Accent1..6).` };
            }
            op.style = canonical;
        }
        for (const key of ['bandedRows', 'bandedColumns', 'firstColumn', 'lastColumn', 'totalRow']) {
            if (args[key] === undefined || args[key] === null) continue;
            if (typeof args[key] !== 'boolean') return { ok: false, error: `"${key}" must be true or false.` };
            op[key] = args[key];
        }
        if (Object.keys(op).length <= 3) return { ok: false, error: 'Give at least one of style, bandedRows, bandedColumns, firstColumn, lastColumn, totalRow.' };
        draft.styleOps.push(op);
        return _ok({ staged: `table ${draft.tableIndex} style${op.style ? ` ${op.style}` : ''}` });
    }

    /**
     * Stages border changes for one table, or one row of that table.
     *
     * @param {{tableIndex?: number, row?: number, borders?: object}} args
     */
    function setBorders(args = {}) {
        const resolved = _draft(args);
        if (resolved.error) return { ok: false, error: resolved.error };
        const draft = resolved.draft;
        const budgetError = _styleBudgetError(draft);
        if (budgetError) return { ok: false, error: budgetError };
        const expanded = expandBorderSet(args.borders);
        if (expanded.error) return { ok: false, error: expanded.error };
        /** @type {Record<string, *>} */
        const op = { type: 'borders', tool: 'set_borders', tableIndex: draft.tableIndex, borders: expanded.borders };
        if (args.row !== undefined && args.row !== null) {
            if (!Number.isInteger(args.row) || args.row < 1 || args.row > draft.rowCount) {
                return { ok: false, error: `Row out of bounds (row=${args.row}); table ${draft.tableIndex} has ${draft.rowCount} rows.` };
            }
            if (args.row < draft.bounds.startRow || args.row > draft.bounds.endRow) {
                return { ok: false, error: `Row ${args.row} is outside the covered region.` };
            }
            if (draft.deletedRows().has(args.row)) {
                return { ok: false, error: `Row ${args.row} has a pending delete — cannot style it.` };
            }
            op.row = args.row;
        }
        draft.styleOps.push(op);
        return _ok({ staged: `table ${draft.tableIndex}${op.row ? ` row ${op.row}` : ''} borders: ${Object.keys(op.borders).join(', ')}` });
    }

    /**
     * Stages cell shading / alignment for a target region (or the whole
     * table when no coordinates are given).
     *
     * @param {object} args - {tableIndex?, row?, col?, rows?, cols?, shadingColor?,
     *   horizontalAlignment?, verticalAlignment?}
     */
    function setCellFormat(args = {}) {
        const resolved = _draft(args);
        if (resolved.error) return { ok: false, error: resolved.error };
        const draft = resolved.draft;
        const budgetError = _styleBudgetError(draft);
        if (budgetError) return { ok: false, error: budgetError };
        const target = normalizeStyleTarget(_pickTarget(args), draft.bounds);
        if (target.error) return { ok: false, error: target.error };
        if (_intersectsDeletedRow(draft, target.region)) {
            return { ok: false, error: 'Cannot format rows with a pending delete.' };
        }
        /** @type {Record<string, *>} */
        const op = { type: 'cellFormat', tool: 'set_cell_format', tableIndex: draft.tableIndex, region: target.region };
        if (args.shadingColor !== undefined && args.shadingColor !== null) {
            const color = normalizeColor(args.shadingColor);
            if (!color) return { ok: false, error: `Invalid shadingColor "${args.shadingColor}" (use "#RRGGBB", "auto", or a color name).` };
            op.shadingColor = color;
        }
        if (args.horizontalAlignment !== undefined && args.horizontalAlignment !== null) {
            const align = normalizeAlignment(args.horizontalAlignment, HORIZONTAL_ALIGNMENTS);
            if (!align) return { ok: false, error: '"horizontalAlignment" must be left | centered | right | justified.' };
            op.horizontalAlignment = align;
        }
        if (args.verticalAlignment !== undefined && args.verticalAlignment !== null) {
            const align = normalizeAlignment(args.verticalAlignment, VERTICAL_ALIGNMENTS);
            if (!align) return { ok: false, error: '"verticalAlignment" must be top | center | bottom.' };
            op.verticalAlignment = align;
        }
        if (!op.shadingColor && !op.horizontalAlignment && !op.verticalAlignment) {
            return { ok: false, error: 'Give at least one of shadingColor, horizontalAlignment, verticalAlignment.' };
        }
        draft.styleOps.push(op);
        return _ok({ staged: `format table ${draft.tableIndex} ${_regionLabel(op.region)}` });
    }

    /**
     * Stages font properties for a target region (or the whole table).
     *
     * @param {object} args - {tableIndex?, row?, col?, rows?, cols?, font}
     */
    function setFont(args = {}) {
        const resolved = _draft(args);
        if (resolved.error) return { ok: false, error: resolved.error };
        const draft = resolved.draft;
        const budgetError = _styleBudgetError(draft);
        if (budgetError) return { ok: false, error: budgetError };
        const target = normalizeStyleTarget(_pickTarget(args), draft.bounds);
        if (target.error) return { ok: false, error: target.error };
        if (_intersectsDeletedRow(draft, target.region)) {
            return { ok: false, error: 'Cannot set fonts on rows with a pending delete.' };
        }
        const font = normalizeFontPayload(args.font);
        if (font.error) return { ok: false, error: font.error };
        draft.styleOps.push({ type: 'font', tool: 'set_font', tableIndex: draft.tableIndex, region: target.region, font: font.font });
        return _ok({ staged: `font table ${draft.tableIndex} ${_regionLabel(target.region)}: ${Object.keys(font.font).join(', ')}` });
    }

    /**
     * Stages the table-header setting: headerRowCount plus optional header
     * row formatting. Header rows are the table's FIRST rows, so the covered
     * region must include row 1.
     *
     * @param {{tableIndex?: number, rows?: number, font?: object, shadingColor?: string}} args
     */
    function setHeaderRow(args = {}) {
        const resolved = _draft(args);
        if (resolved.error) return { ok: false, error: resolved.error };
        const draft = resolved.draft;
        const budgetError = _styleBudgetError(draft);
        if (budgetError) return { ok: false, error: budgetError };
        if (draft.bounds.startRow !== 1) {
            return { ok: false, error: 'Header rows are the table\'s first rows — the covered region must include row 1.' };
        }
        const rows = args.rows === undefined || args.rows === null ? 1 : args.rows;
        if (!Number.isInteger(rows) || rows < 1 || rows > draft.bounds.endRow) {
            return { ok: false, error: `"rows" must be 1–${draft.bounds.endRow}.` };
        }
        if (_intersectsDeletedRow(draft, { startRow: 1, endRow: rows, startCol: 1, endCol: draft.colCount })) {
            return { ok: false, error: 'Cannot make a row with a pending delete a header.' };
        }
        /** @type {Record<string, *>} */
        const op = { type: 'headerRow', tool: 'set_header_row', tableIndex: draft.tableIndex, rows };
        if (args.font !== undefined && args.font !== null) {
            const font = normalizeFontPayload(args.font);
            if (font.error) return { ok: false, error: font.error };
            op.font = font.font;
        }
        if (args.shadingColor !== undefined && args.shadingColor !== null) {
            const color = normalizeColor(args.shadingColor);
            if (!color) return { ok: false, error: `Invalid shadingColor "${args.shadingColor}".` };
            op.shadingColor = color;
        }
        draft.styleOps.push(op);
        return _ok({ staged: `table ${draft.tableIndex} header row(s) 1–${rows}` });
    }

    /**
     * Stages table-level layout: page alignment, width, autofit, column
     * distribution, cell padding.
     *
     * @param {{tableIndex?: number, alignment?: string, widthPt?: number,
     *   autoFitWindow?: boolean, distributeColumns?: boolean, cellPaddingPt?: number}} args
     */
    function setLayout(args = {}) {
        const resolved = _draft(args);
        if (resolved.error) return { ok: false, error: resolved.error };
        const draft = resolved.draft;
        const budgetError = _styleBudgetError(draft);
        if (budgetError) return { ok: false, error: budgetError };
        /** @type {Record<string, *>} */
        const op = { type: 'layout', tool: 'set_layout', tableIndex: draft.tableIndex };
        if (args.alignment !== undefined && args.alignment !== null) {
            const align = normalizeAlignment(args.alignment, ['left', 'centered', 'right']);
            if (!align) return { ok: false, error: '"alignment" must be left | centered | right.' };
            op.alignment = align;
        }
        if (args.widthPt !== undefined && args.widthPt !== null) {
            const width = Number(args.widthPt);
            if (!Number.isFinite(width) || width < 10 || width > TABLE_STYLE_LIMITS.MAX_TABLE_WIDTH_PT) {
                return { ok: false, error: `"widthPt" must be 10–${TABLE_STYLE_LIMITS.MAX_TABLE_WIDTH_PT} points.` };
            }
            op.widthPt = Math.round(width);
        }
        for (const key of ['autoFitWindow', 'distributeColumns']) {
            if (args[key] === undefined || args[key] === null) continue;
            if (typeof args[key] !== 'boolean') return { ok: false, error: `"${key}" must be true or false.` };
            op[key] = args[key];
        }
        if (args.cellPaddingPt !== undefined && args.cellPaddingPt !== null) {
            const padding = Number(args.cellPaddingPt);
            if (!Number.isFinite(padding) || padding < TABLE_STYLE_LIMITS.MIN_CELL_PADDING_PT
                || padding > TABLE_STYLE_LIMITS.MAX_CELL_PADDING_PT) {
                return { ok: false, error: `"cellPaddingPt" must be ${TABLE_STYLE_LIMITS.MIN_CELL_PADDING_PT}–${TABLE_STYLE_LIMITS.MAX_CELL_PADDING_PT} points.` };
            }
            op.cellPaddingPt = Math.round(padding * 10) / 10;
        }
        if (Object.keys(op).length <= 3) {
            return { ok: false, error: 'Give at least one of alignment, widthPt, autoFitWindow, distributeColumns, cellPaddingPt.' };
        }
        draft.styleOps.push(op);
        return _ok({ staged: `table ${draft.tableIndex} layout` });
    }

    /**
     * Stages per-column widths. Uniform tables only — merged grids have no
     * well-defined columns.
     *
     * @param {{tableIndex?: number, widthsPt?: number[]}} args
     */
    function setColumnWidths(args = {}) {
        const resolved = _draft(args);
        if (resolved.error) return { ok: false, error: resolved.error };
        const draft = resolved.draft;
        const budgetError = _styleBudgetError(draft);
        if (budgetError) return { ok: false, error: budgetError };
        if (draft.merged) {
            return { ok: false, error: 'Column widths are not supported: the table contains merged cells.' };
        }
        const widths = args && args.widthsPt;
        if (!Array.isArray(widths) || widths.length !== draft.colCount) {
            return { ok: false, error: `"widthsPt" must be an array of exactly ${draft.colCount} numbers (one per column).` };
        }
        const normalized = [];
        for (const value of widths) {
            const width = Number(value);
            if (!Number.isFinite(width) || width < TABLE_STYLE_LIMITS.MIN_COLUMN_WIDTH_PT
                || width > TABLE_STYLE_LIMITS.MAX_COLUMN_WIDTH_PT) {
                return { ok: false, error: `Every width must be ${TABLE_STYLE_LIMITS.MIN_COLUMN_WIDTH_PT}–${TABLE_STYLE_LIMITS.MAX_COLUMN_WIDTH_PT}pt.` };
            }
            normalized.push(Math.round(width));
        }
        draft.styleOps.push({ type: 'columnWidths', tool: 'set_column_widths', tableIndex: draft.tableIndex, widthsPt: normalized });
        return _ok({ staged: `table ${draft.tableIndex} column widths: ${normalized.join('/')}pt` });
    }

    /** Target sub-object shared by set_cell_format / set_font. @private */
    function _pickTarget(args) {
        if (args.row === undefined && args.col === undefined
            && args.rows === undefined && args.cols === undefined) return null;
        return { row: args.row, col: args.col, rows: args.rows, cols: args.cols };
    }

    /** Compact region label for tool observations. @private */
    function _regionLabel(region) {
        if (!region) return 'whole table';
        const single = region.startRow === region.endRow && region.startCol === region.endCol;
        return single
            ? `R${region.startRow}C${region.startCol}`
            : `R${region.startRow}C${region.startCol}–R${region.endRow}C${region.endCol}`;
    }

    /**
     * Translates the recorded ops into the tablePatch shape consumed by the
     * existing proposal card and applySelectionAmendment. Multi-table sessions
     * carry `tableIndex` on every element; single-table sessions emit the
     * legacy flat shapes (no tableIndex) so the selection path and its
     * apply/review code stay source-compatible.
     *
     * @returns {{tableCount: number, rowCount: number, colCount: number,
     *   cells: Array<object>, rowOps: Array<object>, merges: Array<object>,
     *   styleOps: Array<object>, bounds?: object, originals: string[][],
     *   tableOriginals?: Object<string, string[][]>}}
     */
    function toTablePatch() {
        /** @type {Array<object>} */
        const cells = [];
        /** @type {Array<object>} */
        const allRowOps = [];
        /** @type {Array<object>} */
        const allMerges = [];
        /** @type {Array<object>} */
        const allStyleOps = [];
        for (const draft of drafts) {
            const draftCells = [...draft.cellEdits.values()]
                .filter((c) => !draft.deletedRows().has(c.row))
                .map((c) => ({ tableIndex: draft.tableIndex, ...c }));
            cells.push(...draftCells);
            allRowOps.push(...draft.rowOps.map((o) => ({ tableIndex: draft.tableIndex, ...o })));
            allMerges.push(...draft.mergeOps.map((m) => ({ tableIndex: draft.tableIndex, ...m })));
            allStyleOps.push(...draft.styleOps);
        }
        cells.sort((a, b) => (a.tableIndex - b.tableIndex) || (a.row - b.row) || (a.col - b.col));
        // Single-table sessions keep the legacy patch shape (no tableIndex
        // anywhere) so the selection path and its apply/review code stay
        // source-compatible; multi-table sessions carry tableIndex on every
        // element — the apply side anchors them per table.
        if (drafts.length === 1) {
            const strip = (list) => list.map((op) => {
                const { tableIndex: _ti, ...rest } = op;
                return rest;
            });
            return {
                tableCount: 1,
                rowCount: drafts[0].rowCount,
                colCount: drafts[0].colCount,
                cells: strip(cells),
                rowOps: planRowOpOrder(allRowOps).map((op) => {
                    const { tableIndex: _ti, ...rest } = op;
                    return rest;
                }),
                merges: strip(allMerges),
                styleOps: strip(allStyleOps),
                bounds: drafts[0].bounds,
                originals: drafts[0].values,
            };
        }
        return {
            tableCount: drafts.length,
            rowCount: drafts[0].rowCount,
            colCount: drafts[0].colCount,
            cells,
            rowOps: planRowOpOrder(allRowOps),
            merges: allMerges,
            styleOps: allStyleOps,
            bounds: drafts[0].bounds,
            originals: drafts[0].values,
            tableOriginals: Object.fromEntries(drafts.map((d) => [d.tableIndex, d.values])),
        };
    }

    function opCount() {
        return drafts.reduce((sum, d) => sum + d.cellEdits.size + d.rowOps.length + d.mergeOps.length + d.styleOps.length, 0);
    }

    return {
        getState,
        setCell,
        insertRow,
        deleteRow,
        mergeCells,
        setTableStyle,
        setBorders,
        setCellFormat,
        setFont,
        setHeaderRow,
        setLayout,
        setColumnWidths,
        toTablePatch,
        get opCount() { return opCount(); },
    };
}

/**
 * Dispatches a table tool call against a draft model instance.
 *
 * @param {object} model - createTableModel instance
 * @param {string} name - Tool name
 * @param {object} args - Tool args
 * @returns {{ok: boolean, result?: *, error?: string}} Loop observation
 */
export function executeTableTool(model, name, args) {
    switch (name) {
        case 'get_state': return model.getState();
        case 'set_cell': return model.setCell(args.row, args.col, args.text, args.tableIndex);
        case 'insert_row': return model.insertRow(args || {});
        case 'delete_row': return model.deleteRow(args.row, args.tableIndex);
        case 'merge_cells': return model.mergeCells(args || {});
        case 'set_table_style': return model.setTableStyle(args || {});
        case 'set_borders': return model.setBorders(args || {});
        case 'set_cell_format': return model.setCellFormat(args || {});
        case 'set_font': return model.setFont(args || {});
        case 'set_header_row': return model.setHeaderRow(args || {});
        case 'set_layout': return model.setLayout(args || {});
        case 'set_column_widths': return model.setColumnWidths(args || {});
        default: return { ok: false, error: `Unknown table tool "${name}".` };
    }
}
