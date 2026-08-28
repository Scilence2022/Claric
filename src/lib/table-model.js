/**
 * Table Tool Model
 *
 * L2 of the tool-calling stack: the draft model table tools operate on.
 * The model is seeded from a readSelectionTableRegion result and NEVER
 * touches Word — ops accumulate as a transaction, translated at the end
 * into the existing coordinate patch shape (table-patch.js), so the
 * proposal card and applySelectionAmendment are reused unchanged.
 *
 * Two op families: content/structure (set_cell, insert_row, delete_row,
 * merge_cells) and styling (set_table_style, set_borders, set_cell_format,
 * set_font, set_header_row, set_layout, set_column_widths — validated by
 * lib/table-style.js). Style ops with a `region` are coordinate-bound and
 * apply before row structure changes; table-level style ops apply after.
 *
 * Coordinate discipline (mirrors the apply path): cell coordinates always
 * refer to ORIGINAL row numbering during the loop — row ops are queued,
 * never shift cell coordinates mid-loop. Insert values must therefore be
 * final. This keeps the translated patch expressible in the existing
 * descending-row-op apply order.
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
 */
export const TABLE_TOOL_SPECS = Object.freeze([
    defineTool({
        name: 'get_state',
        description: 'Re-read the covered table region: the coordinate grid (merged slots are marked read-only), the current table style (borders, shading, alignments, header rows, font), the pending operations recorded so far, and whether row operations are allowed.',
        argsExample: {},
    }),
    defineTool({
        name: 'set_cell',
        description: 'Set the FULL new text of one cell. Coordinates are 1-based, from the grid listing, and must be inside the covered region. Slots marked "(merged — read-only)" cannot be set. Rows with a pending delete cannot be set. Setting the current text is a no-op.',
        argsExample: { row: 2, col: 1, text: 'new text' },
    }),
    defineTool({
        name: 'insert_row',
        description: 'Insert one new row next to an existing row. "values" must contain the FINAL text of every cell in the new row (exactly colCount entries, plain text). Only allowed when row operations are enabled (uniform table, full-width selection).',
        argsExample: { position: 'after', row: 3, values: ['cell A', 'cell B'] },
    }),
    defineTool({
        name: 'delete_row',
        description: 'Delete one existing row by its 1-based row number. Only allowed when row operations are enabled (uniform table, full-width selection). The table must keep at least one row.',
        argsExample: { row: 4 },
    }),
    defineTool({
        name: 'merge_cells',
        description: 'Merge a rectangular block of cells into one cell. "row"/"col" is the top-left cell (1-based, from the grid); "rows"/"cols" is the block size (the block must span 2+ cells). All cells in the block must be editable (not already merged away, not in a pending-delete row). The top-left cell becomes the anchor: set its text before merging to control the final merged content (the other cells are cleared and merged into it). Only one merge can be staged at a time, and staging it disables further row operations.',
        argsExample: { row: 1, col: 1, rows: 2, cols: 2 },
    }),
    defineTool({
        name: 'set_table_style',
        description: 'Set the overall LOOK of the whole table: a built-in Word table style plus its banding/emphasis options. "style" must be a built-in style name: TableGrid, TableGridLight, PlainTable1..PlainTable5, or GridTable/ListTable families (GridTable1Light, GridTable2, GridTable3, GridTable4, GridTable5Dark, GridTable6Colorful, GridTable7Colorful, ListTable1Light, ListTable2, ..., ListTable7Colorful), each optionally suffixed _Accent1.._Accent6. Optional boolean flags tune the style: bandedRows, bandedColumns, firstColumn, lastColumn, totalRow. To remove or draw specific borders use set_borders instead.',
        argsExample: { style: 'GridTable4_Accent1', bandedRows: true, firstColumn: true },
    }),
    defineTool({
        name: 'set_borders',
        description: 'Set table borders (whole table), or one row\'s borders when "row" is given. "borders" keys: top, bottom, left, right, insideH, insideV, plus shorthands all / outside / inside. Each value is a border type string ("none", "single", "double", "dotted", "dashed", "dotDashed", "triple", "wave", ...) or {type, color?, width?} with color "#RRGGBB"/"auto"/a color name and width in points. Academic three-line table: set_borders({borders:{top:{type:"single",width:1.5}, bottom:{type:"single",width:1.5}, inside:"none"}}) then set_borders({row:1, borders:{bottom:{type:"single",width:0.75}}}).',
        argsExample: { borders: { top: { type: 'single', width: 1.5 }, bottom: { type: 'single', width: 1.5 }, inside: 'none' } },
    }),
    defineTool({
        name: 'set_cell_format',
        description: 'Set cell shading (background color) and content alignment. Target: {row, col} one cell, {row, rows} a row band, {col, cols} a column band, {row, col, rows, cols} a block, or omit target for the whole table. Payload keys: shadingColor ("#RRGGBB", "auto", or a color name), horizontalAlignment (left | centered | right | justified), verticalAlignment (top | center | bottom). Coordinates covered by a merge resolve to the merged cell. Cannot target rows with a pending delete.',
        argsExample: { row: 1, shadingColor: '#DEEBF7', horizontalAlignment: 'centered' },
    }),
    defineTool({
        name: 'set_font',
        description: 'Set font properties. Same targeting as set_cell_format ({row?, col?, rows?, cols?}; omit for the whole table). "font": {bold?, italic?, underline? ("none"|"single"|"double"|...), size? (points), name?, color?}. Example header emphasis: {row: 1, font: {bold: true}}.',
        argsExample: { row: 1, font: { bold: true, size: 11 } },
    }),
    defineTool({
        name: 'set_header_row',
        description: 'Mark the first N rows as the table header row(s) — Word repeats them on every page — and optionally style them. args: {rows? (default 1), font?, shadingColor?}. Only allowed when the covered region includes row 1.',
        argsExample: { rows: 1, font: { bold: true }, shadingColor: '#DEEBF7' },
    }),
    defineTool({
        name: 'set_layout',
        description: 'Table-level layout: alignment of the table against the page column ("left" | "centered" | "right"), widthPt (total width in points), autoFitWindow (true = stretch columns to the window), distributeColumns (true = equal column widths), cellPaddingPt (uniform cell padding on all sides, 0–100pt).',
        argsExample: { alignment: 'centered', autoFitWindow: true },
    }),
    defineTool({
        name: 'set_column_widths',
        description: 'Set per-column widths in points: {widthsPt: [...]} with exactly colCount entries (5–1584 each). Only allowed on uniform tables (no merged cells) — merged tables reject it.',
        argsExample: { widthsPt: [120, 80, 200] },
    }),
]);

/**
 * Creates the table draft model.
 *
 * @param {object} region - Result of readSelectionTableRegion:
 *   { rowCount, colCount, values (string[][]), bounds {startRow,endRow,startCol,endCol},
 *     merged?: boolean, shadowKeys?: Set<string> ("row,col" 1-based),
 *     style?: object|null (advisory style snapshot for get_state) }
 * @returns {{getState: Function, setCell: Function, insertRow: Function,
 *   deleteRow: Function, mergeCells: Function, setTableStyle: Function,
 *   setBorders: Function, setCellFormat: Function, setFont: Function,
 *   setHeaderRow: Function, setLayout: Function, setColumnWidths: Function,
 *   toTablePatch: Function, opCount: number}}
 */
export function createTableModel(region) {
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
    /**
     * @type {Array<Record<string, *>>} Normalized style ops. Ops with a
     * `region` are coordinate-bound (applied before row structure changes);
     * the rest are table-level (applied after).
     */
    const styleOps = [];
    /** Cells swallowed by a staged merge (1-based "row,col") — read-only. */
    const mergedAway = new Set();
    const deletedRows = () => new Set(rowOps.filter((o) => o.op === 'delete').map((o) => o.row));
    const rowOpsBlocked = () => mergeOps.length > 0;

    const _err = (error) => ({ ok: false, error });
    const _ok = (result) => ({ ok: true, result });

    /** Rejects style ops once the budget is exhausted. @private */
    function _checkStyleBudget() {
        if (styleOps.length >= TABLE_STYLE_LIMITS.MAX_STYLE_OPS) {
            return `Style op limit reached (${TABLE_STYLE_LIMITS.MAX_STYLE_OPS}) — finish with what is staged.`;
        }
        return null;
    }

    /** True when the region intersects any pending-delete row. @private */
    function _intersectsDeletedRow(region) {
        if (!region) return false;
        const deleted = deletedRows();
        for (let r = region.startRow; r <= region.endRow; r++) {
            if (deleted.has(r)) return true;
        }
        return false;
    }

    function getState() {
        const grid = [];
        for (let r = bounds.startRow; r <= bounds.endRow; r++) {
            for (let c = bounds.startCol; c <= bounds.endCol; c++) {
                const key = `${r},${c}`;
                const current = cellEdits.has(key) ? cellEdits.get(key).text : (values[r - 1] && values[r - 1][c - 1]) || '';
                grid.push(`[R${r}C${c}]${shadowKeys.has(key) || mergedAway.has(key) ? ' (merged — read-only)' : ''} ${current}`);
            }
        }
        return _ok({
            rowCount,
            colCount,
            coveredRegion: `R${bounds.startRow}C${bounds.startCol}–R${bounds.endRow}C${bounds.endCol}`,
            mergedTable: merged,
            rowOpsAllowed: allowRowOps,
            grid: grid.join('\n'),
            style: _describeStyleSnapshot(),
            pendingOps: [
                ...[...cellEdits.values()].map((e) => ({ tool: 'set_cell', ...e })),
                ...rowOps,
                ...mergeOps.map((m) => ({ tool: 'merge_cells', ...m })),
                ...styleOps,
            ],
        });
    }

    /**
     * Renders the style snapshot seeded from readSelectionTableRegion into a
     * readable summary for get_state. Missing snapshot → "unknown".
     * @private
     */
    function _describeStyleSnapshot() {
        const s = styleSnapshot;
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

    function setCell(row, col, text) {
        if (!Number.isInteger(row) || !Number.isInteger(col)
            || row < 1 || row > rowCount || col < 1 || col > colCount) {
            return _err(`Coordinates out of bounds (row=${row}, col=${col}); the table is ${rowCount}x${colCount}.`);
        }
        if (row < bounds.startRow || row > bounds.endRow || col < bounds.startCol || col > bounds.endCol) {
            return _err(`Cell R${row}C${col} is outside the covered region R${bounds.startRow}C${bounds.startCol}–R${bounds.endRow}C${bounds.endCol}.`);
        }
        if (shadowKeys.has(`${row},${col}`)) {
            return _err(`Cell R${row}C${col} is covered by a merged cell — it is read-only. Edit the merge anchor slot instead.`);
        }
        if (mergedAway.has(`${row},${col}`)) {
            return _err(`Cell R${row}C${col} is set to be merged away — it is read-only. Set the merge anchor (top-left of the block) instead.`);
        }
        if (typeof text !== 'string') {
            return _err('"text" must be a string.');
        }
        if (deletedRows().has(row)) {
            return _err(`Row ${row} has a pending delete — setting cells in it is not allowed.`);
        }
        const key = `${row},${col}`;
        const original = (values[row - 1] && values[row - 1][col - 1]) || '';
        if (text.trim() === original.trim() && !cellEdits.has(key)) {
            return _ok({ note: `R${row}C${col} already holds that text — no change.` });
        }
        cellEdits.set(key, { row, col, text });
        return _ok({ set: `R${row}C${col}` });
    }

    function insertRow({ position, row, values: newValues }) {
        if (merged) {
            return _err('Row operations are not allowed: the table contains merged cells.');
        }
        if (rowOpsBlocked()) {
            return _err('Row operations are not allowed while a cell merge is pending — apply the merge or clear it first.');
        }
        if (!allowRowOps) {
            return _err('Row operations require a full-width selection covering every column.');
        }
        if (position !== 'after' && position !== 'before') {
            return _err('"position" must be "after" or "before".');
        }
        if (!Number.isInteger(row) || row < 1 || row > rowCount) {
            return _err(`Anchor row out of bounds (row=${row}); the table has ${rowCount} rows.`);
        }
        if (row < bounds.startRow || row > bounds.endRow) {
            return _err(`Anchor row ${row} is outside the covered region.`);
        }
        if (!Array.isArray(newValues) || newValues.length !== colCount) {
            return _err(`"values" must be an array of exactly ${colCount} strings (one per column).`);
        }
        if (!newValues.every((v) => typeof v === 'string')) {
            return _err('Every entry in "values" must be a string.');
        }
        const op = position === 'after' ? 'insertAfter' : 'insertBefore';
        rowOps.push({ op, row, values: [...newValues] });
        return _ok({ inserted: `${op} row ${row}` });
    }

    function deleteRow(row) {
        if (merged) {
            return _err('Row operations are not allowed: the table contains merged cells.');
        }
        if (rowOpsBlocked()) {
            return _err('Row operations are not allowed while a cell merge is pending — apply the merge or clear it first.');
        }
        if (!allowRowOps) {
            return _err('Row operations require a full-width selection covering every column.');
        }
        if (!Number.isInteger(row) || row < 1 || row > rowCount) {
            return _err(`Row out of bounds (row=${row}); the table has ${rowCount} rows.`);
        }
        if (row < bounds.startRow || row > bounds.endRow) {
            return _err(`Row ${row} is outside the covered region.`);
        }
        if (deletedRows().size + 1 >= rowCount) {
            return _err('Deleting this row would remove every row of the table — not allowed.');
        }
        if (deletedRows().has(row)) {
            return _ok({ note: `Row ${row} already has a pending delete.` });
        }
        rowOps.push({ op: 'delete', row });
        return _ok({ deleted: `row ${row}` });
    }

    /**
     * Stages a rectangular cell merge. The top-left cell (\`row\`,\`col\`) is
     * the anchor that holds the merged content; every other cell in the block
     * is cleared and merged into it at apply time. Only one merge may be
     * staged at a time, and staging it blocks further row operations (a merge
     * makes the grid non-uniform).
     *
     * @param {{row: number, col: number, rows: number, cols: number}} args
     * @returns {{ok: boolean, result?: *, error?: string}}
     */
    function mergeCells({ row, col, rows, cols }) {
        if (!Number.isInteger(row) || !Number.isInteger(col) || !Number.isInteger(rows) || !Number.isInteger(cols)) {
            return _err('"row", "col", "rows", "cols" must be integers.');
        }
        if (rows < 1 || cols < 1) {
            return _err('"rows" and "cols" must be ≥ 1.');
        }
        const endRow = row + rows - 1;
        const endCol = col + cols - 1;
        if (row < 1 || endRow > rowCount || col < 1 || endCol > colCount) {
            return _err(`Merge region R${row}C${col}–R${endRow}C${endCol} is outside the ${rowCount}x${colCount} table.`);
        }
        if (row < bounds.startRow || endRow > bounds.endRow || col < bounds.startCol || endCol > bounds.endCol) {
            return _err(`Merge region R${row}C${col}–R${endRow}C${endCol} is outside the covered region R${bounds.startRow}C${bounds.startCol}–R${bounds.endRow}C${bounds.endCol}.`);
        }
        if (rows * cols < 2) {
            return _err('Merging a single cell is a no-op — the block must span 2+ cells.');
        }
        if (rowOpsBlocked()) {
            return _err('Only one cell merge can be staged at a time.');
        }
        for (let r = row; r <= endRow; r++) {
            for (let c = col; c <= endCol; c++) {
                const key = `${r},${c}`;
                if (shadowKeys.has(key)) {
                    return _err(`Cell R${r}C${c} is covered by an existing merged cell — cannot merge.`);
                }
                if (mergedAway.has(key)) {
                    return _err(`Cell R${r}C${c} is already part of a staged merge.`);
                }
                if (deletedRows().has(r)) {
                    return _err(`Row ${r} has a pending delete — cannot merge.`);
                }
                if (r !== row || c !== col) mergedAway.add(key);
            }
        }
        mergeOps.push({ op: 'merge', startRow: row, startCol: col, endRow, endCol });
        return _ok({ merged: `R${row}C${col}–R${endRow}C${endCol} (anchor R${row}C${col})` });
    }

    /**
     * Stages the whole-table look: built-in style plus banding/emphasis flags.
     *
     * @param {{style?: string, bandedRows?: boolean, bandedColumns?: boolean,
     *   firstColumn?: boolean, lastColumn?: boolean, totalRow?: boolean}} args
     */
    function setTableStyle(args = {}) {
        const budgetError = _checkStyleBudget();
        if (budgetError) return _err(budgetError);
        /** @type {Record<string, *>} */
        const op = { type: 'tableStyle', tool: 'set_table_style' };
        if (args.style !== undefined && args.style !== null) {
            if (typeof args.style !== 'string') return _err('"style" must be a built-in table style name string.');
            const canonical = BUILT_IN_TABLE_STYLES.find((name) =>
                name.replace(/[\s_-]+/g, '').toLowerCase() === args.style.replace(/[\s_-]+/g, '').toLowerCase());
            if (!canonical) {
                return _err(`Unknown built-in table style "${args.style}". Use e.g. TableGrid, TableGridLight, PlainTable1..5, or a GridTable/ListTable family (optionally _Accent1..6).`);
            }
            op.style = canonical;
        }
        for (const key of ['bandedRows', 'bandedColumns', 'firstColumn', 'lastColumn', 'totalRow']) {
            if (args[key] === undefined || args[key] === null) continue;
            if (typeof args[key] !== 'boolean') return _err(`"${key}" must be true or false.`);
            op[key] = args[key];
        }
        if (Object.keys(op).length <= 2) return _err('Give at least one of style, bandedRows, bandedColumns, firstColumn, lastColumn, totalRow.');
        styleOps.push(op);
        return _ok({ staged: `table style${op.style ? ` ${op.style}` : ''}` });
    }

    /**
     * Stages border changes for the whole table, or for one row's cells when
     * `row` is given (a header separator line is a row bottom border).
     *
     * @param {{row?: number, borders?: object}} args
     */
    function setBorders(args = {}) {
        const budgetError = _checkStyleBudget();
        if (budgetError) return _err(budgetError);
        const { borders } = args || {};
        const expanded = expandBorderSet(borders);
        if (expanded.error) return _err(expanded.error);
        /** @type {Record<string, *>} */
        const op = { type: 'borders', tool: 'set_borders', borders: expanded.borders };
        if (args.row !== undefined && args.row !== null) {
            if (!Number.isInteger(args.row) || args.row < 1 || args.row > rowCount) {
                return _err(`Row out of bounds (row=${args.row}); the table has ${rowCount} rows.`);
            }
            if (args.row < bounds.startRow || args.row > bounds.endRow) {
                return _err(`Row ${args.row} is outside the covered region.`);
            }
            if (deletedRows().has(args.row)) {
                return _err(`Row ${args.row} has a pending delete — cannot style it.`);
            }
            op.row = args.row;
        }
        styleOps.push(op);
        return _ok({ staged: `${op.row ? `row ${op.row} ` : ''}borders: ${Object.keys(op.borders).join(', ')}` });
    }

    /**
     * Stages cell shading / alignment for a target region (or the whole
     * table when no coordinates are given).
     *
     * @param {object} args - {row?, col?, rows?, cols?, shadingColor?,
     *   horizontalAlignment?, verticalAlignment?}
     */
    function setCellFormat(args = {}) {
        const budgetError = _checkStyleBudget();
        if (budgetError) return _err(budgetError);
        const target = normalizeStyleTarget(_pickTarget(args), bounds);
        if (target.error) return _err(target.error);
        if (_intersectsDeletedRow(target.region)) {
            return _err('Cannot format rows with a pending delete.');
        }
        /** @type {Record<string, *>} */
        const op = { type: 'cellFormat', tool: 'set_cell_format', region: target.region };
        if (args.shadingColor !== undefined && args.shadingColor !== null) {
            const color = normalizeColor(args.shadingColor);
            if (!color) return _err(`Invalid shadingColor "${args.shadingColor}" (use "#RRGGBB", "auto", or a color name).`);
            op.shadingColor = color;
        }
        if (args.horizontalAlignment !== undefined && args.horizontalAlignment !== null) {
            const align = normalizeAlignment(args.horizontalAlignment, HORIZONTAL_ALIGNMENTS);
            if (!align) return _err('"horizontalAlignment" must be left | centered | right | justified.');
            op.horizontalAlignment = align;
        }
        if (args.verticalAlignment !== undefined && args.verticalAlignment !== null) {
            const align = normalizeAlignment(args.verticalAlignment, VERTICAL_ALIGNMENTS);
            if (!align) return _err('"verticalAlignment" must be top | center | bottom.');
            op.verticalAlignment = align;
        }
        if (!op.shadingColor && !op.horizontalAlignment && !op.verticalAlignment) {
            return _err('Give at least one of shadingColor, horizontalAlignment, verticalAlignment.');
        }
        styleOps.push(op);
        return _ok({ staged: `format ${_regionLabel(op.region)}` });
    }

    /**
     * Stages font properties for a target region (or the whole table).
     *
     * @param {object} args - {row?, col?, rows?, cols?, font}
     */
    function setFont(args = {}) {
        const budgetError = _checkStyleBudget();
        if (budgetError) return _err(budgetError);
        const target = normalizeStyleTarget(_pickTarget(args), bounds);
        if (target.error) return _err(target.error);
        if (_intersectsDeletedRow(target.region)) {
            return _err('Cannot set fonts on rows with a pending delete.');
        }
        const font = normalizeFontPayload(args.font);
        if (font.error) return _err(font.error);
        styleOps.push({ type: 'font', tool: 'set_font', region: target.region, font: font.font });
        return _ok({ staged: `font ${_regionLabel(target.region)}: ${Object.keys(font.font).join(', ')}` });
    }

    /**
     * Stages the table-header setting: headerRowCount plus optional header
     * row formatting. Header rows are the table's FIRST rows, so the covered
     * region must include row 1.
     *
     * @param {{rows?: number, font?: object, shadingColor?: string}} args
     */
    function setHeaderRow(args = {}) {
        const budgetError = _checkStyleBudget();
        if (budgetError) return _err(budgetError);
        if (bounds.startRow !== 1) {
            return _err('Header rows are the table\'s first rows — the covered region must include row 1.');
        }
        const rows = args.rows === undefined || args.rows === null ? 1 : args.rows;
        if (!Number.isInteger(rows) || rows < 1 || rows > bounds.endRow) {
            return _err(`"rows" must be 1–${bounds.endRow}.`);
        }
        if (_intersectsDeletedRow({ startRow: 1, endRow: rows, startCol: 1, endCol: colCount })) {
            return _err('Cannot make a row with a pending delete a header.');
        }
        /** @type {Record<string, *>} */
        const op = { type: 'headerRow', tool: 'set_header_row', rows };
        if (args.font !== undefined && args.font !== null) {
            const font = normalizeFontPayload(args.font);
            if (font.error) return _err(font.error);
            op.font = font.font;
        }
        if (args.shadingColor !== undefined && args.shadingColor !== null) {
            const color = normalizeColor(args.shadingColor);
            if (!color) return _err(`Invalid shadingColor "${args.shadingColor}".`);
            op.shadingColor = color;
        }
        styleOps.push(op);
        return _ok({ staged: `header row(s) 1–${rows}` });
    }

    /**
     * Stages table-level layout: page alignment, width, autofit, column
     * distribution, cell padding.
     *
     * @param {{alignment?: string, widthPt?: number, autoFitWindow?: boolean,
     *   distributeColumns?: boolean, cellPaddingPt?: number}} args
     */
    function setLayout(args = {}) {
        const budgetError = _checkStyleBudget();
        if (budgetError) return _err(budgetError);
        /** @type {Record<string, *>} */
        const op = { type: 'layout', tool: 'set_layout' };
        if (args.alignment !== undefined && args.alignment !== null) {
            const align = normalizeAlignment(args.alignment, ['left', 'centered', 'right']);
            if (!align) return _err('"alignment" must be left | centered | right.');
            op.alignment = align;
        }
        if (args.widthPt !== undefined && args.widthPt !== null) {
            const width = Number(args.widthPt);
            if (!Number.isFinite(width) || width < 10 || width > TABLE_STYLE_LIMITS.MAX_TABLE_WIDTH_PT) {
                return _err(`"widthPt" must be 10–${TABLE_STYLE_LIMITS.MAX_TABLE_WIDTH_PT} points.`);
            }
            op.widthPt = Math.round(width);
        }
        for (const key of ['autoFitWindow', 'distributeColumns']) {
            if (args[key] === undefined || args[key] === null) continue;
            if (typeof args[key] !== 'boolean') return _err(`"${key}" must be true or false.`);
            op[key] = args[key];
        }
        if (args.cellPaddingPt !== undefined && args.cellPaddingPt !== null) {
            const padding = Number(args.cellPaddingPt);
            if (!Number.isFinite(padding) || padding < TABLE_STYLE_LIMITS.MIN_CELL_PADDING_PT
                || padding > TABLE_STYLE_LIMITS.MAX_CELL_PADDING_PT) {
                return _err(`"cellPaddingPt" must be ${TABLE_STYLE_LIMITS.MIN_CELL_PADDING_PT}–${TABLE_STYLE_LIMITS.MAX_CELL_PADDING_PT} points.`);
            }
            op.cellPaddingPt = Math.round(padding * 10) / 10;
        }
        if (Object.keys(op).length <= 2) {
            return _err('Give at least one of alignment, widthPt, autoFitWindow, distributeColumns, cellPaddingPt.');
        }
        styleOps.push(op);
        return _ok({ staged: `layout: ${Object.keys(op).filter((k) => k !== 'type' && k !== 'tool').join(', ')}` });
    }

    /**
     * Stages per-column widths. Uniform tables only — merged grids have no
     * well-defined columns.
     *
     * @param {{widthsPt?: number[]}} args
     */
    function setColumnWidths(args = {}) {
        const budgetError = _checkStyleBudget();
        if (budgetError) return _err(budgetError);
        if (merged) {
            return _err('Column widths are not supported: the table contains merged cells.');
        }
        const widths = args && args.widthsPt;
        if (!Array.isArray(widths) || widths.length !== colCount) {
            return _err(`"widthsPt" must be an array of exactly ${colCount} numbers (one per column).`);
        }
        const normalized = [];
        for (const value of widths) {
            const width = Number(value);
            if (!Number.isFinite(width) || width < TABLE_STYLE_LIMITS.MIN_COLUMN_WIDTH_PT
                || width > TABLE_STYLE_LIMITS.MAX_COLUMN_WIDTH_PT) {
                return _err(`Every width must be ${TABLE_STYLE_LIMITS.MIN_COLUMN_WIDTH_PT}–${TABLE_STYLE_LIMITS.MAX_COLUMN_WIDTH_PT}pt.`);
            }
            normalized.push(Math.round(width));
        }
        styleOps.push({ type: 'columnWidths', tool: 'set_column_widths', widthsPt: normalized });
        return _ok({ staged: `column widths: ${normalized.join('/')}pt` });
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
     * existing proposal card and applySelectionAmendment.
     *
     * @returns {{rowCount: number, colCount: number,
     *   cells: Array<{row: number, col: number, text: string}>,
     *   rowOps: Array<{op: string, row: number, values?: string[]}>,
     *   merges: Array<{op: string, startRow: number, startCol: number, endRow: number, endCol: number}>,
     *   styleOps: Array<object>,
     *   bounds: object, originals: string[][]}}
     */
    function toTablePatch() {
        const cells = [...cellEdits.values()]
            .filter((c) => !deletedRows().has(c.row))
            .sort((a, b) => (a.row - b.row) || (a.col - b.col));
        return {
            rowCount,
            colCount,
            cells,
            rowOps: planRowOpOrder(rowOps),
            merges: [...mergeOps],
            styleOps: [...styleOps],
            bounds,
            originals: values,
        };
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
        get opCount() { return cellEdits.size + rowOps.length + mergeOps.length + styleOps.length; },
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
        case 'set_cell': return model.setCell(args.row, args.col, args.text);
        case 'insert_row': return model.insertRow(args || {});
        case 'delete_row': return model.deleteRow(args.row);
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
