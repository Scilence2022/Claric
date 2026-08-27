/**
 * Table Tool Model
 *
 * L2 of the tool-calling stack: the draft model table tools operate on.
 * The model is seeded from a readSelectionTableRegion result and NEVER
 * touches Word — ops accumulate as a transaction, translated at the end
 * into the existing coordinate patch shape (table-patch.js), so the
 * proposal card and applySelectionAmendment are reused unchanged.
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

/**
 * Tool specs for the table draft model. The descriptions are the contract
 * the model sees — constraints live here, enforcement lives in the model.
 */
export const TABLE_TOOL_SPECS = Object.freeze([
    defineTool({
        name: 'get_state',
        description: 'Re-read the covered table region: the coordinate grid (merged slots are marked read-only), the pending operations recorded so far, and whether row operations are allowed.',
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
]);

/**
 * Creates the table draft model.
 *
 * @param {object} region - Result of readSelectionTableRegion:
 *   { rowCount, colCount, values (string[][]), bounds {startRow,endRow,startCol,endCol},
 *     merged?: boolean, shadowKeys?: Set<string> ("row,col" 1-based) }
 * @returns {{getState: Function, setCell: Function, insertRow: Function,
 *   deleteRow: Function, mergeCells: Function, toTablePatch: Function, opCount: number}}
 */
export function createTableModel(region) {
    const rowCount = region.rowCount;
    const colCount = region.colCount;
    const values = region.values || [];
    const bounds = region.bounds || { startRow: 1, endRow: rowCount, startCol: 1, endCol: colCount };
    const merged = !!region.merged;
    const shadowKeys = region.shadowKeys instanceof Set ? region.shadowKeys : new Set();
    const allowRowOps = !merged && bounds.startCol === 1 && bounds.endCol === colCount;

    /** @type {Map<string, {row: number, col: number, text: string}>} */
    const cellEdits = new Map();
    /** @type {Array<{op: string, row: number, values?: string[]}>} */
    const rowOps = [];
    /** @type {Array<{op: string, startRow: number, startCol: number, endRow: number, endCol: number}>} */
    const mergeOps = [];
    /** Cells swallowed by a staged merge (1-based "row,col") — read-only. */
    const mergedAway = new Set();
    const deletedRows = () => new Set(rowOps.filter((o) => o.op === 'delete').map((o) => o.row));
    const rowOpsBlocked = () => mergeOps.length > 0;

    const _err = (error) => ({ ok: false, error });
    const _ok = (result) => ({ ok: true, result });

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
            pendingOps: [
                ...[...cellEdits.values()].map((e) => ({ tool: 'set_cell', ...e })),
                ...rowOps,
                ...mergeOps.map((m) => ({ tool: 'merge_cells', ...m })),
            ],
        });
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
     * Translates the recorded ops into the tablePatch shape consumed by the
     * existing proposal card and applySelectionAmendment.
     *
     * @returns {{rowCount: number, colCount: number,
     *   cells: Array<{row: number, col: number, text: string}>,
     *   rowOps: Array<{op: string, row: number, values?: string[]}>,
     *   merges: Array<{op: string, startRow: number, startCol: number, endRow: number, endCol: number}>,
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
        toTablePatch,
        get opCount() { return cellEdits.size + rowOps.length + mergeOps.length; },
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
        default: return { ok: false, error: `Unknown table tool "${name}".` };
    }
}
