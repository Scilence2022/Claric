/**
 * Table Patch Protocol
 *
 * Coordinate-addressed revision protocol for selections that span multiple
 * table cells. Flattening a multi-cell selection into plain text breaks the
 * word/char diff strategies (cell marks land in the text, token mapping
 * diverges across cell boundaries), so table selections take a different
 * route:
 *
 *   prepare: cells are listed with absolute 1-based coordinates (R2C3) and
 *     the LLM answers with a JSON delta — only changed cells plus optional
 *     row-level structure ops — instead of a rewritten text blob.
 *   apply: each changed cell is revised individually with the existing
 *     granular diff strategies scoped to the cell (Word tracks in-cell text
 *     edits natively); row ops use TableRow.insertRows()/delete().
 *
 * This module is the pure half of that protocol: prompt text, response
 * parsing/validation, and row-op ordering. No Office.js, no network —
 * hermetic-testable. The Word-bound detection/apply lives in word-actions.
 *
 * @module table-patch
 */

/** Row-op verbs understood by the protocol. */
export const ROW_OP = Object.freeze({
    DELETE: 'delete',
    INSERT_AFTER: 'insertAfter',
    INSERT_BEFORE: 'insertBefore',
});

/**
 * Formats covered cells as the coordinate listing embedded in the prompt.
 * One line per cell: `[R{row}C{col}] {text}` (1-based absolute coordinates).
 *
 * @param {Array<{row: number, col: number, text: string}>} cells
 * @returns {string}
 */
export function formatCellGrid(cells) {
    return cells.map((c) => `[R${c.row}C${c.col}] ${c.text}`).join('\n');
}

/**
 * Builds the user message for a table-scope amendment: the user's edit
 * instruction, the coordinate grid of current cell contents, and the JSON
 * patch output rules.
 *
 * @param {string} instruction - The amendment instruction (skill template or
 *   free text; may contain a {selection} placeholder, already NOT substituted
 *   — the grid below plays that role)
 * @param {Array<{row: number, col: number, text: string}>} cells - Covered cells
 * @param {{rowCount: number, colCount: number}} dims - Table dimensions
 * @returns {string}
 */
export function buildTableUserPrompt(instruction, cells, { rowCount, colCount }) {
    return `${instruction}

The selection covers a table region in a Word document (${rowCount} rows x ${colCount} columns total). Current contents of the covered cells (1-based absolute coordinates):

${formatCellGrid(cells)}

CRITICAL OUTPUT RULES (table mode):
- Respond with ONLY a JSON object — no commentary, no markdown code fences:
  {"cells":[{"row":2,"col":1,"text":"new cell text"}],"rowOps":[{"op":"delete","row":3},{"op":"insertAfter","row":4,"values":["cell A","cell B"]}]}
- "cells": list ONLY cells whose text you change. "row"/"col" must be coordinates from the listing above. "text" is the FULL new text of that cell (plain text, no markdown).
- "rowOps": list ONLY structural changes. Supported ops: "delete" (remove the row), "insertAfter" / "insertBefore" (insert one new row next to the given row, "values" = the new row's cell texts, exactly ${colCount} entries).
- If only cell text changes, omit "rowOps" or use an empty array. If nothing should change, respond with {"cells":[],"rowOps":[]}.
- Do NOT rewrite the table as text or markdown. Do NOT include unchanged cells.`;
}

/**
 * Parses and validates the LLM's table-patch JSON.
 *
 * Lenient about transport noise (markdown fences, trailing commas) but strict
 * about semantics: out-of-bounds coordinates and malformed ops are dropped
 * with warnings rather than applied. Cell entries whose text matches the
 * current cell (ignoring surrounding whitespace) are dropped as no-ops so
 * they cannot create meaningless whitespace-only revisions.
 *
 * @param {string} raw - Raw LLM response
 * @param {object} shape
 * @param {number} shape.rowCount - Total table rows (bounds for row coords)
 * @param {number} shape.colCount - Total table columns (bounds for col coords
 *   and insert-values length)
 * @param {string[][]} [shape.originals] - Current cell texts (0-based [r][c])
 *   for no-op detection; omit to keep every parsed cell entry
 * @returns {{cells: Array<{row: number, col: number, text: string}>,
 *   rowOps: Array<{op: string, row: number, values?: string[]}>,
 *   warnings: string[]}}
 *   cells are sorted top-left → bottom-right (document order);
 *   rowOps are pre-sorted into application order via planRowOpOrder.
 * @throws {Error} When no parseable JSON object is present
 */
// @ts-expect-error - the `= {}` default is a defensive fallback; the only caller
// (word-actions.js) always passes the full shape. Removing it would change the
// graceful-degradation behavior for missing bounds.
export function parseTablePatchResponse(raw, { rowCount, colCount, originals } = {}) {
    const warnings = [];
    const parsed = _extractJson(raw);

    const cells = _parseCells(parsed.cells, { rowCount, colCount, originals, warnings });
    const rowOps = planRowOpOrder(_parseRowOps(parsed.rowOps, { rowCount, colCount, warnings }));

    return { cells, rowOps, warnings };
}

/**
 * Orders row ops for application. Coordinates refer to the ORIGINAL table,
 * so ops must run against descending row indexes — inserting/deleting a row
 * shifts every row below it, and descending order keeps not-yet-applied
 * coordinates valid. Ties on the same row resolve inserts-before-deletes,
 * which yields "replace row" semantics for insertAfter+delete pairs.
 * Pure function — exported for tests.
 *
 * @param {Array<{op: string, row: number}>} rowOps
 * @returns {Array<{op: string, row: number}>} New sorted array
 */
export function planRowOpOrder(rowOps) {
    const rank = (op) => (op === ROW_OP.DELETE ? 1 : 0);
    return [...rowOps].sort((a, b) => (b.row - a.row) || (rank(a.op) - rank(b.op)));
}

/**
 * Extracts the JSON object from an LLM response, tolerating markdown code
 * fences and trailing commas (both common with smaller models).
 *
 * @param {string} raw
 * @returns {object}
 * @throws {Error} When no JSON object can be located or parsed
 * @private
 */
function _extractJson(raw) {
    const text = String(raw || '');
    const start = text.indexOf('{');
    const end = text.lastIndexOf('}');
    if (start === -1 || end <= start) {
        throw new Error('Table patch response contains no JSON object');
    }
    const cleaned = text.slice(start, end + 1).replace(/,\s*([}\]])/g, '$1');
    try {
        const parsed = JSON.parse(cleaned);
        if (!parsed || typeof parsed !== 'object' || Array.isArray(parsed)) {
            throw new Error('not an object');
        }
        return parsed;
    } catch (err) {
        throw new Error(`Table patch JSON parse failed: ${err.message}`);
    }
}

/** @private */
function _parseCells(rawCells, { rowCount, colCount, originals, warnings }) {
    if (rawCells === undefined || rawCells === null) return [];
    if (!Array.isArray(rawCells)) {
        warnings.push('"cells" is not an array — ignored');
        return [];
    }

    const byKey = new Map();
    for (const entry of rawCells) {
        const row = _asInt(entry && entry.row);
        const col = _asInt(entry && entry.col);
        if (!row || !col || row < 1 || row > rowCount || col < 1 || col > colCount) {
            warnings.push(`Cell coordinate out of bounds (row=${entry && entry.row}, col=${entry && entry.col}) — dropped`);
            continue;
        }
        const text = _asText(entry.text);
        if (text === null) {
            warnings.push(`Cell R${row}C${col} has no usable "text" — dropped`);
            continue;
        }
        const key = `${row},${col}`;
        if (byKey.has(key)) {
            warnings.push(`Duplicate entry for R${row}C${col} — last one wins`);
        }
        byKey.set(key, { row, col, text });
    }

    const cells = [...byKey.values()].sort((a, b) => (a.row - b.row) || (a.col - b.col));

    if (!originals) return cells;
    return cells.filter((c) => {
        const original = (originals[c.row - 1] && originals[c.row - 1][c.col - 1]) || '';
        return original.trim() !== c.text.trim();
    });
}

/** @private */
function _parseRowOps(rawOps, { rowCount, colCount, warnings }) {
    if (rawOps === undefined || rawOps === null) return [];
    if (!Array.isArray(rawOps)) {
        warnings.push('"rowOps" is not an array — ignored');
        return [];
    }

    const ops = [];
    const validOps = [ROW_OP.DELETE, ROW_OP.INSERT_AFTER, ROW_OP.INSERT_BEFORE];
    for (const entry of rawOps) {
        const op = entry && entry.op;
        const row = _asInt(entry && entry.row);
        if (!validOps.includes(op)) {
            warnings.push(`Unknown row op "${op}" — dropped (use delete | insertAfter | insertBefore)`);
            continue;
        }
        if (!row || row < 1 || row > rowCount) {
            warnings.push(`Row op "${op}" has out-of-bounds row=${entry && entry.row} — dropped`);
            continue;
        }
        if (op === ROW_OP.DELETE) {
            ops.push({ op, row });
            continue;
        }
        const values = _asRowValues(entry.values, colCount, row, warnings);
        if (values === null) continue;
        ops.push({ op, row, values });
    }
    return ops;
}

/** @private */
function _asRowValues(rawValues, colCount, row, warnings) {
    if (!Array.isArray(rawValues)) {
        warnings.push(`Row op at row ${row} has no "values" array — dropped`);
        return null;
    }
    const values = rawValues.slice(0, colCount).map((v) => _asText(v) ?? '');
    if (rawValues.length > colCount) {
        warnings.push(`Row op at row ${row} has ${rawValues.length} values; table has ${colCount} columns — truncated`);
    }
    while (values.length < colCount) values.push('');
    return values;
}

/** @private */
function _asInt(value) {
    const n = typeof value === 'string' && value.trim() !== '' ? Number(value) : value;
    return Number.isInteger(n) ? n : null;
}

/** Coerces LLM cell values to text; numbers/booleans are common in numeric
 *  cells. Returns null for values that cannot become meaningful text.
 *  @private */
function _asText(value) {
    if (typeof value === 'string') return value;
    if (typeof value === 'number' || typeof value === 'boolean') return String(value);
    return null;
}
