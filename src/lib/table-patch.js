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

/** Hard limits for untrusted model output accepted by the patch parser. */
export const TABLE_PATCH_LIMITS = Object.freeze({
    MAX_RESPONSE_CHARS: 256 * 1024,
    MAX_CELL_ENTRIES: 512,
    MAX_ROW_OPS: 128,
    MAX_CELL_TEXT_CHARS: 16 * 1024,
    MAX_TOTAL_TEXT_CHARS: 128 * 1024,
});

/**
 * Formats covered cells as the coordinate listing embedded in the prompt.
 * One line per cell: `[R{row}C{col}] {text}` (1-based absolute coordinates).
 * Cells covered by a merged region carry a read-only marker.
 *
 * @param {Array<{row: number, col: number, text: string, merged?: boolean}>} cells
 * @returns {string}
 */
export function formatCellGrid(cells) {
    return cells.map((c) => `[R${c.row}C${c.col}]${c.merged ? ' (merged — read-only)' : ''} ${c.text}`).join('\n');
}

/**
 * Builds the user message for a table-scope amendment: the user's edit
 * instruction, the coordinate grid of current cell contents, and the JSON
 * patch output rules. When any covered cell is merged-readonly, extra rules
 * restrict the patch to merge anchors and forbid row structure ops.
 *
 * @param {string} instruction - The amendment instruction (skill template or
 *   free text; may contain a {selection} placeholder, already NOT substituted
 *   — the grid below plays that role)
 * @param {Array<{row: number, col: number, text: string, merged?: boolean}>} cells - Covered cells
 * @param {{rowCount: number, colCount: number}} dims - Table dimensions
 * @returns {string}
 */
export function buildTableUserPrompt(instruction, cells, { rowCount, colCount }) {
    const mergedRules = cells.some((c) => c.merged)
        ? '\n- MERGED CELLS: entries marked "(merged — read-only)" are grid slots covered by a merged cell. NEVER include them in "cells" — their coordinates are not editable.' +
          '\n- This table contains merged cells: "rowOps" MUST be an empty array (row insert/delete is not supported on merged tables).'
        : '';
    return `${instruction}

The selection covers a table region in a Word document (${rowCount} rows x ${colCount} columns total). Current contents of the covered cells (1-based absolute coordinates):

${formatCellGrid(cells)}

CRITICAL OUTPUT RULES (table mode):
- Respond with ONLY a JSON object — no commentary, no markdown code fences:
  {"cells":[{"row":2,"col":1,"text":"new cell text"}],"rowOps":[{"op":"delete","row":3},{"op":"insertAfter","row":4,"values":["cell A","cell B"]}]}
- "cells": list ONLY cells whose text you change. "row"/"col" must be coordinates from the listing above. "text" is the FULL new text of that cell (plain text, no markdown).
- "rowOps": list ONLY structural changes. Supported ops: "delete" (remove the row), "insertAfter" / "insertBefore" (insert one new row next to the given row, "values" = the new row's cell texts, exactly ${colCount} entries).
- If only cell text changes, omit "rowOps" or use an empty array. If nothing should change, respond with {"cells":[],"rowOps":[]}.
- Do NOT rewrite the table as text or markdown. Do NOT include unchanged cells.${mergedRules}`;
}

/**
 * Parses and validates the LLM's table-patch JSON.
 *
 * Lenient about transport noise (markdown fences, trailing commas) but strict
 * about semantics: out-of-scope coordinates, malformed operations, conflicts,
 * and content beyond the protocol limits are dropped with explicit warnings.
 * Cell entries whose text matches the current cell (ignoring surrounding
 * whitespace) are dropped as no-ops.
 *
 * @param {string} raw - Raw LLM response
 * @param {object} [shape]
 * @param {number} [shape.rowCount] - Total table rows (bounds for row coords);
 *   omit to reject the patch as undimensioned
 * @param {number} [shape.colCount] - Total table columns (bounds for col coords
 *   and insert-values length); omit to reject the patch as undimensioned
 * @param {string[][]} [shape.originals] - Current cell texts (0-based [r][c])
 *   for no-op detection; omit to keep every parsed cell entry
 * @param {Set<string>} [shape.shadowCoords] - 1-based "row,col" keys of grid
 *   slots covered by a merged cell (read-only); patch entries targeting them
 *   are dropped with a warning
 * @param {{startRow: number, endRow: number, startCol: number, endCol: number,
 *   allowRowOps?: boolean}} [shape.allowedBounds] - Optional selection scope.
 *   Structural changes additionally require allowRowOps=true and full width.
 * @returns {{cells: Array<{row: number, col: number, text: string}>,
 *   rowOps: Array<{op: string, row: number, values?: string[]}>,
 *   warnings: string[]}}
 *   cells are sorted top-left → bottom-right (document order);
 *   rowOps are pre-sorted into application order via planRowOpOrder.
 * @throws {Error} When no parseable JSON object is present
 */
export function parseTablePatchResponse(raw, shape = {}) {
    const warnings = [];
    const rawText = String(raw || '');
    if (rawText.length > TABLE_PATCH_LIMITS.MAX_RESPONSE_CHARS) {
        warnings.push(`Table patch response exceeds the ${TABLE_PATCH_LIMITS.MAX_RESPONSE_CHARS}-character limit (got ${rawText.length}) — entire patch ignored`);
        return { cells: [], rowOps: [], warnings };
    }

    const parsed = _extractJson(rawText);
    if (Object.prototype.hasOwnProperty.call(parsed, 'colOps')) {
        warnings.push('Unsupported "colOps" field — ignored');
    }

    const rowCount = _asInt(shape && shape.rowCount);
    const colCount = _asInt(shape && shape.colCount);
    if (!rowCount || rowCount < 1 || !colCount || colCount < 1) {
        warnings.push('Invalid table dimensions — entire patch ignored');
        return { cells: [], rowOps: [], warnings };
    }

    const boundsResult = _normalizeAllowedBounds(shape && shape.allowedBounds, rowCount, colCount, warnings);
    if (!boundsResult.valid) {
        return { cells: [], rowOps: [], warnings };
    }

    const originals = shape && shape.originals;
    const cells = _parseCells(parsed.cells, {
        rowCount, colCount, originals, allowedBounds: boundsResult.bounds,
        shadowCoords: shape && shape.shadowCoords, warnings,
    });
    const rowOps = planRowOpOrder(_parseRowOps(parsed.rowOps, {
        rowCount, colCount, allowedBounds: boundsResult.bounds, warnings,
    }));
    const limited = _applyTotalTextLimit(cells, rowOps, warnings);

    return { cells: limited.cells, rowOps: limited.rowOps, warnings };
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
function _normalizeAllowedBounds(rawBounds, rowCount, colCount, warnings) {
    if (rawBounds === undefined) return { valid: true, bounds: null };
    if (!rawBounds || typeof rawBounds !== 'object' || Array.isArray(rawBounds)) {
        warnings.push('Invalid "allowedBounds" — entire patch ignored');
        return { valid: false, bounds: null };
    }

    const startRow = _asInt(rawBounds.startRow);
    const endRow = _asInt(rawBounds.endRow);
    const startCol = _asInt(rawBounds.startCol);
    const endCol = _asInt(rawBounds.endCol);
    if (!startRow || !endRow || !startCol || !endCol
        || startRow < 1 || endRow > rowCount || startRow > endRow
        || startCol < 1 || endCol > colCount || startCol > endCol) {
        warnings.push('Invalid "allowedBounds" coordinates — entire patch ignored');
        return { valid: false, bounds: null };
    }

    return {
        valid: true,
        bounds: {
            startRow, endRow, startCol, endCol,
            allowRowOps: rawBounds.allowRowOps === true,
        },
    };
}

/** @private */
function _parseCells(rawCells, { rowCount, colCount, originals, allowedBounds, shadowCoords, warnings }) {
    if (rawCells === undefined || rawCells === null) return [];
    if (!Array.isArray(rawCells)) {
        warnings.push('"cells" is not an array — ignored');
        return [];
    }
    if (rawCells.length > TABLE_PATCH_LIMITS.MAX_CELL_ENTRIES) {
        warnings.push(`"cells" contains ${rawCells.length} entries; limit is ${TABLE_PATCH_LIMITS.MAX_CELL_ENTRIES} — all cell entries ignored`);
        return [];
    }

    const byKey = new Map();
    const seenTextByKey = new Map();
    const conflictedKeys = new Set();
    const duplicateWarnings = new Set();
    for (const entry of rawCells) {
        const row = _asInt(entry && entry.row);
        const col = _asInt(entry && entry.col);
        if (!row || !col || row < 1 || row > rowCount || col < 1 || col > colCount) {
            warnings.push(`Cell coordinate out of bounds (row=${entry && entry.row}, col=${entry && entry.col}) — dropped`);
            continue;
        }
        if (allowedBounds && (row < allowedBounds.startRow || row > allowedBounds.endRow
            || col < allowedBounds.startCol || col > allowedBounds.endCol)) {
            warnings.push(`Cell R${row}C${col} is outside allowedBounds — dropped`);
            continue;
        }
        if (shadowCoords && shadowCoords.has(`${row},${col}`)) {
            warnings.push(`Cell R${row}C${col} is covered by a merged cell — not editable, dropped`);
            continue;
        }

        const text = _asText(entry && entry.text);
        if (text === null) {
            warnings.push(`Cell R${row}C${col} has no usable "text" — dropped`);
            continue;
        }

        const key = `${row},${col}`;
        if (conflictedKeys.has(key)) continue;
        if (seenTextByKey.has(key)) {
            if (seenTextByKey.get(key) !== text) {
                byKey.delete(key);
                conflictedKeys.add(key);
                warnings.push(`Conflicting duplicate cell entries for R${row}C${col} — all dropped`);
            } else if (!duplicateWarnings.has(key)) {
                duplicateWarnings.add(key);
                warnings.push(`Duplicate identical cell entry for R${row}C${col} — coalesced`);
            }
            continue;
        }
        seenTextByKey.set(key, text);

        if (text.length > TABLE_PATCH_LIMITS.MAX_CELL_TEXT_CHARS) {
            warnings.push(`Cell R${row}C${col} text exceeds the ${TABLE_PATCH_LIMITS.MAX_CELL_TEXT_CHARS}-character per-cell limit — dropped`);
            continue;
        }
        byKey.set(key, { row, col, text });
    }

    let cells = [...byKey.values()].sort((a, b) => (a.row - b.row) || (a.col - b.col));
    if (Array.isArray(originals)) {
        cells = cells.filter((cell) => {
            const sourceRow = originals[cell.row - 1];
            const original = _asText(Array.isArray(sourceRow) ? sourceRow[cell.col - 1] : '') ?? '';
            return original.trim() !== cell.text.trim();
        });
    }
    return cells;
}

/** @private */
function _parseRowOps(rawOps, { rowCount, colCount, allowedBounds, warnings }) {
    if (rawOps === undefined || rawOps === null) return [];
    if (!Array.isArray(rawOps)) {
        warnings.push('"rowOps" is not an array — ignored');
        return [];
    }
    if (rawOps.length > TABLE_PATCH_LIMITS.MAX_ROW_OPS) {
        warnings.push(`"rowOps" contains ${rawOps.length} entries; limit is ${TABLE_PATCH_LIMITS.MAX_ROW_OPS} — all row operations ignored`);
        return [];
    }
    if (rawOps.length === 0) return [];
    if (allowedBounds && !allowedBounds.allowRowOps) {
        warnings.push('Row operations are not allowed by allowedBounds — all rowOps ignored');
        return [];
    }
    if (allowedBounds && (allowedBounds.startCol !== 1 || allowedBounds.endCol !== colCount)) {
        warnings.push('Row operations require a full-width allowedBounds selection — all rowOps ignored');
        return [];
    }

    const bySlot = new Map();
    const conflictedSlots = new Set();
    const duplicateWarnings = new Set();
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
        if (allowedBounds && (row < allowedBounds.startRow || row > allowedBounds.endRow)) {
            warnings.push(`Row op "${op}" at row ${row} is outside allowedBounds — dropped`);
            continue;
        }

        let candidate;
        if (op === ROW_OP.DELETE) {
            candidate = { op, row };
        } else {
            const values = _asRowValues(entry && entry.values, colCount, op, row, warnings);
            if (values === null) continue;
            candidate = { op, row, values };
        }

        const slot = `${op}:${row}`;
        if (conflictedSlots.has(slot)) continue;
        const existing = bySlot.get(slot);
        if (!existing) {
            bySlot.set(slot, candidate);
        } else if (_sameRowOp(existing, candidate)) {
            if (!duplicateWarnings.has(slot)) {
                duplicateWarnings.add(slot);
                warnings.push(`Duplicate identical row op "${op}" at row ${row} — coalesced`);
            }
        } else {
            bySlot.delete(slot);
            conflictedSlots.add(slot);
            warnings.push(`Conflicting duplicate row op "${op}" at row ${row} — all variants dropped`);
        }
    }

    const ops = [...bySlot.values()];
    const deletedRows = new Set(ops.filter((op) => op.op === ROW_OP.DELETE).map((op) => op.row));
    if (deletedRows.size === rowCount) {
        warnings.push('Row operations attempt to delete all existing rows — all rowOps ignored');
        return [];
    }
    return ops;
}

/** @private */
function _asRowValues(rawValues, colCount, op, row, warnings) {
    if (!Array.isArray(rawValues)) {
        warnings.push(`Row op "${op}" at row ${row} has no "values" array — dropped`);
        return null;
    }
    if (rawValues.length !== colCount) {
        warnings.push(`Row op "${op}" at row ${row} has ${rawValues.length} values; exactly ${colCount} required — dropped`);
        return null;
    }

    const values = [];
    for (let index = 0; index < rawValues.length; index++) {
        const text = _asText(rawValues[index]);
        if (text === null) {
            warnings.push(`Row op "${op}" at row ${row} value ${index + 1} has no usable text — dropped`);
            return null;
        }
        if (text.length > TABLE_PATCH_LIMITS.MAX_CELL_TEXT_CHARS) {
            warnings.push(`Row op "${op}" at row ${row} value ${index + 1} exceeds the ${TABLE_PATCH_LIMITS.MAX_CELL_TEXT_CHARS}-character per-cell limit — dropped`);
            return null;
        }
        values.push(text);
    }
    return values;
}

/** @private */
function _sameRowOp(a, b) {
    if (a.op !== b.op || a.row !== b.row) return false;
    if (a.op === ROW_OP.DELETE) return true;
    return a.values.length === b.values.length && a.values.every((value, index) => value === b.values[index]);
}

/** @private */
function _applyTotalTextLimit(cells, rowOps, warnings) {
    let total = 0;
    const keptCells = [];
    for (const cell of cells) {
        if (total + cell.text.length > TABLE_PATCH_LIMITS.MAX_TOTAL_TEXT_CHARS) {
            warnings.push(`Cell R${cell.row}C${cell.col} would exceed the ${TABLE_PATCH_LIMITS.MAX_TOTAL_TEXT_CHARS}-character total text limit — dropped`);
            continue;
        }
        total += cell.text.length;
        keptCells.push(cell);
    }

    const keptRowOps = [];
    for (const op of rowOps) {
        const textLength = op.op === ROW_OP.DELETE
            ? 0
            : op.values.reduce((sum, value) => sum + value.length, 0);
        if (total + textLength > TABLE_PATCH_LIMITS.MAX_TOTAL_TEXT_CHARS) {
            warnings.push(`Row op "${op.op}" at row ${op.row} would exceed the ${TABLE_PATCH_LIMITS.MAX_TOTAL_TEXT_CHARS}-character total text limit — dropped`);
            continue;
        }
        total += textLength;
        keptRowOps.push(op);
    }
    return { cells: keptCells, rowOps: keptRowOps };
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
