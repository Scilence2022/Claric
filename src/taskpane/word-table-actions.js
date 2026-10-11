/**
 * Word Table Actions Module
 *
 * The table write paths split out of word-actions.js so their code (and the
 * table patch/creation protocols in lib/table-patch.js and lib/table-ops.js)
 * downloads only when a turn actually edits or creates a table:
 *
 * - prepareTableAmendment / applyTablePatch: the multi-cell selection patch
 *   protocol, plus the region/table style ops the tool loop rides on it.
 * - applyMixedTableAmendment: paragraph-granular apply for selections that
 *   overlap a table without being contained in it.
 * - prepareTableProposal / applyTableProposal: native table creation.
 *
 * word-actions.js keeps same-named wrappers that load this module through
 * task-module-loader.js, so callers and `deps.actions` overrides are
 * unchanged. Selection/table READS stay in word-actions.js: every selection
 * edit runs them to decide whether this module is needed at all.
 *
 * Prepare halves receive the private LLM request helper as
 * `deps.sendActionRequest` (see prepareFormatProposal for the same contract).
 *
 * @module word-table-actions
 */

import { applyTokenMapStrategy } from '../lib/word-diff/index.js';
import { hasCjk, applyCharDiffStrategy } from '../lib/word-diff/char-diff.js';
import { buildTableUserPrompt, parseTablePatchResponse } from '../lib/table-patch.js';
import {
    inferTableCreationSpec, buildTableCreationPrompt, parseTableCreationResponse,
    validateTableCreationSpec,
} from '../lib/table-ops.js';
import { supportsTrackedRowOps } from '../lib/platform.js';
import { extractDocumentStructured } from '../lib/comment-extractor.js';
import { queueRevisionRead, resolveRevisionRead, RevisionSafetyError } from '../lib/word-revisions.js';
import { _alignParagraphs, _normalizeLineEndings } from '../lib/reassembler.js';
import { getActiveBackendConfig } from './app-state.js';
import { enumValue as _enumValue, applyFontOps as _applyFontOps } from './word-format-helpers.js';
import { readSelectionTableContext, readSelectionSnippet, _normalizeSelectionText } from './word-actions.js';

/**
 * Prepare half for multi-cell table selections: sends the covered cells as a
 * coordinate grid and expects a JSON patch (lib/table-patch.js) back. The
 * proposal carries per-cell/per-row items for the review card and the parsed
 * patch for applySelectionAmendment.
 *
 * @param {object} deps - { appState, log }
 * @param {object} args
 * @param {{rowCount: number, colCount: number, cells: Array}} args.tableRegion
 * @param {string} args.promptTemplate - The amendment instruction/template
 * @param {string} [args.commentInstructions] - Ignored for table selections
 * @param {object} args.backendConfig - Active provider config
 * @param {AbortSignal} [args.signal] - Cancellation signal (stop button)
 * @returns {Promise<object>} Proposal with tablePatch + tableItems
 */
export async function prepareTableAmendment(deps, { tableRegion, promptTemplate, commentInstructions, backendConfig, onToken, onReasoning, signal }) {
    const { appState, log } = deps;
    const { rowCount, colCount, cells } = tableRegion;
    log(`Processing table selection (${cells.length} cells) via ${backendConfig.model}...`, 'info');
    if (commentInstructions && commentInstructions.trim()) {
        log('Comment instructions are ignored for table selections — the patch protocol returns JSON.', 'info');
    }

    // The grid plays the {selection} role: templates that name it read
    // naturally when it points at the cell listing below.
    const instruction = (promptTemplate || '').includes('{selection}')
        ? promptTemplate.replace(/{selection}/g, 'the selected table cells')
        : (promptTemplate || '');
    const baseInstruction = tableRegion.mergedUnknown
        ? `${instruction}\n\nNOTE: This table contains merged cells whose layout could not be mapped. Some listed coordinates may be covered by a merge — edits to them are skipped at apply time.`
        : instruction;

    const messages = [];
    const contextPrompt = appState.promptManager.getActivePrompt('context');
    if (contextPrompt) {
        messages.push({ role: 'system', content: contextPrompt.template });
    }
    messages.push({ role: 'user', content: buildTableUserPrompt(baseInstruction, cells, { rowCount, colCount }) });

    const rawResponse = await deps.sendActionRequest(deps, backendConfig, messages, { onToken, onReasoning, signal });
    log(`LLM response received [${backendConfig.model}]`, 'success');

    // Full-table grid for no-op detection, per-cell "before" text, and
    // apply-time staleness checks. The extraction already loaded the whole
    // matrix; using it directly keeps coordinates honest for cells outside
    // the covered region too.
    const originals = Array.isArray(tableRegion.values) && tableRegion.values.length
        ? tableRegion.values
        : Array.from({ length: rowCount }, () => Array(colCount).fill(''));
    if (!tableRegion.values) {
        for (const c of cells) originals[c.row - 1][c.col - 1] = c.text;
    }

    // Restrict the model's patch to the selected rectangle. Structural row
    // ops are allowed only for full-width selections: inserting/deleting a
    // row rewrites cells the user never selected otherwise. Merged tables
    // never allow row ops (values arrays assume a uniform grid; deleting a
    // row that intersects a vertical merge is structurally ambiguous).
    const merged = !!tableRegion.merged;
    const bounds = tableRegion.bounds || {
        startRow: 1, endRow: rowCount, startCol: 1, endCol: colCount,
    };
    const allowedBounds = {
        ...bounds,
        allowRowOps: !merged && bounds.startCol === 1 && bounds.endCol === colCount,
    };

    const patch = parseTablePatchResponse(rawResponse, {
        rowCount, colCount, originals, allowedBounds,
        shadowCoords: tableRegion.shadowKeys,
    });
    for (const warning of patch.warnings) log(`Table patch: ${warning}`, 'warning');

    const tableItems = [
        ...patch.cells.map((c) => ({
            label: `Cell R${c.row}C${c.col}`,
            before: originals[c.row - 1][c.col - 1],
            after: c.text,
            searchText: originals[c.row - 1][c.col - 1].trim().slice(0, 60) || undefined,
        })),
        ...patch.rowOps.map((op) => (op.op === 'delete'
            ? {
                label: `Delete row ${op.row}`,
                before: (originals[op.row - 1] || []).join(' | '),
                after: '',
                searchText: ((originals[op.row - 1] || [])[0] || '').trim().slice(0, 60) || undefined,
            }
            : {
                label: `${op.op === 'insertAfter' ? 'Insert row after' : 'Insert row before'} row ${op.row}`,
                before: '',
                after: op.values.join(' | '),
                searchText: ((originals[op.row - 1] || [])[0] || '').trim().slice(0, 60) || undefined,
            })),
    ];

    return {
        selectionText: cells.map((c) => c.text).join('\n'),
        amendedText: null,
        commentText: null,
        model: backendConfig.model,
        tablePatch: {
            rowCount, colCount, cells: patch.cells, rowOps: patch.rowOps,
            bounds, originals,
        },
        tableItems,
    };
}

/**
 * Applies a parsed table patch: per-cell text revisions first (original
 * coordinates — cell edits never shift coordinates), then row structure ops
 * in pre-sorted descending order.
 *
 * Row insertions/deletions are only recorded as tracked revisions by Word
 * desktop (see lib/platform.js). On hosts without that support the structure
 * phase runs with tracking OFF and a warning, so no half-tracked state leaks;
 * cell text edits are still tracked wherever possible.
 *
 * @param {object} deps - { appState, log }
 * @param {object} proposal - Result of prepareSelectionAmendment (table route)
 * @returns {Promise<{cellsApplied: number, cellsSkipped: number,
 *   rowOpsApplied: number, warnings: string[]}>} Actual application counts —
 *   the caller settles the card honestly instead of assuming success.
 */
export async function applyTablePatch(deps, proposal) {
    const { appState, log } = deps;
    const { tablePatch } = proposal;
    const trackChanges = !!appState.config.trackChangesEnabled;
    const rowTracking = trackChanges && supportsTrackedRowOps(appState.platform);
    const warnings = [];

    if (tablePatch.rowOps.length > 0 && trackChanges && !rowTracking) {
        const warning = `Row insertions/deletions cannot be tracked as revisions on this host (${appState.platform}) — applied directly.`;
        warnings.push(warning);
        log(`${warning} Cell text edits are still tracked.`, 'warning');
    }

    const merges = tablePatch.merges || [];
    if (merges.length > 0) {
        const warning = 'Cell merges are structural and cannot be tracked as revisions — applied directly.';
        warnings.push(warning);
        log(warning, 'warning');
    }

    // Style ops split by binding: region-bound ops run while ORIGINAL
    // coordinates are still valid (before row structure changes); table-level
    // ops run after the structure settles. Font/shading writes record as
    // format revisions wherever the host supports them.
    const styleOps = tablePatch.styleOps || [];
    const regionStyleOps = styleOps.filter((op) => op.region);
    const tableStyleOps = styleOps.filter((op) => !op.region);

    log('Applying table patch...', 'info');
    let cellsApplied = 0;
    let cellsSkipped = 0;
    let rowOpsApplied = 0;
    let mergesApplied = 0;
    let styleOpsApplied = 0;

    await Word.run(async (context) => {
        // Multi-table patches (document-scope table_management) anchor each
        // op by its tableIndex against body.tables; single-table selection
        // patches keep the legacy parentTable anchoring. A document-scope
        // session on a ONE-table document also anchors by body.tables (the
        // selection may sit in a different table than the one it manages).
        const multiTablePatch = (tablePatch.tableCount || 1) > 1 || proposal.tableSource === 'document';
        /** @type {Map<number, object>} */
        let tableByIndex = new Map();
        let tablesLoaded = false;
        const loadTablesIfNeeded = async () => {
            if (!multiTablePatch || tablesLoaded) return;
            const tablesProxy = context.document.body.tables;
            tablesProxy.load('items');
            await context.sync();
            const items = (tablesProxy.items || []).filter((t) => !t.isNullObject);
            tableByIndex = new Map();
            for (let i = 0; i < items.length; i++) {
                const t = items[i];
                t.load('isNullObject,rowCount,values,isUniform');
                await context.sync();
                tableByIndex.set(i + 1, t);
            }
            tablesLoaded = true;
        };
        /** Resolves the table proxy for one op (multi-table) or the single
         *  anchored table (selection path). @private */
        const tableFor = async (op) => {
            if (!multiTablePatch) return table;
            await loadTablesIfNeeded();
            const t = tableByIndex.get((op && op.tableIndex) || 1);
            if (!t) {
                warnings.push(`Table ${(op && op.tableIndex) || 1} no longer exists — op skipped.`);
                return null;
            }
            return t;
        };

        const selection = context.document.getSelection();
        const table = selection.parentTableOrNullObject;
        if (multiTablePatch) {
            // Document-scope patch: no selection anchor; validate the table
            // map is loadable and each referenced table exists.
            await loadTablesIfNeeded();
            const referenced = new Set();
            for (const op of tablePatch.cells) referenced.add((op.tableIndex || 1));
            for (const op of tablePatch.rowOps) referenced.add((op.tableIndex || 1));
            for (const m of tablePatch.merges || []) referenced.add((m.tableIndex || 1));
            for (const op of tablePatch.styleOps || []) referenced.add((op.tableIndex || 1));
            for (const idx of referenced) {
                if (!tableByIndex.has(idx)) {
                    throw new Error(`Table ${idx} no longer exists in the document (${tableByIndex.size} table(s) found). Draft a new edit instead.`);
                }
            }
        } else {
            table.load('isNullObject,rowCount,values,isUniform');
            await context.sync();
            if (table.isNullObject) {
                throw new Error('The selection is no longer inside the table — re-select the region and apply again.');
            }
            if (table.rowCount !== tablePatch.rowCount) {
                throw new Error(`Table changed since this proposal was drafted (${tablePatch.rowCount} → ${table.rowCount} rows). Draft a new edit instead.`);
            }
        }

        // Staleness guard: the patch's coordinates were bound to the table(s)
        // as they were at prepare time. If a touched cell (or a row referenced
        // by a structure op) no longer holds its original text, the proposal
        // is stale — applying it would silently overwrite the user's newer edit.
        const stale = [];
        for (const cellPatch of tablePatch.cells) {
            const t = await tableFor(cellPatch);
            if (!t) continue;
            const currentValues = (t && t.values) || [];
            const before = tablePatch.tableOriginals
                ? tablePatch.tableOriginals[cellPatch.tableIndex || 1]
                : tablePatch.originals || [];
            const beforeText = (before[cellPatch.row - 1] || [])[cellPatch.col - 1] || '';
            const now = (currentValues[cellPatch.row - 1] || [])[cellPatch.col - 1];
            if (now === undefined || now.trim() !== String(beforeText).trim()) {
                stale.push(`R${cellPatch.row}C${cellPatch.col}`);
            }
        }
        for (const op of tablePatch.rowOps) {
            const t = await tableFor(op);
            if (!t) continue;
            const beforeRow = tablePatch.tableOriginals
                ? (tablePatch.tableOriginals[op.tableIndex || 1] || [])[op.row - 1] || []
                : (tablePatch.originals || [])[op.row - 1] || [];
            const beforeRowText = Array.isArray(beforeRow) ? beforeRow.join('').trim() : String(beforeRow).trim();
            const nowRow = ((t.values || [])[op.row - 1] || []).join('').trim();
            if (beforeRowText !== nowRow) stale.push(`row ${op.row}`);
        }
        if (stale.length > 0) {
            throw new Error(
                `The table changed since this proposal was drafted (${stale.slice(0, 3).join(', ')}` +
                `${stale.length > 3 ? ', …' : ''}). Draft a new edit instead.`
            );
        }

        try {
            // Phase 1: cell text patches (coordinates still original).
            if (tablePatch.cells.length > 0) {
                if (Word.ChangeTrackingMode) {
                    context.document.changeTrackingMode = trackChanges
                        ? Word.ChangeTrackingMode.trackAll
                        : Word.ChangeTrackingMode.off;
                }
                for (const cellPatch of tablePatch.cells) {
                    const t = await tableFor(cellPatch);
                    if (!t) continue;
                    const applied = await _patchCell(context, t, cellPatch, log);
                    if (applied) cellsApplied++; else cellsSkipped++;
                }
            }

            // Phase 2: region-bound style ops (shading/alignment/font).
            // Still ahead of the row structure ops, so original coordinates
            // remain valid. Each op syncs separately — one unsupported op
            // degrades to a warning instead of failing the batch.
            if (regionStyleOps.length > 0) {
                if (Word.ChangeTrackingMode) {
                    context.document.changeTrackingMode = trackChanges
                        ? Word.ChangeTrackingMode.trackAll
                        : Word.ChangeTrackingMode.off;
                }
                for (const op of regionStyleOps) {
                    const t = await tableFor(op);
                    if (!t) continue;
                    const applied = await _applyRegionStyleOp(context, t, op, log, warnings);
                    if (applied) styleOpsApplied++;
                }
            }

            // Phase 2: row structure ops (descending coordinates).
            if (tablePatch.rowOps.length > 0) {
                if (Word.ChangeTrackingMode) {
                    context.document.changeTrackingMode = rowTracking
                        ? Word.ChangeTrackingMode.trackAll
                        : Word.ChangeTrackingMode.off;
                    await context.sync();
                }
                for (const op of tablePatch.rowOps) {
                    const t = await tableFor(op);
                    if (!t) continue;
                    const row = t.getCell(op.row - 1, 0).parentRow;
                    if (op.op === 'delete') {
                        row.delete();
                    } else if (typeof row.insertRows === 'function') {
                        row.insertRows(
                            op.op === 'insertAfter' ? Word.InsertLocation.after : Word.InsertLocation.before,
                            1,
                            [op.values]
                        );
                    } else {
                        log(`Row insert is not supported by this Word API (skipped row ${op.row})`, 'warning');
                        continue;
                    }
                    rowOpsApplied++;
                }
                await context.sync();
            }

            // Phase 3: cell merges. Merge is a structure op — hosts do not
            // track it as a revision, so it always runs untracked here. The
            // non-anchor cells are cleared first so the merged cell ends up
            // holding the anchor's content deterministically.
            if (merges.length > 0) {
                if (Word.ChangeTrackingMode) {
                    context.document.changeTrackingMode = Word.ChangeTrackingMode.off;
                    await context.sync();
                }
                for (const m of merges) {
                    const t = await tableFor(m);
                    if (!t) continue;
                    const anchor = t.getCell(m.startRow - 1, m.startCol - 1);
                    // Word.js: Table.mergeCells(range) requires a selection-like
                    // range — a body.getRange().expandTo() span yields
                    // InvalidArgument. The cell-level TableCell.merge(other)
                    // takes TableCells directly (no range), which is what Word
                    // accepts. Re-fetch the anchor each iteration since the
                    // grid changes as cells merge in.
                    if (typeof anchor.merge !== 'function') {
                        warnings.push('Cell merge is not supported on this Word host — skipped.');
                        log('Cell merge skipped (Word host lacks TableCell.merge).', 'warning');
                        continue;
                    }
                    // Clear the non-anchor cells so Word's concatenation on
                    // merge yields just the anchor's content.
                    for (let r = m.startRow; r <= m.endRow; r++) {
                        for (let c = m.startCol; c <= m.endCol; c++) {
                            if (r === m.startRow && c === m.startCol) continue;
                            const cell = t.getCell(r - 1, c - 1);
                            const paras = cell.body.paragraphs;
                            paras.load('items');
                            await context.sync();
                            try {
                                for (const para of paras.items) para.clear();
                            } catch (clearErr) {
                                log(`Merge: could not clear R${r}C${c} (${clearErr.message || clearErr})`, 'warning');
                            }
                        }
                    }
                    await context.sync();

                    for (let r = m.startRow; r <= m.endRow; r++) {
                        for (let c = m.startCol; c <= m.endCol; c++) {
                            if (r === m.startRow && c === m.startCol) continue;
                            t.getCell(m.startRow - 1, m.startCol - 1).merge(t.getCell(r - 1, c - 1));
                        }
                    }
                    await context.sync();
                    mergesApplied++;
                }
            }

            // Phase 4: table-level style ops — table style/banding, borders,
            // header rows, layout, column widths, whole-table formats. Run
            // after the structure settled; each op syncs separately so one
            // unsupported op degrades to a warning.
            if (tableStyleOps.length > 0) {
                if (Word.ChangeTrackingMode) {
                    context.document.changeTrackingMode = trackChanges
                        ? Word.ChangeTrackingMode.trackAll
                        : Word.ChangeTrackingMode.off;
                    await context.sync();
                }
                for (const op of tableStyleOps) {
                    const t = await tableFor(op);
                    if (!t) continue;
                    const applied = await _applyTableStyleOp(context, t, op, log, warnings);
                    if (applied) styleOpsApplied++;
                }
            }
        } finally {
            // Later turns and comments must not inherit tracking state.
            if (Word.ChangeTrackingMode) {
                context.document.changeTrackingMode = Word.ChangeTrackingMode.off;
                await context.sync();
            }
        }
    });

    log(`Table patch applied: ${cellsApplied} cell(s) revised, ${rowOpsApplied} row op(s)` +
        (mergesApplied ? `, ${mergesApplied} merge(s)` : '') +
        (styleOpsApplied ? `, ${styleOpsApplied} style op(s)` : '') +
        (cellsSkipped ? `, ${cellsSkipped} cell(s) already up to date` : ''), 'success');
    return { cellsApplied, cellsSkipped, rowOpsApplied, mergesApplied, styleOpsApplied, warnings };
}

/**
 * Applies one region-bound style op (cellFormat / font with a target).
 * Full-width regions use the TableRow properties (one write per row);
 * narrower regions fall back to per-cell writes — a row write would spill
 * outside the covered region on partial-width selections. Returns false and
 * pushes a warning when the op fails.
 *
 * @param {Word.RequestContext} context
 * @param {Word.Table} table
 * @param {object} op - {type: 'cellFormat'|'font', region, ...payload}
 * @param {function} log
 * @param {string[]} warnings
 * @returns {Promise<boolean>}
 * @private
 */
async function _applyRegionStyleOp(context, table, op, log, warnings) {
    const { region } = op;
    try {
        const fullRowBands = region.startCol === 1 && region.endCol === _tableColumnCount(table);
        if (op.type === 'font') {
            if (fullRowBands) {
                for (let r = region.startRow; r <= region.endRow; r++) {
                    const row = table.getCell(r - 1, 0).parentRow;
                    if (row && row.font) _applyFontOps(row.font, op.font, log);
                }
                await context.sync();
                return true;
            }
            // Partial-width font: write through each cell's paragraph ranges
            // (TableCell has no font property).
            const cellFonts = [];
            for (let r = region.startRow; r <= region.endRow; r++) {
                for (let c = region.startCol; c <= region.endCol; c++) {
                    const cell = table.getCell(r - 1, c - 1);
                    const paragraphs = cell.body.paragraphs;
                    paragraphs.load('items');
                    cellFonts.push({ cell, paragraphs });
                }
            }
            await context.sync();
            for (const { paragraphs } of cellFonts) {
                for (const para of paragraphs.items || []) {
                    _applyFontOps(para.getRange(Word.RangeLocation.content).font, op.font, log);
                }
            }
            await context.sync();
            return true;
        }
        // cellFormat
        if (fullRowBands) {
            for (let r = region.startRow; r <= region.endRow; r++) {
                const row = table.getCell(r - 1, 0).parentRow;
                _applyCellFormatProps(row, op, log);
            }
            await context.sync();
            return true;
        }
        for (let r = region.startRow; r <= region.endRow; r++) {
            for (let c = region.startCol; c <= region.endCol; c++) {
                _applyCellFormatProps(table.getCell(r - 1, c - 1), op, log);
            }
        }
        await context.sync();
        return true;
    } catch (err) {
        const warning = `Style op (${op.type}) failed: ${err.message || err} — skipped.`;
        warnings.push(warning);
        log(warning, 'warning');
        return false;
    }
}

/** Column count of a table whose values were already loaded. @private */
function _tableColumnCount(table) {
    const values = table.values;
    return Array.isArray(values) && values[0] ? values[0].length : 0;
}

/**
 * Writes a cellFormat payload onto a Table / TableRow / TableCell proxy —
 * all three expose shadingColor and the two alignment properties.
 * @private
 */
function _applyCellFormatProps(target, op, log) {
    if (op.shadingColor) target.shadingColor = op.shadingColor;
    if (op.horizontalAlignment) {
        const value = _enumValue(Word.Alignment, op.horizontalAlignment);
        if (value === undefined) log(`Table style: unknown alignment "${op.horizontalAlignment}"`, 'warning');
        else target.horizontalAlignment = value;
    }
    if (op.verticalAlignment) {
        const value = _enumValue(Word.VerticalAlignment, op.verticalAlignment);
        if (value === undefined) log(`Table style: unknown vertical alignment "${op.verticalAlignment}"`, 'warning');
        else target.verticalAlignment = value;
    }
}

/**
 * Applies one table-level style op (tableStyle / borders / headerRow /
 * layout / columnWidths, and target-less cellFormat/font). Returns false
 * and pushes a warning when the op fails.
 *
 * @param {Word.RequestContext} context
 * @param {Word.Table} table
 * @param {object} op
 * @param {function} log
 * @param {string[]} warnings
 * @returns {Promise<boolean>}
 * @private
 */
async function _applyTableStyleOp(context, table, op, log, warnings) {
    try {
        switch (op.type) {
            case 'tableStyle': {
                if (op.style) {
                    const builtIn = _enumValue(Word.BuiltInStyleName, op.style);
                    if (builtIn !== undefined) table.styleBuiltIn = builtIn;
                    else table.style = op.style; // custom/localized style name
                }
                for (const [key, prop] of [
                    ['bandedRows', 'styleBandedRows'], ['bandedColumns', 'styleBandedColumns'],
                    ['firstColumn', 'styleFirstColumn'], ['lastColumn', 'styleLastColumn'],
                    ['totalRow', 'styleTotalRow'],
                ]) {
                    if (op[key] !== undefined) table[prop] = op[key];
                }
                await context.sync();
                return true;
            }
            case 'borders': {
                if (op.row !== undefined) {
                    const row = table.getCell(op.row - 1, 0).parentRow;
                    for (const [location, spec] of Object.entries(op.borders)) {
                        _setBorderSpec(row.getBorder(_borderLocation(location)), spec, log);
                    }
                } else {
                    for (const [location, spec] of Object.entries(op.borders)) {
                        _setBorderSpec(table.getBorder(_borderLocation(location)), spec, log);
                    }
                }
                await context.sync();
                return true;
            }
            case 'headerRow': {
                table.headerRowCount = op.rows;
                await context.sync();
                if (op.font || op.shadingColor) {
                    for (let r = 1; r <= op.rows; r++) {
                        const row = table.getCell(r - 1, 0).parentRow;
                        if (op.shadingColor) row.shadingColor = op.shadingColor;
                        if (op.font && row.font) _applyFontOps(row.font, op.font, log);
                    }
                    await context.sync();
                }
                return true;
            }
            case 'layout': {
                if (op.alignment) {
                    const value = _enumValue(Word.Alignment, op.alignment);
                    if (value === undefined) log(`Table style: unknown alignment "${op.alignment}"`, 'warning');
                    else table.alignment = value;
                }
                if (op.widthPt) table.width = op.widthPt;
                if (op.autoFitWindow && typeof table.autoFitWindow === 'function') table.autoFitWindow();
                if (op.distributeColumns && typeof table.distributeColumns === 'function') table.distributeColumns();
                if (op.cellPaddingPt !== undefined && typeof table.setCellPadding === 'function') {
                    for (const side of ['Top', 'Left', 'Bottom', 'Right']) {
                        table.setCellPadding(side, op.cellPaddingPt);
                    }
                }
                await context.sync();
                return true;
            }
            case 'columnWidths': {
                for (let c = 0; c < op.widthsPt.length; c++) {
                    table.getCell(0, c).columnWidth = op.widthsPt[c];
                }
                await context.sync();
                return true;
            }
            case 'cellFormat': {
                // Target-less cellFormat = whole table via Table properties.
                _applyCellFormatProps(table, op, log);
                await context.sync();
                return true;
            }
            case 'font': {
                if (table.font) {
                    _applyFontOps(table.font, op.font, log);
                    await context.sync();
                    return true;
                }
                // No table.font on this host — fall back to row fonts.
                const rowCount = table.rowCount || 0;
                for (let r = 0; r < rowCount; r++) {
                    const row = table.getCell(r, 0).parentRow;
                    if (row && row.font) _applyFontOps(row.font, op.font, log);
                }
                await context.sync();
                return true;
            }
            default: {
                const warning = `Unknown style op type "${op.type}" — skipped.`;
                warnings.push(warning);
                log(warning, 'warning');
                return false;
            }
        }
    } catch (err) {
        const warning = `Style op (${op.type}) failed: ${err.message || err} — skipped.`;
        warnings.push(warning);
        log(warning, 'warning');
        return false;
    }
}

/** Maps a border-set key onto the Word border location. @private */
function _borderLocation(key) {
    const map = {
        top: 'top', bottom: 'bottom', left: 'left', right: 'right',
        insideH: 'insideHorizontal', insideV: 'insideVertical',
    };
    const enumKey = map[key] || key;
    return (Word.BorderLocation && Word.BorderLocation[enumKey]) || enumKey;
}

/** Writes one validated border spec onto a TableBorder proxy. @private */
function _setBorderSpec(border, spec, log) {
    const type = _enumValue(Word.BorderType, spec.type);
    if (type === undefined) {
        log(`Table style: unknown border type "${spec.type}"`, 'warning');
        return;
    }
    border.type = type;
    if (spec.type !== 'none') {
        if (spec.color) border.color = spec.color;
        if (spec.width !== undefined) border.width = spec.width;
    }
}

/**
 * Revises one cell's text. Single-paragraph cells go through the granular
 * diff strategies (tracked in-cell edits are native to Word); the rare
 * multi-paragraph cell falls back to per-paragraph replacement when the line
 * count matches, else a coarse whole-content replace. Returns true when the
 * cell was written, false when its text already matched.
 *
 * Addressing a cell can itself fail on merged tables — some hosts throw
 * ItemNotFound for getCell on merge-covered coordinates — which degrades to
 * a skip with a warning instead of failing the whole patch.
 *
 * @private
 */
async function _patchCell(context, table, cellPatch, log) {
    const label = `R${cellPatch.row}C${cellPatch.col}`;
    let cell;
    try {
        cell = table.getCell(cellPatch.row - 1, cellPatch.col - 1);
        const paragraphs = cell.body.paragraphs;
        paragraphs.load('items');
        await context.sync();
        if (paragraphs.items.length === 0) return false; // a cell always holds ≥1 paragraph
    } catch (addrErr) {
        log(`Cell ${label}: could not be addressed (${addrErr.name || 'Error'} — likely merge-covered), skipped`, 'warning');
        return false;
    }
    const paragraphs = cell.body.paragraphs;
    const items = paragraphs.items;

    if (items.length === 1) {
        const range = items[0].getRange(Word.RangeLocation.content);
        range.load('text');
        const revisionRead = queueRevisionRead(range, { paragraph: true });
        await context.sync();
        const current = resolveRevisionRead(range, revisionRead, { paragraph: true });
        if (current.text.trim() === cellPatch.text.trim()) return false;
        try {
            // The outer scope owns the tracking mode for the whole patch.
            const diffOptions = { trackChanges: false, paragraph: true, verificationParagraph: items[0] };
            if (hasCjk(current.text) || hasCjk(cellPatch.text)) {
                await applyCharDiffStrategy(context, range, current.text, cellPatch.text, log, diffOptions);
            } else {
                await applyTokenMapStrategy(context, range, current.text, cellPatch.text, log, diffOptions);
            }
        } catch (diffErr) {
            if (diffErr instanceof RevisionSafetyError || current.hasRevisions) throw diffErr;
            // Loses edit granularity but never leaves a failed apply.
            log(`Cell ${label}: granular diff failed (${diffErr.message}), replacing cell text`, 'warning');
            range.insertText(cellPatch.text, Word.InsertLocation.replace);
            await context.sync();
        }
        return true;
    }

    const whole = items[0].getRange(Word.RangeLocation.content)
        .expandTo(items[items.length - 1].getRange(Word.RangeLocation.content));
    whole.load('text');
    const reads = items.map((p) => queueRevisionRead(p, { paragraph: true }));
    await context.sync();
    const states = items.map((p, i) => resolveRevisionRead(p, reads[i], { paragraph: true }));
    if (states.map((s) => s.text).join('\n').trim() === cellPatch.text.trim()) return false;

    const newLines = cellPatch.text.split(/\r?\n/);
    if (newLines.length === items.length) {
        for (let i = 0; i < items.length; i++) {
            const paraRange = items[i].getRange(Word.RangeLocation.content);
            paraRange.load('text');
            await context.sync();
            const before = states[i].text;
            if (before !== newLines[i]) {
                if (states[i].hasRevisions) {
                    const strategy = hasCjk(before) || hasCjk(newLines[i]) ? applyCharDiffStrategy : applyTokenMapStrategy;
                    await strategy(context, paraRange, before, newLines[i], log,
                        { trackChanges: false, paragraph: true, verificationParagraph: items[i] });
                } else paraRange.insertText(newLines[i], Word.InsertLocation.replace);
            }
        }
        await context.sync();
        return true;
    }

    if (states.some((state) => state.hasRevisions)) {
        throw new RevisionSafetyError('Cannot change paragraph boundaries in a cell with unresolved revisions. Revise its paragraphs in place.');
    }
    log(`Cell ${label}: paragraph count changed (${items.length} → ${newLines.length}); replacing content as one paragraph`, 'warning');
    whole.insertText(newLines.join(' '), Word.InsertLocation.replace);
    await context.sync();
    return true;
}

/**
 * Applies an amendment for a mixed paragraph+table selection. Works at
 * paragraph granularity — the same LCS/similarity alignment the document
 * pipeline uses (reassembler._alignParagraphs), with the same table guards:
 * insert/delete ops never touch table paragraphs, and in-cell keep edits go
 * through the granular diff strategies.
 *
 * Deliberately NO whole-selection replacement fallback: on a mixed selection
 * that fallback is the failure mode this path exists to prevent. Alignment
 * failures (e.g. truncated model output) propagate to the proposal card
 * instead.
 *
 * @param {object} deps - { appState, log }
 * @param {object} proposal - Result of prepareSelectionAmendment (mixed route)
 * @returns {Promise<boolean>} True when at least one op was written
 */
export async function applyMixedTableAmendment(deps, proposal) {
    const { appState, log } = deps;
    const { amendedText } = proposal;
    const trackChangesEnabled = !!appState.config.trackChangesEnabled;
    let applied = false;

    await Word.run(async (context) => {
        const selection = context.document.getSelection();
        const paragraphs = selection.paragraphs;
        paragraphs.load('items');
        await context.sync();

        const allParaItems = paragraphs.items;
        if (allParaItems.length === 0) throw new Error('No paragraphs found in the selection');

        for (const para of allParaItems) para.load('text');
        const revisionReads = allParaItems.map((p) => queueRevisionRead(p, { paragraph: true }));
        const tableChecks = allParaItems.map((p) => {
            const t = p.parentTableOrNullObject;
            t.load('isNullObject');
            return t;
        });
        await context.sync();

        // Blank paragraphs stay out of the alignment, mirroring the
        // reassembler (the amendment text cannot represent them).
        const paraItems = [];
        const inTable = [];
        const states = [];
        allParaItems.forEach((p, i) => {
            const state = resolveRevisionRead(p, revisionReads[i], { paragraph: true });
            if (state.text.trim() !== '') {
                paraItems.push(p);
                inTable.push(!tableChecks[i].isNullObject);
                states.push(state);
            }
        });
        if (paraItems.length === 0) {
            log('Mixed apply: selection contains only blank paragraphs, nothing to amend');
            return;
        }

        const origTexts = states.map((s) => s.text);
        if (_normalizeSelectionText(origTexts.join('\n')) !== _normalizeSelectionText(proposal.selectionText)
            || proposal.revisionFingerprints?.some((fingerprint, i) => fingerprint && states[i]?.fingerprint !== fingerprint)) {
            throw new RevisionSafetyError('Mixed selection changed since staging. Generate a fresh proposal.');
        }
        const amendedLines = _normalizeLineEndings(amendedText).split('\n');
        while (amendedLines.length > 0 && amendedLines[amendedLines.length - 1].trim() === '') amendedLines.pop();
        while (amendedLines.length > 0 && amendedLines[0].trim() === '') amendedLines.shift();

        // Truncation guard (same 30% rule as the document pipeline).
        const origTotalChars = origTexts.reduce((sum, t) => sum + t.length, 0);
        const amendedTotalChars = amendedLines.reduce((sum, t) => sum + t.length, 0);
        if (origTotalChars > 0 && amendedTotalChars < origTotalChars * 0.3) {
            throw new Error(`LLM output appears truncated (${amendedTotalChars} chars vs ${origTotalChars} original) — refusing to apply`);
        }

        if (origTexts.length === amendedLines.length &&
            origTexts.every((t, i) => t.trim() === amendedLines[i].trim())) {
            return; // no changes
        }

        const alignment = origTexts.length === amendedLines.length && states.some((s) => s.hasRevisions)
            ? origTexts.map((_text, index) => ({ type: 'keep', origIdx: index, newIdx: index }))
            : _alignParagraphs(origTexts, amendedLines);
        if (alignment.some((op) => op.type === 'delete' && states[op.origIdx].hasRevisions)) {
            throw new RevisionSafetyError('Cannot delete a paragraph with unresolved revisions. Revise its text in place.');
        }

        if (Word.ChangeTrackingMode) {
            context.document.changeTrackingMode = trackChangesEnabled
                ? Word.ChangeTrackingMode.trackAll
                : Word.ChangeTrackingMode.off;
        }

        try {
            // Reverse document order so ops never invalidate later indexes.
            for (const op of [...alignment].reverse()) {
                if (op.type === 'keep') {
                    const origText = origTexts[op.origIdx];
                    const newText = amendedLines[op.newIdx];
                    if (origText.trim() === newText.trim()) continue;

                    const paraRange = paraItems[op.origIdx].getRange(Word.RangeLocation.content);
                    paraRange.load('text');
                    await context.sync();
                    try {
                        const diffOptions = { trackChanges: false, paragraph: true, verificationParagraph: paraItems[op.origIdx] };
                        if (hasCjk(origText) || hasCjk(newText)) {
                            await applyCharDiffStrategy(context, paraRange, origText, newText.trim(), log, diffOptions);
                        } else {
                            await applyTokenMapStrategy(context, paraRange, origText, newText.trim(), log, diffOptions);
                        }
                    } catch (diffErr) {
                        if (diffErr instanceof RevisionSafetyError || states[op.origIdx].hasRevisions) throw diffErr;
                        // Content-range replacement loses run formatting but is
                        // structurally safe even inside a cell (the cell mark
                        // is outside the content range).
                        log(`Para ${op.origIdx}: granular diff failed (${diffErr.message}), using paragraph text replacement`, 'warning');
                        paraRange.insertText(newText.trim(), Word.InsertLocation.replace);
                        await context.sync();
                    }
                    applied = true;
                } else if (op.type === 'delete') {
                    if (inTable[op.origIdx]) {
                        log(`Para ${op.origIdx}: skipping delete — paragraph is inside a table`, 'warning');
                        continue;
                    }
                    paraItems[op.origIdx].delete();
                    applied = true;
                } else if (op.type === 'insert') {
                    const insertText = amendedLines[op.newIdx].trim();
                    if (!insertText) continue;
                    let anchorOrigIdx = -1;
                    const opIndex = alignment.indexOf(op);
                    for (let k = opIndex - 1; k >= 0; k--) {
                        if (alignment[k].origIdx !== undefined) {
                            anchorOrigIdx = alignment[k].origIdx;
                            break;
                        }
                    }
                    if (anchorOrigIdx >= 0 && inTable[anchorOrigIdx]) {
                        log(`Skipping insert after para ${anchorOrigIdx} — anchor is inside a table`, 'warning');
                        continue;
                    }
                    if (anchorOrigIdx >= 0) {
                        paraItems[anchorOrigIdx].insertParagraph(insertText, Word.InsertLocation.after);
                    } else {
                        paraItems[0].insertParagraph(insertText, Word.InsertLocation.before);
                    }
                    applied = true;
                }
            }
            await context.sync();
        } finally {
            if (Word.ChangeTrackingMode) {
                context.document.changeTrackingMode = Word.ChangeTrackingMode.off;
                await context.sync();
            }
        }
    });

    return applied;
}

/**
 * Instruction wording that implies the model must invent cell content rather
 * than just size an empty grid. Chinese verbs/nouns plus common English ones.
 */
const TABLE_CONTENT_HINT_RE = /填充|填写|填入|内容|数据|资料|根据|依据|基于|\b(?:fill|content|data|populate|based on)\b/i;

/**
 * Drafts a table creation proposal — the prepare half of the staged table
 * flow (see lib/table-ops.js for the creation protocol).
 *
 * Two paths:
 * - Explicit dimensions with no content request ("插入一个三行三列的表格")
 *   resolve deterministically to an empty grid — no model round-trip, so the
 *   result cannot be derailed by a hallucinating model.
 * - Anything else (content hints like 填充/数据, or no dimensions at all) goes
 *   to the model with the strict JSON contract. When dimensions were inferred,
 *   they are restated as a hard constraint and the model's output is checked
 *   against them; a mismatch rejects the whole proposal.
 *
 * @param {object} deps - { appState, log }
 * @param {object} args
 * @param {string} args.instruction - The user's table creation request
 * @param {function} [args.onToken] - Called with each streamed content token
 * @param {function} [args.onReasoning] - Called with each streamed thinking token
 * @returns {Promise<{ instruction: string, spec: object, model: string|null,
 *   warnings: string[] }>} `model` is null on the deterministic empty-table path
 */
export async function prepareTableProposal(deps, { instruction, onToken, onReasoning, signal } = {}) {
    const { appState, log } = deps;
    const text = String(instruction || '').trim();
    if (!text) {
        throw new Error('No table instruction given.');
    }

    const inferred = inferTableCreationSpec(text);
    if (inferred && !TABLE_CONTENT_HINT_RE.test(text)) {
        const rowCount = inferred.rows.length;
        const columnCount = inferred.rows[0].length;
        log(`Table request has explicit dimensions (${rowCount}×${columnCount}) — drafting an empty table without the model.`, 'info');
        return { instruction: text, spec: inferred, model: null, warnings: [] };
    }

    // Content-bearing or dimensionless request: the model fills the grid.
    // Selection is the preferred context — a selected table renders as a
    // markdown grid (structure survives), plain selections pass through as
    // flat text; fall back to the document.
    const tableContext = await readSelectionTableContext({ log });
    let scopeText = tableContext
        ? tableContext.contextText
        : (await readSelectionSnippet()).trim();
    if (!scopeText) {
        const richness = (appState.config.docExtraction || {}).richness || 'structured';
        log('Extracting document text for context...', 'info');
        scopeText = await extractDocumentStructured({ richness });
    }

    let promptInstruction = text;
    let expectedDimensions = null;
    if (inferred) {
        expectedDimensions = { rowCount: inferred.rows.length, columnCount: inferred.rows[0].length };
        promptInstruction =
            `The table MUST have exactly ${expectedDimensions.rowCount} rows and ` +
            `${expectedDimensions.columnCount} columns. ${text}`;
    }

    const prompt = buildTableCreationPrompt(promptInstruction, scopeText);
    const backendConfig = getActiveBackendConfig(appState);
    log(`Generating table content [${backendConfig.model}]...`, 'info');
    const rawResponse = await deps.sendActionRequest(deps, backendConfig, prompt, { onToken, onReasoning, signal });

    const parsed = parseTableCreationResponse(rawResponse);
    if (!parsed.spec) {
        const first = parsed.errors[0];
        throw new Error(
            `表格内容生成失败（${first ? first.message : '模型未返回有效的表格 JSON'}）。` +
            '建议：在指令中明确行数和列数（如"三行三列"），或简化单元格内容后重试。'
        );
    }
    for (const warning of parsed.warnings) {
        log(`Table creation: ${warning.message}`, 'warning');
    }

    if (expectedDimensions) {
        const { rowCount, columnCount } = expectedDimensions;
        const actual = { rowCount: parsed.spec.rows.length, columnCount: parsed.spec.rows[0].length };
        if (actual.rowCount !== rowCount || actual.columnCount !== columnCount) {
            throw new Error(
                `模型返回了 ${actual.rowCount}×${actual.columnCount} 的表格，与要求的 ${rowCount}×${columnCount} 不符。` +
                '请重试，或改用不要求生成内容的建表指令（如"插入一个三行三列的表格"）。'
            );
        }
    }

    return {
        instruction: text,
        spec: parsed.spec,
        model: backendConfig.model,
        warnings: parsed.warnings.map((w) => w.message),
    };
}

/**
 * Applies a prepared table creation proposal: inserts one native Word table at
 * the spec's position (document start/end, or before/after the selection).
 *
 * Tracking: a table insertion is recorded as a revision only on hosts with
 * structural-revision support (see lib/platform.js). Elsewhere the insert runs
 * untracked with a warning rather than risking a half-tracked table. Tracking
 * state is always restored to off afterwards.
 *
 * @param {object} deps - { appState, log }
 * @param {object} proposal - Result of prepareTableProposal
 * @returns {Promise<{ inserted: boolean, rowCount: number, columnCount: number,
 *   tracked: boolean, warnings: string[] }>}
 */
export async function applyTableProposal(deps, proposal) {
    const { appState, log } = deps;
    // Re-validate at apply time: the proposal may have round-tripped through
    // session persistence, so its spec is treated as untrusted input again.
    const validation = validateTableCreationSpec(proposal && proposal.spec);
    if (!validation.spec) {
        const first = validation.errors[0];
        throw new Error(`表格提案无效，无法应用（${first ? first.message : 'unknown error'}）。请重新起草表格。`);
    }
    const spec = validation.spec;
    const rowCount = spec.rows.length;
    const columnCount = spec.rows[0].length;
    const warnings = validation.warnings.map((w) => w.message);

    const trackChanges = !!appState.config.trackChangesEnabled;
    const canTrackInsert = trackChanges && supportsTrackedRowOps(appState.platform);
    if (trackChanges && !canTrackInsert) {
        const warning = `Table insertion cannot be tracked as a revision on this host (${appState.platform}) — applied directly.`;
        warnings.push(warning);
        log(warning, 'warning');
    }

    await Word.run(async (context) => {
        const trackingApiAvailable = !!Word.ChangeTrackingMode;
        if (trackingApiAvailable) {
            context.document.changeTrackingMode = canTrackInsert
                ? Word.ChangeTrackingMode.trackAll
                : Word.ChangeTrackingMode.off;
        } else if (trackChanges) {
            warnings.push('Change tracking is unavailable on this host — the table was inserted directly.');
        }

        try {
            let table;
            if (spec.position === 'start' || spec.position === 'end') {
                table = context.document.body.insertTable(
                    rowCount, columnCount, Word.InsertLocation[spec.position], spec.rows
                );
            } else {
                // before/after anchor on the current selection.
                table = context.document.getSelection().insertTable(
                    rowCount, columnCount, Word.InsertLocation[spec.position], spec.rows
                );
            }
            if (Word.BuiltInStyleName && Word.BuiltInStyleName.tableGrid !== undefined) {
                table.styleBuiltIn = Word.BuiltInStyleName.tableGrid;
            }
            if (spec.headerRowCount > 0) {
                table.headerRowCount = spec.headerRowCount;
            }
            if (spec.autoFit && typeof table.autoFitWindow === 'function') {
                table.autoFitWindow();
            }
            await context.sync();
        } finally {
            // Later turns must not inherit tracking state.
            if (trackingApiAvailable) {
                context.document.changeTrackingMode = Word.ChangeTrackingMode.off;
                await context.sync();
            }
        }
    });

    log(`Inserted a ${rowCount}×${columnCount} table (${spec.position}).`, 'success');
    return { inserted: true, rowCount, columnCount, tracked: canTrackInsert, warnings };
}
