/**
 * Word Actions Module
 *
 * The document/LLM operations behind chat turns, adapted from the original
 * taskpane.js handlers. Differences from the old handlers:
 *
 * - Handlers take explicit args ({ category, promptTemplate, commentInstructions })
 *   instead of reading the PromptManager active prompt and DOM fields.
 * - Selection-scope amendment runs are split into prepare (LLM call, returns
 *   the proposed edit) and apply (diff into Word as tracked changes), so the
 *   chat UI can stage a proposal card between the two.
 * - Document-scope runs take an onProgress callback and return citation data
 *   instead of writing to a fixed progress bar.
 *
 * Dependencies flow in via a `deps` object: { appState, log, logWithRetry,
 * updateStatusBar }. Module-level code is side-effect free; Word/Office
 * globals are only touched inside functions.
 *
 * @module word-actions
 */

import { applyTokenMapStrategy, applySentenceDiffStrategy } from '../lib/word-diff/index.js';
import { hasCjk, applyCharDiffStrategy } from '../lib/word-diff/char-diff.js';
import { sendPrompt, sendPromptStream, stripMarkdown } from '../lib/llm-client.js';
import { buildTableUserPrompt, parseTablePatchResponse } from '../lib/table-patch.js';
import {
    inferTableCreationSpec, buildTableCreationPrompt, parseTableCreationResponse,
    validateTableCreationSpec,
} from '../lib/table-ops.js';
import { supportsTrackedRowOps } from '../lib/platform.js';
import { fireCommentRequest } from '../lib/comment-request.js';
import { extractAllComments, extractDocumentStructured, estimateTokenCount, extractTrackedChanges, extractCommentsOnRange } from '../lib/comment-extractor.js';
import { formatSelectionWithComments } from '../lib/selection-with-comments.js';
import { formatTableMarkdown, formatMixedContext, formatCursorContext } from '../lib/selection-context.js';
import { createSummaryDocument, buildSummaryHtml } from '../lib/document-generator.js';
import { parseDelimitedResponse, buildFallbackClassificationPrompt } from '../lib/response-parser.js';
import { parseDocument } from '../lib/document-parser.js';
import { chunkDocument } from '../lib/document-chunker.js';
import { extractContext } from '../lib/context-extractor.js';
import { processChunksParallel } from '../lib/orchestrator.js';
import { bookmarkChunkRanges, applyChunkResults, cleanupBookmarks, _alignParagraphs, _normalizeLineEndings } from '../lib/reassembler.js';
import { buildFormatPrompt, parseFormatOps } from '../lib/format-ops.js';
import {
    buildIllustrationPrompt, parseIllustration, sanitizeSvg,
    ensureSvgDimensions, svgDimensions, illustrationPositionFromInstruction,
    illustrationPositionLabel,
} from '../lib/illustration.js';
import { buildPlanPrompt, parsePlan } from '../lib/task-planner.js';
import { getActiveBackendConfig } from './app-state.js';

/**
 * Flattens a chat-completions messages array into a single prompt string
 * (system content first), matching the old sendPrompt compatibility shim.
 *
 * @param {Array<{role: string, content: string}>} messages
 * @returns {string}
 * @private
 */
function _flattenMessages(messages) {
    if (messages.length >= 2 && messages[0].role === 'system') {
        return messages[0].content + '\n\n' + messages.slice(1).map(m => m.content).join('\n\n');
    }
    return messages.map(m => m.content).join('\n\n');
}

/**
 * Builds a PromptManager-compatible shim that serves `template` as the active
 * prompt for `category` while delegating everything else (context prompt,
 * other categories) to the real promptManager. Used to run the existing
 * orchestrator/comment pipelines against an explicit skill template.
 *
 * @param {object} promptManager - Real PromptManager instance
 * @param {string} category - Pipeline category the template belongs to
 * @param {string} template - Explicit prompt template
 * @returns {object} PromptManager-compatible shim
 */
export function makePromptShim(promptManager, category, template) {
    return {
        getActivePrompt(cat) {
            if (cat === category) {
                return { id: `skill-${category}`, name: 'Skill template', template };
            }
            return promptManager.getActivePrompt(cat);
        },
        getActiveMode() {
            return category;
        },
        composeMessages: (selectionText, cat) =>
            cat === category
                ? promptManager.composeMessages(selectionText, cat, template)
                : promptManager.composeMessages(selectionText, cat),
    };
}

/**
 * Reads the current selection, optionally enriched with anchored comments
 * (config.includeCommentsInSelection). Throws when the selection is empty.
 *
 * @param {object} deps - { appState, log }
 * @returns {Promise<{ selectionText: string, plainSelectionText: string }>}
 */
export async function readSelectionText(deps) {
    const { appState, log } = deps;
    const includeComments = !!appState.config.includeCommentsInSelection;
    let selectionText = '';
    let plainSelectionText = '';
    let enrichmentError = null;

    await Word.run(async (context) => {
        const selection = context.document.getSelection();
        selection.load('text');
        // OOXML fetch only when enrichment is requested (toggle ON) — saves a sync round-trip on the default path.
        const ooxmlResult = includeComments ? selection.getOoxml() : null;
        await context.sync();
        if (!selection.text || !selection.text.trim()) {
            throw new Error('Please select some text first.');
        }
        plainSelectionText = selection.text;

        if (!includeComments) {
            selectionText = plainSelectionText;
            return;
        }

        let comments = [];
        try {
            comments = await extractCommentsOnRange(context, selection);
        } catch (err) {
            console.error('[readSelectionText] extractCommentsOnRange failed', { err });
            enrichmentError = err;
        }

        if (!enrichmentError) {
            try {
                selectionText = formatSelectionWithComments(ooxmlResult.value, comments);
            } catch (err) {
                console.error('[readSelectionText] formatSelectionWithComments failed', { err });
                enrichmentError = err;
            }
        }
    });

    if (includeComments && enrichmentError) {
        log(`Comment enrichment failed (${enrichmentError.message}); falling back to plain selection.`, 'warning');
        selectionText = plainSelectionText;
    } else if (includeComments && selectionText.length > plainSelectionText.length) {
        log(`Selection enriched with comment threads (+${selectionText.length - plainSelectionText.length} chars)`, 'info');
    }

    return { selectionText, plainSelectionText };
}

/**
 * Detects whether the current selection spans MULTIPLE table cells and, if
 * so, extracts the covered region as coordinate-addressed cell data for the
 * table patch protocol (lib/table-patch.js).
 *
 * Returns null for selections outside a table or contained within a single
 * cell — those stay on the flat-text pipelines, whose in-cell text edits
 * track natively and granularly. Only multi-cell regions need the protocol:
 * flattening them would leak cell marks into the diff strategies.
 *
 * Word cell selections are rectangular, so the zero-width selection start
 * and end points each sit in the top-left / bottom-right covered cell —
 * two coordinate reads bound the whole region (no per-cell location
 * compares, no row-collection iteration). An endpoint that lands exactly
 * ON a table boundary (whole-table selections hit both) resolves to the
 * unit OUTSIDE the table and reads as a null cell; such endpoints clamp
 * to the table edge instead of failing the read. Merged layouts are
 * probed via getCell anchors; hosts that throw on merge-covered
 * coordinates degrade to "merged, layout unknown" (mergedUnknown=true,
 * empty shadowKeys) with apply-time validation instead of failing.
 *
 * @param {object} deps - { appState, log }
 * @returns {Promise<null | {rowCount: number, colCount: number,
 *   bounds: {startRow: number, endRow: number, startCol: number, endCol: number},
 *   cells: Array<{row: number, col: number, text: string}>,
 *   values: string[][]}>}
 *   Row/col in cells and bounds are 1-based absolute table coordinates;
 *   values is the full-table matrix (0-based) for no-op/staleness checks.
 */
export async function readSelectionTableRegion(deps) {
    const { log } = deps;
    let region = null;
    await Word.run(async (context) => {
        const selection = context.document.getSelection();
        const table = selection.parentTableOrNullObject;
        const anchorCell = selection.parentTableCellOrNullObject;
        table.load('isNullObject');
        anchorCell.load('isNullObject');
        await context.sync();

        // Outside any table, or fully inside one cell: not table mode.
        if (table.isNullObject || !anchorCell.isNullObject) return;

        const startCell = selection.getRange(Word.RangeLocation.start).parentTableCellOrNullObject;
        const endCell = selection.getRange(Word.RangeLocation.end).parentTableCellOrNullObject;
        // Office.js exposes the column coordinate as `cellIndex` (there is no
        // `columnIndex` on TableCell) — the wrong property silently reads
        // undefined and collapses the region to a single column.
        startCell.load('isNullObject,rowIndex,cellIndex');
        endCell.load('isNullObject,rowIndex,cellIndex');
        table.load('rowCount,values,isUniform');
        await context.sync();

        // Boundary-ambiguous selections: a zero-width range sitting exactly
        // ON a table boundary belongs to the OUTSIDE unit in Word's position
        // model (the paragraph before the table for a selection starting at
        // the very top of the first cell; the row-end mark / following
        // paragraph for one ending at the very end of the last cell).
        // Selecting the WHOLE table hits both. The containment checks above
        // (parentTable non-null, anchorCell null) already prove a multi-cell
        // in-table selection, so clamp missing endpoints to the table edge
        // instead of erroring: resolved corners keep their exact coordinates.
        const values = table.values || [];
        const colCount = values[0] ? values[0].length : 0;
        const startRow = Math.min(
            startCell.isNullObject ? 0 : startCell.rowIndex,
            endCell.isNullObject ? (table.rowCount || values.length) - 1 : endCell.rowIndex
        );
        const endRow = Math.max(
            startCell.isNullObject ? 0 : startCell.rowIndex,
            endCell.isNullObject ? (table.rowCount || values.length) - 1 : endCell.rowIndex
        );
        const startCol = Math.min(
            startCell.isNullObject ? 0 : startCell.cellIndex,
            endCell.isNullObject ? colCount - 1 : endCell.cellIndex
        );
        const endCol = Math.max(
            startCell.isNullObject ? 0 : startCell.cellIndex,
            endCell.isNullObject ? colCount - 1 : endCell.cellIndex
        );
        if (startCell.isNullObject || endCell.isNullObject) {
            log('Selection endpoint(s) sat on a table boundary — clamping the region to the table edge.', 'warning');
        }

        // Merged tables: Word JS has no merge/unmerge surface and row ops
        // assume a uniform grid, but CONTENT edits are safe on merge anchors
        // — getCell(r, c) resolves a coordinate inside a merged region to the
        // merged cell, whose rowIndex/cellIndex are its anchor slot. Probe
        // every covered coordinate once: a probe resolving to coordinates
        // other than its own is a read-only shadow of that merge anchor.
        // Some hosts throw ItemNotFound for getCell on merge-covered
        // coordinates instead of resolving to the anchor — the probe sync is
        // guarded, and a failure degrades to "merged, layout unknown" (edits
        // validated only at apply) rather than failing the whole turn.
        const merged = table.isUniform === false;
        const shadowKeys = new Set();
        let mergedUnknown = false;
        if (merged) {
            const probes = [];
            for (let r = startRow; r <= endRow; r++) {
                for (let c = startCol; c <= endCol; c++) {
                    const probe = table.getCell(r, c);
                    probe.load('rowIndex,cellIndex');
                    probes.push({ r, c, probe });
                }
            }
            try {
                await context.sync();
            } catch (probeErr) {
                mergedUnknown = true;
                log(`Merge layout could not be probed (${probeErr.name || 'Error'}: ${probeErr.message}) — continuing with unknown merge layout; invalid edits will be skipped at apply.`, 'warning');
            }
            if (!mergedUnknown) {
                for (const { r, c, probe } of probes) {
                    if (probe.rowIndex !== r || probe.cellIndex !== c) {
                        shadowKeys.add(`${r + 1},${c + 1}`);
                    }
                }
                log(`Table contains merged cells — ${shadowKeys.size} covered slot(s) read-only, row ops disabled; anchor-cell text edits stay available.`, 'warning');
            }
        }

        const cells = [];
        for (let r = startRow; r <= endRow; r++) {
            for (let c = startCol; c <= endCol; c++) {
                const entry = { row: r + 1, col: c + 1, text: (values[r] && values[r][c]) || '' };
                if (shadowKeys.has(`${r + 1},${c + 1}`)) entry.merged = true;
                cells.push(entry);
            }
        }

        // Style snapshot for the tool loop's style tools (get_state reporting
        // and before/after context). Read after the merge probe so the probe
        // sync ordinal is unchanged; best effort — a host lacking a property
        // degrades to an unknown-style snapshot, never a failed read.
        const style = await _readTableStyleSnapshot(context, table);
        region = {
            rowCount: table.rowCount,
            colCount,
            bounds: { startRow: startRow + 1, endRow: endRow + 1, startCol: startCol + 1, endCol: endCol + 1 },
            cells,
            values,
            merged,
            style,
            ...(merged ? { shadowKeys, mergedUnknown } : {}),
        };
    });
    if (region) {
        const first = region.cells[0];
        const last = region.cells[region.cells.length - 1];
        log(`Table selection: ${region.cells.length} cell(s), R${first.row}C${first.col} → R${last.row}C${last.col}`, 'info');
    }
    return region;
}

/**
 * Reads EVERY table in the document body and returns an array of regions
 * covering each whole table — used by document-scope `table_management`
 * turns where no selection is available. Regions are numbered by document
 * order (tableIndex 1-based). Returns an empty array when the document has
 * no tables, or null when a table is malformed/empty (callers treat null
 * as "nothing to operate on").
 *
 * @param {object} deps - { log }
 * @returns {Promise<Array<object>|null>}
 */
export async function readDocumentTableRegions(deps) {
    const { log } = deps;
    let regions = [];
    await Word.run(async (context) => {
        const tables = context.document.body.tables;
        tables.load('items');
        await context.sync();
        const items = tables.items || [];
        if (items.length === 0) return;

        /** @type {Array<object>} */
        const found = [];
        for (const table of items) {
            table.load('isNullObject,rowCount,values,isUniform');
            await context.sync();

            const values = table.values || [];
            const rowCount = table.rowCount || values.length;
            const colCount = values[0] ? values[0].length : 0;
            if (rowCount === 0 || colCount === 0) {
                log(`Document table (${found.length + 1}) is empty/malformed — skipped.`, 'warning');
                continue;
            }

            const cells = [];
            for (let r = 0; r < rowCount; r++) {
                for (let c = 0; c < colCount; c++) {
                    cells.push({ row: r + 1, col: c + 1, text: (values[r] && values[r][c]) || '' });
                }
            }

            const style = await _readTableStyleSnapshot(context, table);

            found.push({
                tableIndex: found.length + 1,
                rowCount,
                colCount,
                bounds: { startRow: 1, endRow: rowCount, startCol: 1, endCol: colCount },
                cells,
                values,
                merged: false,
                style,
            });
        }
        regions = found;
    });
    if (regions.length > 0) {
        log(`Document has ${regions.length} table(s); this turn manages all of them by tableIndex.`, 'info');
    }
    return regions;
}

/**
 * Reads the table's current styling into a plain snapshot for the tool
 * loop: built-in/custom style name, alignment, header rows, banding flags,
 * shading, whole-table font, and the six border locations. Best effort —
 * any failure degrades to null ("snapshot unavailable") so a host without
 * one of the properties still serves the read.
 *
 * @param {Word.RequestContext} context
 * @param {Word.Table} table - Loaded table proxy
 * @returns {Promise<object|null>}
 * @private
 */
async function _readTableStyleSnapshot(context, table) {
    try {
        table.load(
            'styleBuiltIn,style,alignment,horizontalAlignment,verticalAlignment,' +
            'headerRowCount,styleBandedRows,styleBandedColumns,styleFirstColumn,' +
            'styleLastColumn,styleTotalRow,shadingColor,width'
        );
        if (table.font) table.font.load('bold,name,size,color');
        const borderSpots = [
            ['top', 'top', 'Top'], ['bottom', 'bottom', 'Bottom'],
            ['left', 'left', 'Left'], ['right', 'right', 'Right'],
            ['insideH', 'insideHorizontal', 'InsideHorizontal'],
            ['insideV', 'insideVertical', 'InsideVertical'],
        ];
        const borderObjects = typeof table.getBorder === 'function'
            ? borderSpots.map(([key, enumKey, fallback]) => {
                // The API takes the enum value ("Top") or its string literal;
                // hosts without the enum still accept the capitalized literal.
                const location = (globalThis.Word && globalThis.Word.BorderLocation
                    && globalThis.Word.BorderLocation[enumKey]) || fallback;
                const border = table.getBorder(location);
                border.load('type,color,width');
                return { key, border };
            })
            : [];
        await context.sync();
        const borders = {};
        for (const { key, border } of borderObjects) {
            borders[key] = border.type === undefined && border.color === undefined
                ? null
                : { type: border.type || null, color: border.color || null, width: border.width || null };
        }
        return {
            styleBuiltIn: table.styleBuiltIn || null,
            style: table.style || null,
            alignment: table.alignment || null,
            horizontalAlignment: table.horizontalAlignment || null,
            verticalAlignment: table.verticalAlignment || null,
            headerRowCount: table.headerRowCount || 0,
            bandedRows: table.styleBandedRows,
            bandedColumns: table.styleBandedColumns,
            firstColumn: table.styleFirstColumn,
            lastColumn: table.styleLastColumn,
            totalRow: table.styleTotalRow,
            shadingColor: table.shadingColor || null,
            widthPt: table.width || null,
            font: table.font
                ? { bold: table.font.bold, name: table.font.name, size: table.font.size, color: table.font.color }
                : null,
            borders: borderObjects.length ? borders : null,
        };
    } catch (_styleErr) {
        return null;
    }
}

/**
 * Detects a MIXED selection: not contained in a table (readSelectionTableRegion
 * covers that case) but overlapping one — e.g. a caption paragraph, the whole
 * table, and the note paragraph below it. The flat range-diff strategies
 * would destroy the table (whole-selection replacement leaks cell marks), so
 * such selections switch to paragraph-granular handling: the prompt text and
 * the apply-time alignment both work on whole paragraphs (one line per
 * paragraph/cell), and table paragraphs are guarded against structural ops.
 *
 * Paragraph granularity means a partial selection of a boundary paragraph
 * acts on the whole paragraph — the resulting redlines are individually
 * rejectable, which is the safe trade-off against table corruption.
 *
 * @param {object} deps - { appState, log }
 * @returns {Promise<null | {selectionText: string, paraCount: number, tableParaCount: number}>}
 *   selectionText is the non-blank paragraph texts joined with '\n'.
 */
export async function readMixedTableSelection(deps) {
    const { log } = deps;
    let mixed = null;
    await Word.run(async (context) => {
        const selection = context.document.getSelection();
        const table = selection.parentTableOrNullObject;
        table.load('isNullObject');
        await context.sync();
        // Inside a table (single cell or multi-cell region): other routes own it.
        if (!table.isNullObject) return;

        const paragraphs = selection.paragraphs;
        paragraphs.load('items');
        await context.sync();
        if (paragraphs.items.length === 0) return;

        for (const para of paragraphs.items) para.load('text');
        const tableChecks = paragraphs.items.map((p) => {
            const t = p.parentTableOrNullObject;
            t.load('isNullObject');
            return t;
        });
        await context.sync();

        const tableParaCount = tableChecks.filter((t) => !t.isNullObject).length;
        if (tableParaCount === 0) return;

        const texts = paragraphs.items
            .map((p) => p.text)
            .filter((t) => t && t.trim() !== '');
        mixed = { selectionText: texts.join('\n'), paraCount: paragraphs.items.length, tableParaCount };
    });
    if (mixed) {
        log(`Mixed selection: ${mixed.paraCount} paragraph(s), ${mixed.tableParaCount} inside table(s) — paragraph-granular mode`, 'info');
    }
    return mixed;
}

/**
 * Reads the current selection as STRUCTURED prompt context: table content
 * becomes a markdown grid (lib/selection-context.js) instead of the
 * flattened cell string selection.text gives. Read-only, so unlike the
 * edit-side readers above there is no isUniform constraint — merged cells
 * simply render as they read.
 *
 * Shapes, mirroring the edit-side routing:
 * - selection contained in a table (single or multi cell) → the WHOLE
 *   table matrix, with the covered region noted in the marker. Questions
 *   about selected data cells still need the header row for meaning.
 * - selection overlapping but not contained in a table (caption + table +
 *   note) → document-ordered paragraphs + table grids.
 * - plain text selection / no table overlap → null (caller keeps the flat
 *   selection text).
 *
 * Any failure resolves to null so the QA path falls back to plain text
 * instead of erroring the turn.
 *
 * @param {object} deps - { log }
 * @returns {Promise<null | {kind: 'table', contextText: string, rowCount: number,
 *   colCount: number} | {kind: 'mixed', contextText: string, tableCount: number}>}
 */
export async function readSelectionTableContext(deps) {
    const { log } = deps;
    let result = null;
    try {
        await Word.run(async (context) => {
            const selection = context.document.getSelection();
            const table = selection.parentTableOrNullObject;
            const anchorCell = selection.parentTableCellOrNullObject;
            table.load('isNullObject');
            anchorCell.load('isNullObject');
            await context.sync();

            // Contained in a table (single- or multi-cell selection).
            if (!table.isNullObject) {
                let startCell = null;
                let endCell = null;
                if (!anchorCell.isNullObject) {
                    anchorCell.load('rowIndex,cellIndex');
                } else {
                    // Rectangular covered region bounded by the start/end
                    // cells (same two-coordinate trick as
                    // readSelectionTableRegion).
                    startCell = selection.getRange(Word.RangeLocation.start).parentTableCellOrNullObject;
                    endCell = selection.getRange(Word.RangeLocation.end).parentTableCellOrNullObject;
                    startCell.load('isNullObject,rowIndex,cellIndex');
                    endCell.load('isNullObject,rowIndex,cellIndex');
                }
                table.load('values,isUniform');
                await context.sync();

                const values = table.values || [];
                const rowCount = values.length;
                const colCount = rowCount > 0 && Array.isArray(values[0]) ? values[0].length : 0;

                let note = '';
                if (!anchorCell.isNullObject) {
                    note = `user selected cell R${anchorCell.rowIndex + 1}C${anchorCell.cellIndex + 1}`;
                } else if (!startCell.isNullObject && !endCell.isNullObject) {
                    const startRow = Math.min(startCell.rowIndex, endCell.rowIndex) + 1;
                    const endRow = Math.max(startCell.rowIndex, endCell.rowIndex) + 1;
                    const startCol = Math.min(startCell.cellIndex, endCell.cellIndex) + 1;
                    const endCol = Math.max(startCell.cellIndex, endCell.cellIndex) + 1;
                    note = (startRow === 1 && endRow === rowCount && startCol === 1 && endCol === colCount)
                        ? 'user selected the whole table'
                        : `user selected R${startRow}C${startCol}–R${endRow}C${endCol}`;
                }
                if (table.isUniform === false) {
                    note = note ? `${note}; contains merged cells` : 'contains merged cells';
                }

                const contextText = formatTableMarkdown(values, { note });
                if (contextText) {
                    log(`Selection context: table ${rowCount}×${colCount}${note ? ` (${note})` : ''}`, 'info');
                    result = { kind: 'table', contextText, rowCount, colCount };
                }
                return;
            }

            // Mixed selection: table overlap without containment (mirrors
            // readMixedTableSelection's detection).
            const paragraphs = selection.paragraphs;
            paragraphs.load('items');
            await context.sync();
            if (paragraphs.items.length === 0) return;

            for (const para of paragraphs.items) para.load('text');
            const tableChecks = paragraphs.items.map((p) => {
                const t = p.parentTableOrNullObject;
                t.load('isNullObject');
                return t;
            });
            await context.sync();

            const inTable = tableChecks.map((t) => !t.isNullObject);
            if (!inTable.some(Boolean)) return;

            // Group consecutive table paragraphs into runs. Adjacent Word
            // tables merge unless a paragraph separates them, so each run is
            // one table; its matrix is read from the run's first paragraph.
            const runs = [];
            for (let i = 0; i < inTable.length; i++) {
                if (!inTable[i]) continue;
                const last = runs[runs.length - 1];
                if (last && last.end === i - 1) {
                    last.end = i;
                } else {
                    runs.push({ start: i, end: i });
                }
            }

            const runTables = runs.map((run) => {
                const t = paragraphs.items[run.start].parentTableOrNullObject;
                t.load('values');
                return t;
            });
            await context.sync();

            const parts = [];
            let cursor = 0;
            runs.forEach((run, i) => {
                for (let j = cursor; j < run.start; j++) {
                    parts.push({ type: 'paragraph', text: paragraphs.items[j].text });
                }
                parts.push({ type: 'table', values: runTables[i].values || [] });
                cursor = run.end + 1;
            });
            for (let j = cursor; j < paragraphs.items.length; j++) {
                parts.push({ type: 'paragraph', text: paragraphs.items[j].text });
            }

            const contextText = formatMixedContext(parts);
            if (contextText) {
                log(`Selection context: mixed selection, ${runs.length} table(s) — markdown mode`, 'info');
                result = { kind: 'mixed', contextText, tableCount: runs.length };
            }
        });
    } catch (err) {
        if (log) log(`Table context read failed (${err.message}); using plain selection text.`, 'warning');
        result = null;
    }
    return result;
}

/**
 * Cap on the backward paragraph walk when hunting the nearest preceding
 * heading. Bounds the queued proxy chain (one sync) so a pathological
 * document cannot turn a cursor read into a full-body scan.
 */
const CURSOR_HEADING_WALK_LIMIT = 120;

/**
 * Reads the CURSOR location as prompt context — used when no text is
 * selected, so document-scope answers know where the user is working.
 *
 * Returns null unless the selection is collapsed (a bare caret): a real
 * selection belongs to the selection-context paths instead. Context is
 * the caret's paragraph (clipped) plus the nearest preceding heading,
 * found by a bounded getPreviousOrNullObject chain queued in ONE sync.
 * When the caret itself sits in a heading paragraph, that heading wins
 * (it is the section being read/edited, not the one before it).
 *
 * Any failure resolves to null — QA proceeds without location context.
 *
 * @param {object} deps - { log }
 * @returns {Promise<null | {kind: 'cursor', contextText: string}>}
 */
export async function readCursorContext(deps) {
    const { log } = deps;
    let result = null;
    try {
        await Word.run(async (context) => {
            const selection = context.document.getSelection();
            selection.load('text');
            await context.sync();
            if (selection.text && selection.text.trim()) return; // real selection

            const paragraphs = selection.paragraphs;
            paragraphs.load('items');
            await context.sync();
            const cursorPara = paragraphs.items[0];
            if (!cursorPara) return;

            cursorPara.load('text,styleBuiltIn');
            const table = cursorPara.parentTableOrNullObject;
            table.load('isNullObject');

            // Bounded walk-back chain for the nearest preceding heading:
            // proxy getPreviousOrNullObject calls queue up and all resolve
            // in the single sync below.
            const chain = [];
            let node = cursorPara.getPreviousOrNullObject();
            for (let i = 0; i < CURSOR_HEADING_WALK_LIMIT && node; i++) {
                node.load('isNullObject,text,styleBuiltIn');
                chain.push(node);
                node = node.getPreviousOrNullObject();
            }
            await context.sync();

            const headingOf = (p) => {
                const match = /^Heading([1-9])$/.exec(p.styleBuiltIn || '');
                return match ? { text: (p.text || '').trim(), level: Number(match[1]) } : null;
            };

            // Caret in a heading: that heading IS the current section.
            let heading = headingOf(cursorPara);
            if (!heading) {
                for (const p of chain) {
                    if (p.isNullObject) break; // walked off the body start
                    heading = headingOf(p);
                    if (heading) break;
                }
            }

            const contextText = formatCursorContext({
                paragraphText: cursorPara.text,
                headingText: heading ? heading.text : '',
                headingLevel: heading ? heading.level : 0,
                inTable: !table.isNullObject,
            });
            if (contextText) {
                log(`Cursor context: ${heading ? `section "${heading.text.slice(0, 40)}"` : 'no heading found'}` +
                    `${!table.isNullObject ? ' (in table)' : ''}`, 'info');
                result = { kind: 'cursor', contextText };
            }
        });
    } catch (err) {
        if (log) log(`Cursor context read failed (${err.message}).`, 'warning');
        result = null;
    }
    return result;
}

/**
 * Reads the current selection as plain text (no comment enrichment).
 * Returns '' when the selection is empty or the Word API is unavailable.
 * Used for the live selection preview and for QA-turn context.
 *
 * @returns {Promise<string>}
 */
export async function readSelectionSnippet() {
    try {
        let text = '';
        await Word.run(async (context) => {
            const selection = context.document.getSelection();
            selection.load('text');
            await context.sync();
            text = selection.text || '';
        });
        return text;
    } catch (_err) {
        return '';
    }
}

/**
 * Wraps a raw base64 image payload (as returned by
 * InlinePicture.getBase64ImageSrc) into a data URL, sniffing the MIME type
 * from the base64 magic prefix — Word does not report the encoded format.
 * Already-wrapped values pass through unchanged.
 *
 * @param {string} base64 - Raw base64 (or an existing data: URL)
 * @returns {string} data: URL
 */
export function imageDataUrl(base64) {
    const raw = String(base64 || '');
    if (raw.startsWith('data:')) return raw;
    let mime = 'image/png';
    if (raw.startsWith('/9j/')) mime = 'image/jpeg';
    else if (raw.startsWith('R0lGO')) mime = 'image/gif';
    else if (raw.startsWith('UklGR')) mime = 'image/webp';
    else if (raw.startsWith('iVBOR')) mime = 'image/png';
    else if (raw.startsWith('PHN2Z')) mime = 'image/svg+xml';
    return `data:${mime};base64,${raw}`;
}

/**
 * Cap on selection images read fully (base64 + metadata) per read. Word
 * selections can span dozens of pictures; the preview and the image-tool
 * selection focus only need a handful, and every image costs a base64
 * round-trip.
 */
const MAX_SELECTION_IMAGES = 6;

/**
 * Reads the current selection as FULL content: plain text plus the inline
 * pictures inside the selection (base64 data URL + size + alt text), AND
 * the selection's table shape (multi-cell region + corner coords). So
 * ANY selected content is visible to the add-in — picture selections have
 * empty selection.text, single-cell selections stay on the text path, and
 * multi-cell selections route into the table tool session.
 *
 * The table shape is detected in the SAME Word.run as the image read
 * (a single sync covers both); the full matrix that the table tool loop
 * needs is loaded separately inside prepareTableToolEdit — cheap here.
 *
 * Image-bearing selections are consumed by the image tool session (object
 * reference + tools, see agent-actions); the preview shows thumbnails.
 * Multi-cell selections route into the table tool session (same pattern,
 * TABLE_TOOL turn). Errors resolve to empty content (same contract as
 * readSelectionSnippet).
 *
 * @returns {Promise<{ text: string,
 *   images: Array<{ base64: string, dataUrl: string, width: number, height: number, altText: string }>,
 *   totalImages: number,
 *   hasMultiCellTableRegion: boolean,
 *   tableRegion: null | { startRow: number, endRow: number, startCol: number, endCol: number } }>}
 *   totalImages counts every picture in the selection (images.length ≤
 *   totalImages when the cap truncates). tableRegion is non-null only
 *   when the selection covers multiple cells AND both endpoint cells
 *   resolve cleanly (whole-table boundary selections may return
 *   hasMultiCellTableRegion=true with tableRegion=null — the corner
 *   coords are inferred inside the full region read instead).
 */
export async function readSelectionContent() {
    try {
        let text = '';
        let images = [];
        let totalImages = 0;
        let hasMultiCellTableRegion = false;
        let tableRegion = null;
        await Word.run(async (context) => {
            const selection = context.document.getSelection();
            selection.load('text');
            const pictures = selection.inlinePictures;
            pictures.load('items');
            // Shape probes — same Word.run, no extra sync.
            const parentTable = selection.parentTableOrNullObject;
            const anchorCell = selection.parentTableCellOrNullObject;
            parentTable.load('isNullObject');
            anchorCell.load('isNullObject');
            const startRange = selection.getRange(Word.RangeLocation.start);
            const endRange = selection.getRange(Word.RangeLocation.end);
            const startCell = startRange.parentTableCellOrNullObject;
            const endCell = endRange.parentTableCellOrNullObject;
            startCell.load('isNullObject,rowIndex,cellIndex');
            endCell.load('isNullObject,rowIndex,cellIndex');

            await context.sync();
            text = selection.text || '';

            const items = pictures.items || [];
            totalImages = items.length;
            const shown = items.slice(0, MAX_SELECTION_IMAGES);
            const reads = [];
            for (const pic of shown) {
                pic.load('width,height,altTextDescription');
                reads.push({ pic, b64: pic.getBase64ImageSrc() });
            }
            if (reads.length > 0) await context.sync();
            images = reads
                .filter(({ b64 }) => b64.value)
                .map(({ pic, b64 }) => ({
                    base64: b64.value,
                    dataUrl: imageDataUrl(b64.value),
                    width: pic.width,
                    height: pic.height,
                    altText: pic.altTextDescription || '',
                }));

            // Multi-cell shape: inside a table BUT not wholly inside a single
            // anchor cell — same predicate as readSelectionTableRegion's gate.
            hasMultiCellTableRegion = !parentTable.isNullObject && anchorCell.isNullObject;
            if (hasMultiCellTableRegion) {
                if (!startCell.isNullObject && !endCell.isNullObject) {
                    tableRegion = {
                        startRow: Math.min(startCell.rowIndex, endCell.rowIndex) + 1,
                        endRow: Math.max(startCell.rowIndex, endCell.rowIndex) + 1,
                        startCol: Math.min(startCell.cellIndex, endCell.cellIndex) + 1,
                        endCol: Math.max(startCell.cellIndex, endCell.cellIndex) + 1,
                    };
                }
                // Whole-table boundary selections leave one or both endpoints
                // as null proxies; the full region read in prepareTableToolEdit
                // clamps those to the table edge — handle there.
            }
        });
        return { text, images, totalImages, hasMultiCellTableRegion, tableRegion };
    } catch (_err) {
        return { text: '', images: [], totalImages: 0, hasMultiCellTableRegion: false, tableRegion: null };
    }
}

/**
 * Watches the Word selection and invokes the callback with the current
 * selection CONTENT ({ text, images, totalImages, hasMultiCellTableRegion,
 * tableRegion } — '' / [] when nothing is selected; image-only selections
 * carry empty text with images; multi-cell table regions carry corner
 * coords in tableRegion). Events are debounced so a drag selection fires
 * one Word.run at the end.
 *
 * @param {function(object)} callback - Receives the live selection content
 * @param {object} [opts]
 * @param {number} [opts.debounceMs=200] - Trailing debounce for change events
 * @returns {function()} Unsubscribe function
 */
export function watchSelection(callback, { debounceMs = 200 } = {}) {
    if (typeof Office === 'undefined' || !Office.context || !Office.context.document
        || !Office.EventType || !Office.EventType.DocumentSelectionChanged) {
        return () => {};
    }

    let timer = null;
    const onChange = () => {
        if (timer) clearTimeout(timer);
        timer = setTimeout(async () => {
            timer = null;
            callback(await readSelectionContent());
        }, debounceMs);
    };

    Office.context.document.addHandlerAsync(
        Office.EventType.DocumentSelectionChanged,
        onChange,
        () => {} // registration failure is non-fatal: the preview just stays static
    );

    // Emit the initial state (the user may have selected content before opening the pane)
    onChange();

    return () => {
        if (timer) clearTimeout(timer);
        try {
            Office.context.document.removeHandlerAsync(
                Office.EventType.DocumentSelectionChanged,
                { handler: onChange },
                () => {}
            );
        } catch (_err) { /* best-effort cleanup */ }
    };
}

/**
 * Runs the selection-scope amendment LLM call WITHOUT applying the result.
 * The returned proposal is staged in the chat UI; applySelectionAmendment
 * writes it to the document.
 *
 * @param {object} deps - { appState, log }
 * @param {object} args
 * @param {string} args.promptTemplate - Amendment instruction/template ({selection} placeholder supported)
 * @param {string} [args.commentInstructions] - When non-empty, merged amendment + comment mode
 * @param {function} [args.onToken] - Called with each streamed content token
 * @param {function} [args.onReasoning] - Called with each streamed thinking token
 * @param {AbortSignal} [args.signal] - Cancellation signal (stop button)
 * @returns {Promise<{ selectionText: string, amendedText: string|null, commentText: string|null, model: string }>}
 */
export async function prepareSelectionAmendment(deps, { promptTemplate, commentInstructions, onToken, onReasoning, signal } = {}) {
    const { appState, log } = deps;
    const backendConfig = getActiveBackendConfig(appState);

    // Multi-cell table selections take the coordinate patch protocol — the
    // flat-text pipelines below cannot represent cell boundaries.
    const tableRegion = await readSelectionTableRegion(deps);
    if (tableRegion) {
        return _prepareTableAmendment(deps, {
            tableRegion, promptTemplate, commentInstructions, backendConfig, onToken, onReasoning, signal,
        });
    }

    // Mixed selections (paragraphs + table content) keep the flat prompt but
    // switch to paragraph-granular text, so apply-time alignment can map
    // lines back onto paragraphs without touching table structure.
    const mixed = await readMixedTableSelection(deps);

    const { selectionText } = mixed ? { selectionText: mixed.selectionText } : await readSelectionText(deps);
    log(`Processing selection (${selectionText.length} chars) via ${backendConfig.model}...`, 'info');

    const merged = !!(commentInstructions && commentInstructions.trim());
    const messages = merged
        ? appState.promptManager.composeMergedMessages(selectionText, commentInstructions, promptTemplate)
        : appState.promptManager.composeMessages(selectionText, 'amendment', promptTemplate);

    if (messages.length === 0) {
        throw new Error('No prompt composed — check the skill template');
    }

    if (mixed) {
        // One line per paragraph/cell: the guarded alignment at apply time
        // maps lines back onto paragraphs, so the model must not fuse them.
        messages[messages.length - 1].content += '\n\nNOTE: The selection contains a Word table; each cell appears as its own line. Keep every line present and in the same order, edit text in place, and never merge, split, reorder, or drop lines.';
    }

    const promptText = _flattenMessages(messages);
    const rawResponse = (onToken || onReasoning)
        ? await sendPromptStream(backendConfig, promptText, { onContent: onToken, onReasoning }, log, signal)
        : await sendPrompt(backendConfig, promptText, log, signal);
    log(`LLM response received [${backendConfig.model}]`, 'success');

    if (!merged) {
        return {
            selectionText,
            amendedText: stripMarkdown(rawResponse, log),
            commentText: null,
            model: backendConfig.model,
            ...(mixed ? { mixedTable: true } : {}),
        };
    }

    // Merged mode: parse the ===AMENDMENT=== / ===COMMENT=== protocol,
    // with the fallback classification call preserved from the old handler.
    let parsed = parseDelimitedResponse(rawResponse);

    if (parsed.amendment === null) {
        log('Response missing delimiters, attempting to classify...', 'info');
        const fallbackMessages = buildFallbackClassificationPrompt(rawResponse, selectionText);
        try {
            const fallbackResponse = await sendPrompt(backendConfig, _flattenMessages(fallbackMessages), log, signal);
            parsed = parseDelimitedResponse(fallbackResponse);
            if (parsed.amendment === null) {
                log('Could not split response into amendment and comment', 'warning');
                parsed = { amendment: rawResponse.trim(), comment: null, raw: rawResponse };
            }
        } catch (fallbackError) {
            // A user cancel must propagate, not stage a proposal from
            // partial output (the catch below would swallow it).
            if (fallbackError.name === 'AbortError') throw fallbackError;
            log(`Fallback classification failed: ${fallbackError.message}`, 'warning');
            parsed = { amendment: rawResponse.trim(), comment: null, raw: rawResponse };
        }
    }

    return {
        selectionText,
        amendedText: parsed.amendment ? stripMarkdown(parsed.amendment, log) : null,
        commentText: parsed.comment || null,
        model: backendConfig.model,
        ...(mixed ? { mixedTable: true } : {}),
    };
}

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
 * @private
 */
async function _prepareTableAmendment(deps, { tableRegion, promptTemplate, commentInstructions, backendConfig, onToken, onReasoning, signal }) {
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

    const promptText = _flattenMessages(messages);
    const rawResponse = (onToken || onReasoning)
        ? await sendPromptStream(backendConfig, promptText, { onContent: onToken, onReasoning }, log, signal)
        : await sendPrompt(backendConfig, promptText, log, signal);
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
 * Applies a prepared selection amendment to the document as tracked changes
 * (per config.trackChangesEnabled), then inserts the optional comment.
 *
 * @param {object} deps - { appState, log }
 * @param {object} proposal - Result of prepareSelectionAmendment
 * @returns {Promise<object|undefined>} Table patches resolve to
 *   { cellsApplied, cellsSkipped, rowOpsApplied, warnings }; the flat text
 *   route resolves { skipped: true, reason } when the apply-time selection
 *   no longer matches the staged text (nothing is written); other routes
 *   resolve undefined.
 */
export async function applySelectionAmendment(deps, proposal) {
    const { appState, log } = deps;
    const { selectionText, amendedText, commentText } = proposal;

    // Table route: per-cell tracked revisions plus row-level structure ops.
    // Table proposals never carry a comment (merged mode is skipped).
    if (proposal.tablePatch) {
        return _applyTablePatch(deps, proposal);
    }

    if (amendedText) {
        if (proposal.mixedTable) {
            // Mixed paragraph+table selection: guarded paragraph alignment.
            // The flat strategies' whole-selection fallback would destroy the
            // table, so this path has NO such fallback by design.
            log('Applying changes (paragraph-granular, table-guarded)...', 'info');
            const applied = await _applyMixedTableAmendment(deps, proposal);
            log(applied ? 'Changes applied successfully' : 'No differences found — nothing applied',
                applied ? 'success' : 'info');
        } else {
        log('Applying changes...', 'info');
        const trackChanges = !!appState.config.trackChangesEnabled;
        // Staleness guard: the staged amendment is anchored to the selection
        // text captured at prepare time. If the selection moved on between
        // staging and Apply, granular diffing would throw and the whole-
        // selection fallback below would overwrite whatever is selected NOW.
        // Refuse instead of writing to the wrong range.
        let stale = false;
        await Word.run(async (context) => {
            const selection = context.document.getSelection();
            selection.load('text');
            await context.sync();
            if (_normalizeSelectionText(selection.text) !== _normalizeSelectionText(selectionText)) {
                stale = true;
                return;
            }
            // Baseline mode per config; the strategies manage tracking from
            // here (and always restore off when they enabled it).
            if (Word.ChangeTrackingMode) {
                context.document.changeTrackingMode = trackChanges
                    ? Word.ChangeTrackingMode.trackAll
                    : Word.ChangeTrackingMode.off;
            }
            try {
                const strategyOptions = { trackChanges };
                if (appState.config.lineDiffEnabled) {
                    await applySentenceDiffStrategy(context, selection, selectionText, amendedText, log, strategyOptions);
                } else if (hasCjk(selectionText) || hasCjk(amendedText)) {
                    // CJK text has no word boundaries for the token map — use
                    // char-level diff so e.g. a one-comma edit stays minimal.
                    await applyCharDiffStrategy(context, selection, selectionText, amendedText, log, strategyOptions);
                } else {
                    await applyTokenMapStrategy(context, selection, selectionText, amendedText, log, strategyOptions);
                }
            } catch (diffErr) {
                // Last resort: replace the whole selection text. Loses edit
                // granularity but never leaves a failed apply. A failed
                // strategy may have left tracking in any state, so re-assert
                // the config mode first — an untracked write here would
                // silently bypass the redline.
                log(`Granular diff failed (${diffErr.message}), using whole-selection replacement`, 'warning');
                if (Word.ChangeTrackingMode) {
                    context.document.changeTrackingMode = trackChanges
                        ? Word.ChangeTrackingMode.trackAll
                        : Word.ChangeTrackingMode.off;
                }
                selection.insertText(amendedText, Word.InsertLocation.replace);
                await context.sync();
            } finally {
                // Comments and later turns must not inherit tracking state.
                if (Word.ChangeTrackingMode) {
                    context.document.changeTrackingMode = Word.ChangeTrackingMode.off;
                    await context.sync();
                }
            }
        });
        if (stale) {
            const reason = 'Selection changed since this proposal was staged; ' +
                'apply was skipped to avoid overwriting the wrong text. ' +
                'Re-run the instruction on the new selection.';
            log(reason, 'warning');
            return { skipped: true, reason };
        }
        log('Changes applied successfully', 'success');
        }
    }

    if (commentText && appState.supportsComments) {
        log('Inserting comment...', 'info');
        try {
            await Word.run(async (context) => {
                const selection = context.document.getSelection();
                selection.load('text');
                await context.sync();
                const contentRange = selection.getRange();
                contentRange.insertComment(commentText);
                await context.sync();
            });
            log('Comment inserted successfully', 'success');
        } catch (commentError) {
            // Comment insertion failed -- log the comment text so it is not lost
            log(`Comment insertion failed: ${commentError.message}. Comment text: "${commentText}"`, 'warning');
        }
    } else if (commentText && !appState.supportsComments) {
        log(`Comment generated but Word API 1.4 not available. Comment: "${commentText}"`, 'warning');
    }
}

/**
 * Normalizes selection text for staleness comparison: line endings and
 * whitespace runs are folded so a re-read selection only fails the guard
 * when the CONTENT actually moved.
 *
 * @param {string} text
 * @returns {string}
 * @private
 */
function _normalizeSelectionText(text) {
    return String(text || '').replace(/\r\n?/g, '\n').replace(/\s+/g, ' ').trim();
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
 * @private
 */
async function _applyTablePatch(deps, proposal) {
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
        await context.sync();
        if (range.text.trim() === cellPatch.text.trim()) return false;
        try {
            // The outer scope owns the tracking mode for the whole patch.
            const diffOptions = { trackChanges: false };
            if (hasCjk(range.text) || hasCjk(cellPatch.text)) {
                await applyCharDiffStrategy(context, range, range.text, cellPatch.text, log, diffOptions);
            } else {
                await applyTokenMapStrategy(context, range, range.text, cellPatch.text, log, diffOptions);
            }
        } catch (diffErr) {
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
    await context.sync();
    if (whole.text.replace(/\r/g, '\n').trim() === cellPatch.text.trim()) return false;

    const newLines = cellPatch.text.split(/\r?\n/);
    if (newLines.length === items.length) {
        for (let i = 0; i < items.length; i++) {
            const paraRange = items[i].getRange(Word.RangeLocation.content);
            paraRange.load('text');
            await context.sync();
            if (paraRange.text !== newLines[i]) {
                paraRange.insertText(newLines[i], Word.InsertLocation.replace);
            }
        }
        await context.sync();
        return true;
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
 * @private
 */
async function _applyMixedTableAmendment(deps, proposal) {
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
        allParaItems.forEach((p, i) => {
            if (p.text && p.text.trim() !== '') {
                paraItems.push(p);
                inTable.push(!tableChecks[i].isNullObject);
            }
        });
        if (paraItems.length === 0) {
            log('Mixed apply: selection contains only blank paragraphs, nothing to amend');
            return;
        }

        const origTexts = paraItems.map((p) => p.text);
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

        const alignment = _alignParagraphs(origTexts, amendedLines);

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
                        const diffOptions = { trackChanges: false };
                        if (hasCjk(paraRange.text) || hasCjk(newText)) {
                            await applyCharDiffStrategy(context, paraRange, paraRange.text, newText.trim(), log, diffOptions);
                        } else {
                            await applyTokenMapStrategy(context, paraRange, paraRange.text, newText.trim(), log, diffOptions);
                        }
                    } catch (diffErr) {
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
 * Generates new content to append at the document end — the prepare half of
 * the staged append flow. The document text is sent as context; the model is
 * instructed to return ONLY the text to append. The result is staged in a
 * proposal card and written by applyDocumentAppend.
 *
 * @param {object} deps - { appState, log }
 * @param {object} args
 * @param {string} args.instruction - The user's generation instruction
 * @param {string} [args.selectionText] - Current selection, sent as focused context
 * @param {function} [args.onToken] - Called with each streamed content token
 * @param {function} [args.onReasoning] - Called with each streamed thinking token
 * @returns {Promise<{ instruction: string, generatedText: string, model: string }>}
 */
export async function prepareDocumentAppend(deps, { instruction, selectionText, onToken, onReasoning, signal } = {}) {
    const { appState, log } = deps;

    const richness = (appState.config.docExtraction || {}).richness || 'structured';
    log('Extracting document text for context...', 'info');
    const documentText = await extractDocumentStructured({ richness });

    let prompt =
        'You are a writing assistant embedded in a Word document. The user wants NEW content generated and appended to the END of the document.\n' +
        'Continue the document in its own language, style, and tone, following the user instruction below.\n\n' +
        'CRITICAL OUTPUT RULES:\n' +
        '- Output ONLY the new content to append, as PLAIN TEXT.\n' +
        '- Do NOT repeat, quote, or paraphrase the existing document text.\n' +
        '- Do NOT use markdown formatting (no bold, italics, asterisks, headings, bullet points).\n' +
        '- Do NOT include explanations, commentary, introductory phrases, or concluding remarks.\n' +
        '- Separate paragraphs with a single blank line.\n\n' +
        'USER INSTRUCTION:\n' + (instruction || '').trim() + '\n';
    if (selectionText && selectionText.trim()) {
        prompt += '\n--- SELECTED TEXT (excerpt the user selected; treat it as the immediate context) ---\n' + selectionText.trim() + '\n';
    }
    prompt += '\n--- EXISTING DOCUMENT (context only — do not repeat it; continue after it) ---\n' + documentText;

    const backendConfig = getActiveBackendConfig(appState);
    log(`Drafting content to append [${backendConfig.model}]...`, 'info');
    const rawResponse = (onToken || onReasoning)
        ? await sendPromptStream(backendConfig, prompt, { onContent: onToken, onReasoning }, log, signal)
        : await sendPrompt(backendConfig, prompt, log, signal);
    log(`LLM response received [${backendConfig.model}]`, 'success');

    return {
        instruction: (instruction || '').trim(),
        generatedText: (stripMarkdown(rawResponse, log) || '').trim(),
        model: backendConfig.model,
    };
}

/**
 * Applies a prepared append proposal: inserts the generated text at the
 * document end, one Word paragraph per line, as tracked changes
 * (per config.trackChangesEnabled).
 *
 * @param {object} deps - { appState, log }
 * @param {object} proposal - Result of prepareDocumentAppend
 * @returns {Promise<{ paragraphsAppended: number, chars: number }>}
 */
export async function applyDocumentAppend(deps, proposal) {
    const { appState, log } = deps;
    const text = ((proposal && proposal.generatedText) || '').trim();
    if (!text) {
        throw new Error('Nothing to append — the model returned empty text.');
    }
    // The model is instructed to blank-line-separate paragraphs; treating
    // every newline as a paragraph break also covers single-\n output.
    const paragraphs = text.split(/\n+/).map((p) => p.trim()).filter(Boolean);

    await Word.run(async (context) => {
        if (Word.ChangeTrackingMode) {
            context.document.changeTrackingMode = appState.config.trackChangesEnabled
                ? Word.ChangeTrackingMode.trackAll
                : Word.ChangeTrackingMode.off;
        }
        const body = context.document.body;
        for (const paragraph of paragraphs) {
            body.insertParagraph(paragraph, Word.InsertLocation.end);
        }
        await context.sync();
        if (Word.ChangeTrackingMode) {
            context.document.changeTrackingMode = Word.ChangeTrackingMode.off;
            await context.sync();
        }
    });
    log(`Appended ${paragraphs.length} paragraph(s) to the document end.`, 'success');
    return { paragraphsAppended: paragraphs.length, chars: text.length };
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
    const rawResponse = (onToken || onReasoning)
        ? await sendPromptStream(backendConfig, prompt, { onContent: onToken, onReasoning }, log, signal)
        : await sendPrompt(backendConfig, prompt, log, signal);

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

/**
 * Collects the body paragraphs that are safe to delete as redundant empty
 * paragraphs: whitespace-only text, not the body's final paragraph (Word
 * requires a trailing paragraph mark), not inside a table cell, and carrying
 * no inline pictures (an image lives in an otherwise "empty" paragraph —
 * deleting it would destroy the image).
 *
 * @param {Word.RequestContext} context - An open Word.run context
 * @returns {Promise<{ paragraphs: Word.ParagraphCollection, indexes: number[] }>}
 *   Indexes into the collection's items.
 */
async function _collectEmptyParagraphs(context) {
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load('items');
    await context.sync();

    const items = paragraphs.items;
    for (const para of items) {
        para.load('text');
        const tableCheck = para.parentTableOrNullObject;
        tableCheck.load('isNullObject');
        const pictures = para.inlinePictures;
        pictures.load('items');
    }
    await context.sync();

    const lastIndex = items.length - 1;
    const indexes = [];
    items.forEach((para, index) => {
        const isEmpty = (para.text || '').trim() === '';
        const isLast = index === lastIndex;
        const inTable = para.parentTableOrNullObject && !para.parentTableOrNullObject.isNullObject;
        const hasPicture = para.inlinePictures && para.inlinePictures.items.length > 0;
        if (isEmpty && !isLast && !inTable && !hasPicture) {
            indexes.push(index);
        }
    });
    return { paragraphs, indexes };
}

/**
 * Scans the document for deletable empty paragraphs — the prepare half of the
 * staged empty-paragraph cleanup flow. No LLM is involved: blank paragraphs
 * are invisible to the document parser and excluded from text-pipeline
 * alignment, so only a deterministic Word.js scan can serve this instruction.
 * The result is staged in a proposal card and deleted by
 * applyEmptyParagraphCleanup.
 *
 * @param {object} deps - { appState, log }
 * @returns {Promise<{ emptyCount: number }>}
 */
export async function prepareEmptyParagraphCleanup(deps) {
    const { log } = deps;
    log('Scanning for empty paragraphs...', 'info');
    const { indexes } = await Word.run((context) => _collectEmptyParagraphs(context));
    log(`Found ${indexes.length} empty paragraph(s).`, 'info');
    return { emptyCount: indexes.length };
}

/**
 * Applies the staged empty-paragraph cleanup. Re-scans at apply time (the
 * document may have changed since staging — honest feedback reports the
 * actual deletions) and deletes the empty paragraphs as tracked changes per
 * config.trackChangesEnabled.
 *
 * @param {object} deps - { appState, log }
 * @returns {Promise<{ deleted: number }>}
 */
export async function applyEmptyParagraphCleanup(deps) {
    const { appState, log } = deps;
    let deleted = 0;
    await Word.run(async (context) => {
        const { paragraphs, indexes } = await _collectEmptyParagraphs(context);
        if (indexes.length === 0) {
            return;
        }
        if (Word.ChangeTrackingMode) {
            context.document.changeTrackingMode = appState.config.trackChangesEnabled
                ? Word.ChangeTrackingMode.trackAll
                : Word.ChangeTrackingMode.off;
        }
        try {
            // Delete in reverse document order via object references held
            // from the same context; index shifts can't affect later deletes.
            for (let i = indexes.length - 1; i >= 0; i--) {
                paragraphs.items[indexes[i]].delete();
            }
            deleted = indexes.length;
            await context.sync();
        } finally {
            if (Word.ChangeTrackingMode) {
                context.document.changeTrackingMode = Word.ChangeTrackingMode.off;
                await context.sync();
            }
        }
    });
    log(`Deleted ${deleted} empty paragraph(s).`, 'success');
    return { deleted };
}

/**
 * Decomposes a compound instruction into ordered pipeline tasks — the
 * planning step of a compound turn ("增加标题，并深度润色修改" →
 * [insert title, edit document]). The document text is NOT sent: the
 * planner only classifies intent, so this is a cheap call. Returns
 * tasks=null when the plan cannot be parsed; the caller then falls back to
 * single-intent routing.
 *
 * @param {object} deps - { appState, log }
 * @param {object} args
 * @param {string} args.instruction - The user's compound instruction
 * @param {boolean} [args.hasSelection] - Whether the document has a non-empty selection
 * @param {function} [args.onToken] - Called with each streamed content token
 * @param {function} [args.onReasoning] - Called with each streamed thinking token
 * @returns {Promise<{ tasks: Array<{ type: string, instruction: string }> | null, model: string }>}
 */
export async function planDocumentTasks(deps, { instruction, hasSelection = false, onToken, onReasoning, signal } = {}) {
    const { appState, log } = deps;

    const prompt = buildPlanPrompt(instruction, hasSelection);
    const backendConfig = getActiveBackendConfig(appState);
    log(`Planning tasks [${backendConfig.model}]...`, 'info');
    const rawResponse = (onToken || onReasoning)
        ? await sendPromptStream(backendConfig, prompt, { onContent: onToken, onReasoning }, log, signal)
        : await sendPrompt(backendConfig, prompt, log, signal);

    const tasks = parsePlan(rawResponse, log);
    if (tasks) log(`Planned ${tasks.length} task(s): ${tasks.map((t) => t.type).join(' → ')}`, 'success');
    return { tasks, model: backendConfig.model };
}

/**
 * Plans formatting changes from a natural-language instruction — the prepare
 * half of the staged format flow. The model returns a JSON op array (see
 * format-ops.js); the ops are validated there and staged in a proposal card,
 * written by applyFormatProposal only when the user applies.
 *
 * @param {object} deps - { appState, log }
 * @param {object} args
 * @param {string} args.instruction - The user's formatting instruction
 * @param {string} args.scope - 'selection' | 'document'
 * @param {string} [args.selectionText] - Current selection text (selection scope)
 * @param {function} [args.onToken] - Called with each streamed content token
 * @param {function} [args.onReasoning] - Called with each streamed thinking token
 * @returns {Promise<{ instruction: string, scope: string, ops: Array<object>, model: string }>}
 */
export async function prepareFormatProposal(deps, { instruction, scope = 'selection', selectionText, onToken, onReasoning, signal } = {}) {
    const { appState, log } = deps;

    let scopeText = '';
    if (scope === 'selection') {
        scopeText = (selectionText && selectionText.trim())
            || (await readSelectionText(deps)).selectionText;
    } else {
        const richness = (appState.config.docExtraction || {}).richness || 'structured';
        log('Extracting document text for context...', 'info');
        scopeText = await extractDocumentStructured({ richness });
    }

    const prompt = buildFormatPrompt(instruction, scopeText, scope);
    const backendConfig = getActiveBackendConfig(appState);
    log(`Planning formatting ops [${backendConfig.model}]...`, 'info');
    const rawResponse = (onToken || onReasoning)
        ? await sendPromptStream(backendConfig, prompt, { onContent: onToken, onReasoning }, log, signal)
        : await sendPrompt(backendConfig, prompt, log, signal);

    const ops = parseFormatOps(rawResponse, log);
    log(`Parsed ${ops.length} formatting op(s) from the model response.`, 'info');
    return { instruction, scope, ops, model: backendConfig.model };
}

/**
 * Applies a prepared format proposal. Insert ops add their paragraph(s) at
 * the scope start/end; other ops' targets are resolved inside the scope
 * range (whole scope, substring matches, or paragraphs of a given built-in
 * style), then font/paragraph properties are set with change tracking per
 * config (Word records them as Formatted revisions). List ops (listType/
 * listLevel) turn target paragraphs into a bulleted/numbered list or detach
 * them from one.
 *
 * @param {object} deps - { appState, log }
 * @param {object} proposal - Result of prepareFormatProposal
 * @returns {Promise<{ applied: boolean, appliedRanges: number, insertedParagraphs: number }>}
 */
export async function applyFormatProposal(deps, proposal) {
    const { appState, log } = deps;
    const { ops, scope } = proposal || {};
    if (!Array.isArray(ops) || ops.length === 0) {
        throw new Error('No formatting ops to apply.');
    }

    let appliedRanges = 0;
    let insertedParagraphs = 0;
    await Word.run(async (context) => {
        if (Word.ChangeTrackingMode) {
            context.document.changeTrackingMode = appState.config.trackChangesEnabled
                ? Word.ChangeTrackingMode.trackAll
                : Word.ChangeTrackingMode.off;
        }
        const scopeRange = scope === 'document'
            ? context.document.body
            : context.document.getSelection();

        let applied = 0;
        let inserted = 0;
        for (const op of ops) {
            if (op.insert) {
                inserted += await _applyInsertOp(context, scopeRange, op, log);
                continue;
            }
            const targets = await _resolveFormatTargets(context, scopeRange, op);
            if (targets.length === 0) {
                log(`Format op found no target (${op.match || op.paragraphStyle || 'scope'})`, 'warning');
                continue;
            }
            for (const target of targets) {
                if (op.font) _applyFontOps(target.font, op.font, log);
                if (op.paragraph) {
                    const paragraphs = target.paragraphs;
                    paragraphs.load('items');
                    await context.sync();
                    for (const paragraph of paragraphs.items) {
                        _applyParagraphOps(paragraph, op.paragraph, log);
                    }
                    if (_hasListOps(op.paragraph)) {
                        await _applyListOps(context, paragraphs.items, op.paragraph, log);
                    }
                }
                applied++;
            }
            await context.sync();
        }

        if (Word.ChangeTrackingMode) {
            context.document.changeTrackingMode = Word.ChangeTrackingMode.off;
            await context.sync();
        }
        appliedRanges = applied;
        insertedParagraphs = inserted;
        log(`Applied formatting to ${applied} range(s) and inserted ${inserted} paragraph(s) across ${ops.length} op(s).`, 'success');
    });
    return { applied: appliedRanges > 0 || insertedParagraphs > 0, appliedRanges, insertedParagraphs };
}

/**
 * Applies an insert op: queues the op's paragraph(s) at the start or end of
 * the scope range, styled by the op's font/paragraph payload. Paragraphs
 * inserted at the start are queued in reverse so their final order matches
 * the text. List ops need the paragraphs materialized first, so this syncs
 * internally when they are present. Returns the number of paragraphs
 * inserted.
 * @private
 */
async function _applyInsertOp(context, scopeRange, op, log) {
    const texts = op.insert.text.split(/\n+/).map((p) => p.trim()).filter(Boolean);
    const atStart = op.insert.position === 'start';
    const location = atStart ? Word.InsertLocation.start : Word.InsertLocation.end;
    const ordered = atStart ? [...texts].reverse() : texts;
    const inserted = [];
    for (const text of ordered) {
        const paragraph = scopeRange.insertParagraph(text, location);
        if (op.font) _applyFontOps(paragraph.font, op.font, log);
        if (op.paragraph) _applyParagraphOps(paragraph, op.paragraph, log);
        inserted.push(paragraph);
    }
    if (op.paragraph && _hasListOps(op.paragraph)) {
        await context.sync();
        // At-start inserts were queued last-to-first; restore document order.
        await _applyListOps(context, atStart ? [...inserted].reverse() : inserted, op.paragraph, log);
    }
    return texts.length;
}

/** True when a paragraph payload carries list ops (handled by _applyListOps). @private */
function _hasListOps(paragraphPayload) {
    return paragraphPayload.listType !== undefined || paragraphPayload.listLevel !== undefined;
}

/**
 * Resolves an op's target ranges: explicit substring matches, paragraphs of
 * a built-in style, or the whole scope when neither selector is given.
 * @private
 */
async function _resolveFormatTargets(context, scopeRange, op) {
    if (op.match) {
        // Word search strings are capped at 255 chars.
        const results = scopeRange.search(op.match.slice(0, 255), { matchCase: true, matchWholeWord: false });
        results.load('items');
        await context.sync();
        return results.items;
    }
    if (op.paragraphStyle) {
        const paragraphs = scopeRange.paragraphs;
        paragraphs.load('items/style');
        await context.sync();
        const enumValue = _enumValue(Word.Style, op.paragraphStyle);
        const wanted = String(op.paragraphStyle).toLowerCase();
        return paragraphs.items
            .filter((p) => (p.style || '').toLowerCase() === wanted
                || (enumValue && p.style === enumValue))
            .map((p) => p.getRange());
    }
    return [scopeRange];
}

/**
 * Case-insensitive Word enum lookup (e.g. 'heading1' -> Word.Style.heading1,
 * 'dark blue' -> Word.HighlightColor.darkBlue, 'GridTable4_Accent1' ->
 * Word.BuiltInStyleName.gridTable4_Accent1). Separators are stripped on BOTH
 * sides so snake/camel enum keys match their display spellings. Returns
 * undefined on miss.
 * @private
 */
function _enumValue(enumObj, name) {
    if (!enumObj || name === undefined || name === null) return undefined;
    const strip = (s) => String(s).toLowerCase().replace(/[\s_-]+/g, '');
    const wanted = strip(name);
    const key = Object.keys(enumObj).find((k) => strip(k) === wanted);
    return key ? enumObj[key] : undefined;
}

/**
 * Applies validated font ops to a Word.Font object. Unknown/invalid enum
 * values are skipped with a warning rather than failing the batch.
 * @private
 */
function _applyFontOps(font, ops, log) {
    for (const [key, value] of Object.entries(ops)) {
        try {
            if (key === 'underline') {
                const v = _enumValue(Word.UnderlineType, value);
                if (v === undefined) log(`Format ops: unknown underline "${value}"`, 'warning');
                else font.underline = v;
            } else if (key === 'highlightColor') {
                const v = _enumValue(Word.HighlightColor, value);
                if (v === undefined) log(`Format ops: unknown highlight color "${value}"`, 'warning');
                else font.highlightColor = v;
            } else {
                font[key] = value;
            }
        } catch (e) {
            log(`Format ops: font.${key} failed (${e.message})`, 'warning');
        }
    }
}

/**
 * Applies validated paragraph ops to a Word.Paragraph object. List keys are
 * skipped here — they need multi-paragraph coordination and live in
 * _applyListOps.
 * @private
 */
function _applyParagraphOps(paragraph, ops, log) {
    for (const [key, value] of Object.entries(ops)) {
        if (key === 'listType' || key === 'listLevel') continue;
        try {
            if (key === 'styleBuiltIn') {
                const v = _enumValue(Word.Style, value);
                if (v === undefined) log(`Format ops: unknown built-in style "${value}"`, 'warning');
                else paragraph.styleBuiltIn = v;
            } else if (key === 'alignment') {
                const v = _enumValue(Word.Alignment, value);
                if (v === undefined) log(`Format ops: unknown alignment "${value}"`, 'warning');
                else paragraph.alignment = v;
            } else {
                paragraph[key] = value;
            }
        } catch (e) {
            log(`Format ops: paragraph.${key} failed (${e.message})`, 'warning');
        }
    }
}

/**
 * Applies list ops to a batch of paragraphs (WordApi 1.3). 'bullet'/'number'
 * turn the non-list paragraphs into ONE list — the first starts it, the rest
 * attach — and format the requested level; 'none' detaches paragraphs from
 * their list. listLevel alone re-nests paragraphs already in a list. Syncs
 * internally; hosts without the list API get a warning instead of a failure.
 * @private
 */
async function _applyListOps(context, paragraphs, ops, log) {
    if (!paragraphs || paragraphs.length === 0) return;
    if (typeof Word.List === 'undefined' || typeof paragraphs[0].startNewList !== 'function') {
        log('Format ops: this Word host does not support list editing (needs WordApi 1.3); list ops skipped', 'warning');
        return;
    }
    const listType = ops.listType;
    const level = Number.isInteger(ops.listLevel) ? ops.listLevel : 0;

    if (listType === 'none') {
        for (const paragraph of paragraphs) {
            try {
                paragraph.detachFromList();
            } catch (e) {
                log(`Format ops: detachFromList failed (${e.message})`, 'warning');
            }
        }
        await context.sync();
        return;
    }

    // startNewList/attachToList fail at sync time for paragraphs that already
    // belong to a list, so learn membership up front.
    for (const paragraph of paragraphs) paragraph.load('isListItem');
    await context.sync();

    const fresh = [];
    for (const paragraph of paragraphs) {
        if (paragraph.isListItem) {
            if (ops.listLevel !== undefined) paragraph.listItem.level = level;
            if (listType) log('Format ops: a paragraph is already in a list; its list type was left unchanged', 'warning');
        } else {
            fresh.push(paragraph);
        }
    }
    if (!listType && fresh.length > 0) {
        log('Format ops: listLevel without listType only applies to existing list items', 'warning');
    }

    if (listType && fresh.length > 0) {
        const list = fresh[0].startNewList();
        list.load('id');
        await context.sync();
        if (listType === 'bullet') {
            const bullet = _enumValue(Word.ListBullet, 'solid');
            if (bullet !== undefined) list.setLevelBullet(level, bullet);
        } else {
            const numbering = _enumValue(Word.ListNumbering, 'arabic');
            if (numbering !== undefined) list.setLevelNumbering(level, numbering);
        }
        if (ops.listLevel !== undefined) fresh[0].listItem.level = level;
        for (const paragraph of fresh.slice(1)) paragraph.attachToList(list.id, level);
        await context.sync();
    } else if (ops.listLevel !== undefined) {
        await context.sync();
    }
}

/** Maximum width of an inserted illustration, in points (A4/Letter content width is ~450-470pt). */
const MAX_ILLUSTRATION_WIDTH_PT = 450;

/**
 * Designs an illustration for the document — the prepare half of the staged
 * illustration flow. The model returns ONE self-contained SVG (see
 * illustration.js); it is parsed, sanitized (safe for DOM preview), and
 * staged in a proposal card, written by applyIllustrationProposal only when
 * the user applies.
 *
 * @param {object} deps - { appState, log }
 * @param {object} args
 * @param {string} args.instruction - The user's illustration instruction
 * @param {function} [args.onToken] - Called with each streamed content token
 * @param {function} [args.onReasoning] - Called with each streamed thinking token
 * @returns {Promise<{ instruction: string, svg: string|null, position: string,
 *   positionLabel: string, model: string }>}
 *   svg is null when the model returned nothing usable.
 */
export async function prepareIllustrationProposal(deps, { instruction, onToken, onReasoning, signal } = {}) {
    const { appState, log } = deps;

    const richness = (appState.config.docExtraction || {}).richness || 'structured';
    log('Extracting document text for context...', 'info');
    const documentText = await extractDocumentStructured({ richness });

    const prompt = buildIllustrationPrompt(instruction, documentText);
    const backendConfig = getActiveBackendConfig(appState);
    log(`Designing illustration [${backendConfig.model}]...`, 'info');
    const rawResponse = (onToken || onReasoning)
        ? await sendPromptStream(backendConfig, prompt, { onContent: onToken, onReasoning }, log, signal)
        : await sendPrompt(backendConfig, prompt, log, signal);

    const parsed = parseIllustration(rawResponse, log);
    const svg = parsed ? ensureSvgDimensions(sanitizeSvg(parsed.svg)) : null;
    if (svg) log(`Illustration SVG received (${(svg.length / 1024).toFixed(1)} KB).`, 'success');
    const position = illustrationPositionFromInstruction(instruction);
    return {
        instruction: (instruction || '').trim(),
        svg,
        position,
        positionLabel: illustrationPositionLabel(position),
        model: backendConfig.model,
    };
}

/**
 * Applies a prepared illustration proposal: rasterizes the sanitized SVG to
 * PNG (Word's insertInlinePictureFromBase64 takes PNG/JPEG/GIF/BMP base64,
 * not SVG) and inserts it, as a tracked change per config:
 * - start/end: a centered inline picture in its own paragraph at the
 *   document start or end;
 * - cursor: INLINE at the caret — the insertion point is read at apply
 *   time (the document keeps its selection while focus is in the taskpane),
 *   anchored to the selection end so a non-collapsed selection never has
 *   its text replaced.
 *
 * @param {object} deps - { appState, log }
 * @param {object} proposal - Result of prepareIllustrationProposal
 * @returns {Promise<{ inserted: boolean }>}
 */
export async function applyIllustrationProposal(deps, proposal) {
    const { appState, log } = deps;
    const svg = ((proposal && proposal.svg) || '').trim();
    if (!svg) {
        throw new Error('No illustration to apply — the model returned no usable SVG.');
    }
    const { base64, width, height } = await svgToPngBase64(svg);
    const position = proposal.position === 'start' || proposal.position === 'cursor'
        ? proposal.position
        : 'end';

    await Word.run(async (context) => {
        if (Word.ChangeTrackingMode) {
            context.document.changeTrackingMode = appState.config.trackChangesEnabled
                ? Word.ChangeTrackingMode.trackAll
                : Word.ChangeTrackingMode.off;
        }
        try {
            const picture = insertPngPicture(context, {
                base64,
                position,
            });
            await context.sync();
            // Scaling reads the synced width; alt text is best-effort.
            finalizeInsertedPicture(context, picture, (proposal.instruction || 'Illustration').slice(0, 200));
            await context.sync();
        } finally {
            if (Word.ChangeTrackingMode) {
                context.document.changeTrackingMode = Word.ChangeTrackingMode.off;
                await context.sync();
            }
        }
        log(`Inserted illustration at ${illustrationPositionLabel(position)} (${width}x${height}px PNG).`, 'success');
    });
    return { inserted: true };
}

/**
 * Inserts one PNG (base64) as an inline picture inside an OPEN Word.run
 * context — the shared insertion primitive of the illustration pipeline
 * and the image tool loop. The caller owns the change-tracking mode.
 *
 * start/end: a centered inline picture in its own paragraph at the document
 * start/end. cursor: INLINE at the caret — the insertion point is read when
 * this runs (the document keeps its selection while focus is in the
 * taskpane), anchored to the selection end so a non-collapsed selection
 * never has its text replaced. Oversized pictures scale down to the content
 * width; alt text is best-effort.
 *
 * @param {Word.RequestContext} context - An open Word.run context
 * @param {object} args
 * @param {string} args.base64 - PNG base64 (no data-URL prefix)
 * @param {'start'|'end'|'cursor'} args.position
 * @returns {Word.InlinePicture} The inserted picture proxy
 */
export function insertPngPicture(context, { base64, position }) {
    let picture;
    if (position === 'cursor') {
        // Zero-width range at the selection end: inserting at its
        // start puts the picture right after the selected text —
        // exactly at a collapsed caret, and never over a selection.
        const insertionPoint = context.document.getSelection()
            .getRange(Word.RangeLocation.end);
        picture = insertionPoint.insertInlinePictureFromBase64(base64, Word.InsertLocation.start);
    } else {
        const location = position === 'start' ? Word.InsertLocation.start : Word.InsertLocation.end;
        // A dedicated paragraph keeps the image standalone and centerable.
        const paragraph = context.document.body.insertParagraph('', location);
        picture = paragraph.getRange(Word.RangeLocation.start)
            .insertInlinePictureFromBase64(base64, Word.InsertLocation.start);
        paragraph.alignment = Word.Alignment.centered;
    }
    picture.load('width,height');
    return picture;
}

/**
 * Queues the post-insert fixes for a picture inserted by insertPngPicture:
 * scale oversized images down to the content width (keeping the aspect
 * ratio) and set best-effort alt text. Runs inside the caller's Word.run —
 * sync happens once after all queued ops.
 *
 * @param {Word.RequestContext} context
 * @param {Word.InlinePicture} picture
 * @param {string} [altText]
 */
export function finalizeInsertedPicture(context, picture, altText) {
    // Inline picture dimensions are points; scale oversized images down
    // to the content width, keeping the aspect ratio.
    if (picture.width > MAX_ILLUSTRATION_WIDTH_PT) {
        picture.height = picture.height * (MAX_ILLUSTRATION_WIDTH_PT / picture.width);
        picture.width = MAX_ILLUSTRATION_WIDTH_PT;
    }
    try {
        if (altText) picture.altTextDescription = String(altText).slice(0, 200);
    } catch (_e) {
        // Alt text is best-effort (not critical for the insertion).
    }
}

/**
 * Rasterizes an SVG string to PNG base64 at 1600px wide via an offscreen
 * canvas (browser/WebView2 only — not exercised under node tests). The SVG
 * is vector, so rendering up to 1600px keeps the inserted image crisp.
 * Exported for the image tool loop (agent-actions).
 */
export async function svgToPngBase64(svg) {
    const dims = svgDimensions(svg) || { width: 1200, height: 800 };
    const blob = new Blob([svg], { type: 'image/svg+xml;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    try {
        const img = new Image();
        await new Promise((resolve, reject) => {
            img.onload = resolve;
            img.onerror = () => reject(new Error('The SVG could not be rendered.'));
            img.src = url;
        });
        const canvas = document.createElement('canvas');
        canvas.width = 1600;
        canvas.height = Math.max(1, Math.round(dims.height * (1600 / dims.width)));
        canvas.getContext('2d').drawImage(img, 0, 0, canvas.width, canvas.height);
        return { base64: canvas.toDataURL('image/png').split(',')[1], width: canvas.width, height: canvas.height };
    } finally {
        URL.revokeObjectURL(url);
    }
}

/**
 * Fires the fire-and-forget comment pipeline on the current selection
 * (bookmark capture + async LLM call + comment insertion with retry link).
 *
 * @param {object} deps - { appState, log, logWithRetry, updateStatusBar }
 * @param {object} args
 * @param {string} args.promptTemplate - Explicit comment template
 */
export async function fireSelectionComment(deps, { promptTemplate } = {}) {
    const { appState, log, logWithRetry, updateStatusBar } = deps;
    const { selectionText } = await readSelectionText(deps);
    fireCommentRequest(selectionText, {
        config: getActiveBackendConfig(appState),
        sendPromptFn: sendPrompt,
        promptManager: makePromptShim(appState.promptManager, 'comment', promptTemplate),
        commentQueue: appState.commentQueue,
        log,
        addLogWithRetryFn: logWithRetry,
        updateStatusBarFn: updateStatusBar,
    });
}

/**
 * Runs the whole-document chunked pipeline (parse → chunk → context →
 * parallel LLM → apply as tracked changes/comments). Preserves the old
 * handleProcessDocument semantics; cancellation via appState.processDocController.
 *
 * @param {object} deps - { appState, log, logWithRetry }
 * @param {object} args
 * @param {string} args.category - 'amendment' | 'comment'
 * @param {string} args.promptTemplate - Explicit template for the pipeline
 * @param {string} [args.commentInstructions] - Merged-mode comment instructions
 * @param {function} [args.onProgress] - Progress callback from the orchestrator
 * @param {function} [args.onChunkToken] - Live per-chunk token callback
 *   (chunkInfo, kind, token); kind is 'content' or 'reasoning'
 * @param {boolean} [args.gateApply=false] - When true, stop after the LLM
 *   processing phase and return an apply/discard continuation instead of
 *   writing to the document (used to stage a proposal card for amendments).
 * @returns {Promise<{ results: Array, applicationResult?: object, chunks: Array, cancelled?: boolean,
 *   staged?: boolean, apply?: Function, discard?: Function, failedCount?: number, cancelledCount?: number }>}
 */
export async function runDocumentSkill(deps, { category, promptTemplate, commentInstructions = '', onProgress, onChunkToken, gateApply = false, signal: signalArg } = {}) {
    const { appState, log, logWithRetry } = deps;
    // Honor the caller's signal; fall back to the shared processing slot for
    // legacy callers that rely on it having been set before invocation.
    const signal = signalArg || (appState.processDocController ? appState.processDocController.signal : undefined);
    const promptShim = makePromptShim(appState.promptManager, category, promptTemplate);

    // Step 1: Parse document
    log('Parsing document...', 'info');
    const docModel = await parseDocument();
    log(`Found ${docModel.paragraphs.length} paragraphs (~${docModel.totalTokens} tokens)`, 'info');

    // Step 2: Chunk document
    const chunks = chunkDocument(docModel, { maxTokens: 6000 });
    log(`Split into ${chunks.length} chunks`, 'info');

    // Step 3: Extract context
    const documentContext = extractContext(docModel);
    log(`Extracted ${documentContext.definitions.length} definitions, ${documentContext.outline.length} headings`, 'info');

    // Step 4: Bookmark chunk ranges
    const bookmarkMap = await bookmarkChunkRanges(chunks);

    // Step 5: Process chunks in parallel
    const backendConfig = getActiveBackendConfig(appState);
    const concurrency = chunks.some(c => c.tokenCount > 8000) ? 4 : 6;

    const results = await processChunksParallel(chunks, {
        config: backendConfig,
        promptManager: promptShim,
        documentContext: documentContext,
        log,
        onProgress,
        onChunkToken,
        signal,
        concurrency,
        timeoutMs: 300000,
        commentInstructions,
    });

    const failed = results.filter(r => r.status === 'rejected').length;
    const cancelled = results.filter(r => r.status === 'cancelled').length;

    /**
     * Applies the processed results to the document (tracked changes +
     * comments), cleans up bookmarks, logs the summary, and registers the
     * retry link for failed chunks.
     * @param {Array<string>} [chunkIds] - Selective apply from the proposal
     *   card's change list: only amendment results for these chunks are
     *   applied. Results without an amendment (comment-only) are not
     *   selectable items, so they always ride along.
     * @param {object} [opts]
     * @param {AbortSignal} [opts.signal] - Cooperative pause signal (Stop
     *   button): apply stops at the next chunk boundary, leaving remaining
     *   chunk bookmarks in place so the card can offer "Continue applying".
     * @param {function(string, object)} [opts.onChunkApplied] - Per-chunk
     *   progress callback for the proposal card to mark items applied.
     * @returns {Promise<object>} applicationResult from applyChunkResults
     */
    const apply = async (chunkIds, { signal, onChunkApplied } = {}) => {
        const selected = Array.isArray(chunkIds)
            ? results.filter((r) => !r.amendment || (r.chunk && chunkIds.includes(r.chunk.id)))
            : results;
        log(`Applying changes to document${Array.isArray(chunkIds) ? ` (${selected.length} of ${results.length} section(s))` : ''}...`, 'info');
        const applicationResult = await applyChunkResults(selected, bookmarkMap, {
            trackChangesEnabled: appState.config.trackChangesEnabled,
            lineDiffEnabled: appState.config.lineDiffEnabled,
            log,
            ...(signal ? { signal } : {}),
            ...(onChunkApplied ? { onChunkApplied } : {}),
        });

        // On a pause (interrupted) the not-yet-applied chunks must keep their
        // bookmarks so the "Continue applying" resume can still target them;
        // only clean up once the whole selection has been applied. Failed
        // chunks also keep theirs: the retry link below re-drives exactly
        // those chunks and needs the staged ranges to still exist.
        if (!applicationResult.interrupted) {
            const keepNames = new Set(
                results
                    .filter((r) => r.status === 'rejected')
                    .map((r) => bookmarkMap.get(r.chunkId))
                    .filter(Boolean)
            );
            await cleanupBookmarks(bookmarkMap, { keep: keepNames });
        }

        log(
            `Document processed: ${chunks.length} chunks, ` +
            `${applicationResult.amendmentsApplied} amendments applied, ` +
            `${applicationResult.commentsInserted} comments inserted` +
            (failed > 0 ? `, ${failed} chunks failed` : '') +
            (cancelled > 0 ? `, ${cancelled} chunks cancelled` : ''),
            failed > 0 ? 'warning' : 'success'
        );

        if (failed > 0 && logWithRetry) {
            logWithRetry(
                `${failed} chunk(s) failed. Click to retry failed chunks.`,
                'warning',
                retryFailed
            );
        }

        return applicationResult;
    };

    /**
     * Retries this run's failed chunks with the original bookmark map,
     * backend config, and prompt shim. Exposed on the gated return too, so
     * an all-chunks-failed run (where apply() never runs and no card is
     * staged) can still offer a retry link.
     */
    const retryFailed = () => {
        const failedChunks = results.filter(r => r.status === 'rejected');
        if (failedChunks.length === 0) {
            return Promise.resolve();
        }
        return retryFailedChunks(deps, { failedResults: failedChunks, bookmarkMap, backendConfig, promptShim, onProgress });
    };

    /**
     * Discards a staged run: removes the hidden chunk bookmarks without
     * touching document text.
     */
    const discard = async () => {
        await cleanupBookmarks(bookmarkMap);
        log('Proposed changes discarded; no edits were applied.', 'info');
    };

    // Gated mode (amendment pipeline): stop before writing to the document
    // and hand the caller an apply/discard continuation so the chat UI can
    // stage a proposal card for user confirmation. retryFailed rides along:
    // when every chunk failed there is nothing to stage, and this is the
    // only handle the caller has to offer a retry link.
    if (gateApply) {
        return { staged: true, results, chunks, apply, discard, retryFailed, failedCount: failed, cancelledCount: cancelled };
    }

    const applicationResult = await apply();
    return { results, applicationResult, chunks, cancelled: cancelled > 0 };
}

/**
 * Retries processing only the failed chunks of a document-scope run.
 * Re-runs the orchestrator on the failed chunk subset and applies results.
 *
 * @param {object} deps - { appState, log }
 * @param {object} args
 * @param {Array} args.failedResults - ChunkResult objects with status 'rejected'
 * @param {Map} args.bookmarkMap - Original chunkId -> bookmarkName map
 * @param {object} args.backendConfig - Backend configuration
 * @param {object} args.promptShim - PromptManager shim from the original run
 * @param {function} [args.onProgress] - Progress callback
 */
export async function retryFailedChunks(deps, { failedResults, bookmarkMap, backendConfig, promptShim, onProgress } = {}) {
    const { appState, log } = deps;

    if (appState.isProcessingDoc) {
        log('Document processing is already running. Wait for it to finish before retrying.', 'warning');
        return;
    }

    log(`Retrying ${failedResults.length} failed chunk(s)...`, 'info');
    appState.isProcessingDoc = true;
    appState.processDocController = new AbortController();

    try {
        // Re-drive the ORIGINAL chunk objects. ChunkResult.chunk carries the
        // full DocumentChunk (paragraphs, overlap, id), which is what the
        // orchestrator's composer and the reassembler's alignment both need;
        // rebuilding text-only stubs here used to crash every retry chunk
        // inside chunk.paragraphs.map and re-reject it instantly.
        const retryChunks = failedResults.map((r) => r.chunk).filter(Boolean);

        const results = await processChunksParallel(retryChunks, {
            config: backendConfig,
            promptManager: promptShim,
            // No context prefix on retries: rebuilding the full document
            // context is not worth a re-extraction, and orchestrator treats
            // a null context as "no prefix" rather than crashing.
            documentContext: null,
            log,
            onProgress,
            signal: appState.processDocController.signal,
            concurrency: 4,
            timeoutMs: 300000,
            commentInstructions: '',
        });

        // Retry chunks are rebuilt without paragraphs, so carry the staged
        // original texts forward explicitly: applyChunkResults uses them to
        // re-anchor bookmark ranges that drifted since staging.
        const chunkOriginals = new Map();
        for (const r of failedResults) {
            if (r.chunk && Array.isArray(r.chunk.paragraphs) && r.chunk.paragraphs.length > 0) {
                chunkOriginals.set(r.chunkId, r.chunk.paragraphs.map((p) => p.text));
            }
        }

        const applicationResult = await applyChunkResults(results, bookmarkMap, {
            trackChangesEnabled: appState.config.trackChangesEnabled,
            lineDiffEnabled: appState.config.lineDiffEnabled,
            log,
            chunkOriginals,
        });

        const stillFailed = results.filter(r => r.status === 'rejected').length;
        log(
            `Retry complete: ${applicationResult.amendmentsApplied} amendments, ` +
            `${applicationResult.commentsInserted} comments` +
            (stillFailed > 0 ? `, ${stillFailed} still failed` : ''),
            stillFailed > 0 ? 'warning' : 'success'
        );

        // The retry is the failed chunks' second chance: clean up their
        // bookmarks either way so nothing lingers in the document. (The
        // original apply() kept exactly these bookmarks alive for this call.)
        const retriedBookmarks = new Map(
            failedResults
                .map((r) => [r.chunkId, bookmarkMap.get(r.chunkId)])
                .filter(([, name]) => name)
        );
        if (retriedBookmarks.size > 0) {
            await cleanupBookmarks(retriedBookmarks);
        }
    } catch (error) {
        if (error.name === 'AbortError') {
            log('Retry cancelled.', 'warning');
        } else {
            log(`Retry failed: ${error.message}`, 'error');
        }
    } finally {
        appState.isProcessingDoc = false;
        appState.processDocController = null;
    }
}

/**
 * Runs the summary pipeline: extract comments (+ optional document text and
 * tracked changes per template placeholders), send to LLM, open a new Word
 * document with the formatted summary.
 *
 * @param {object} deps - { appState, log }
 * @param {object} args
 * @param {string} args.promptTemplate - Summary template; supports {comments},
 *   {whole document}, {tracked changes} placeholders
 * @param {function} [args.onToken] - Called with each streamed content token
 * @param {function} [args.onReasoning] - Called with each streamed thinking token
 */
export async function runSummarySkill(deps, { promptTemplate, onToken, onReasoning, signal } = {}) {
    const { appState, log } = deps;

    log('Extracting document comments...', 'info');
    const comments = await extractAllComments();
    if (comments.length === 0) {
        log('No review comments found — summarizing from document text only.', 'info');
    }

    const summaryOpts = { templateOverride: promptTemplate };

    if (promptTemplate.includes('{whole document}')) {
        const richness = (appState.config.docExtraction || {}).richness || 'structured';
        log(`Extracting document text (${richness})...`, 'info');
        summaryOpts.documentText = await extractDocumentStructured({ richness });
        log(`Document text extracted (${summaryOpts.documentText.length} chars, ~${estimateTokenCount(summaryOpts.documentText)} tokens)`, 'info');
    }

    if (appState.config.trackedChangesExtraction && promptTemplate.includes('{tracked changes}')) {
        log('Extracting tracked changes (OOXML parsing)...', 'info');
        const tcResult = await extractTrackedChanges();
        log(`Tracked changes extracted (${tcResult.changes.length} change(s))`, 'info');

        let tcText = '';
        if (tcResult.changes.length > 0) {
            tcText = tcResult.changes.map((c, i) => {
                const num = i + 1;
                const author = c.author || 'Unknown';
                const date = c.date || '';
                const dateStr = date ? ` on ${date}` : '';

                if (c.type === 'Replaced') {
                    return `[Change ${num}] REPLACED by ${author}${dateStr}:\n` +
                           `  BEFORE: "${c.beforeText}"\n` +
                           `  AFTER:  "${c.afterText}"` +
                           (c.paragraphText ? `\n  IN CLAUSE: "${c.paragraphText}"` : '');
                } else if (c.type === 'Deleted') {
                    return `[Change ${num}] DELETED by ${author}${dateStr}:\n` +
                           `  REMOVED: "${c.text}"` +
                           (c.paragraphText ? `\n  IN CLAUSE: "${c.paragraphText}"` : '');
                } else if (c.type === 'Added') {
                    return `[Change ${num}] ADDED by ${author}${dateStr}:\n` +
                           `  INSERTED: "${c.text}"` +
                           (c.paragraphText ? `\n  IN CLAUSE: "${c.paragraphText}"` : '');
                } else if (c.type.startsWith('Moved')) {
                    return `[Change ${num}] ${c.type.toUpperCase()} by ${author}${dateStr}:\n` +
                           `  TEXT: "${c.text}"` +
                           (c.paragraphText ? `\n  IN CLAUSE: "${c.paragraphText}"` : '');
                }
                return `[Change ${num}] ${c.type} by ${author}${dateStr}: "${c.text}"`;
            }).join('\n\n');
        }

        if (tcText) {
            summaryOpts.trackedChangesText = tcText;
        } else if (tcResult.changes.length === 0) {
            summaryOpts.trackedChangesText = '(No tracked changes found in document)';
        }
    }

    const messages = appState.promptManager.composeSummaryMessages(comments, summaryOpts);
    if (messages.length === 0) {
        throw new Error('No summary prompt composed — check the skill template');
    }

    const backendConfig = getActiveBackendConfig(appState);
    log(`Sending summary request [${backendConfig.model}]...`, 'info');
    const promptText = _flattenMessages(messages);
    const llmResponse = (onToken || onReasoning)
        ? await sendPromptStream(backendConfig, promptText, { onContent: onToken, onReasoning }, log, signal)
        : await sendPrompt(backendConfig, promptText, log, signal);
    log(`Summary received (${llmResponse.length} chars). Creating document...`, 'info');

    let docTitle = 'Document Summary';
    try {
        await Word.run(async (context) => {
            const props = context.document.properties;
            props.load('title');
            await context.sync();
            if (props.title) {
                docTitle = `Summary - ${props.title}`;
            }
        });
    } catch (_titleErr) {
        // Title lookup failed -- use default
    }

    const html = buildSummaryHtml(llmResponse, comments, docTitle);
    await createSummaryDocument(html, docTitle, log);
    log('Summary document opened successfully.', 'success');

    return { chars: llmResponse.length, commentCount: comments.length };
}

/**
 * Answers a free-text question in chat using the document as context.
 * Streams tokens via sendPromptStream when the backend supports SSE.
 *
 * @param {object} deps - { appState, log }
 * @param {object} args
 * @param {string} args.question - The user's question
 * @param {string} [args.skillTemplate] - Persona/instruction template from a chat skill
 * @param {string} [args.selectionText] - Currently selected text, added as a
 *   focused excerpt before the full document context; when EMPTY, the cursor
 *   location (caret paragraph + nearest heading) is injected instead
 * @param {Array<{width: number, height: number, altText: string}>} [args.selectionImages] -
 *   Metadata of inline pictures inside the selection. They enter the prompt
 *   as object references (size + alt text), not raw bytes — the model is
 *   pointed at the image tool session (read_image) for visual content
 * @param {function} [args.onToken] - Called with each streamed token
 * @param {function} [args.onReasoning] - Called with each streamed thinking token
 * @param {function} [args.onStatus] - Called with stage updates ("Reading the document...", "Waiting for model...")
 * @param {AbortSignal} [args.signal] - Cancellation signal
 * @returns {Promise<string>} The full answer text
 */
export async function answerQuestion(deps, { question, skillTemplate, selectionText, selectionImages, onToken, onReasoning, onStatus, signal } = {}) {
    const { appState, log } = deps;

    const richness = (appState.config.docExtraction || {}).richness || 'structured';
    log('Extracting document text for context...', 'info');
    if (onStatus) onStatus('Reading the document...');
    const documentText = await extractDocumentStructured({ richness });

    let prompt = '';
    const contextPrompt = appState.promptManager.getActivePrompt('context');
    if (contextPrompt) {
        prompt += contextPrompt.template + '\n\n';
    }
    if (skillTemplate) {
        prompt += skillTemplate + '\n\n';
    }
    prompt += question;
    // Image metadata for MIXED text+image selections: object references in
    // the prompt, never the bytes (visual reading belongs to the image tool
    // session's read_image). Pure image selections never reach QA — they
    // route to the image tool turn.
    const imageMeta = (Array.isArray(selectionImages) ? selectionImages : [])
        .slice(0, MAX_SELECTION_IMAGES);
    if (selectionText && selectionText.trim()) {
        // Table selections carry structure that selection.text flattens away:
        // read the grid as markdown, fall back to the flat text otherwise
        // (readSelectionTableContext swallows its own failures → null).
        const tableContext = await readSelectionTableContext(deps);
        if (tableContext) {
            const label = tableContext.kind === 'mixed'
                ? 'SELECTED CONTENT (paragraphs and tables the user is asking about)'
                : 'SELECTED TABLE (the user is asking about this table)';
            prompt += `\n\n--- ${label} ---\n` + tableContext.contextText;
        } else {
            prompt += '\n\n--- SELECTED TEXT (the user is asking about this excerpt) ---\n' + selectionText.trim();
        }
        if (imageMeta.length > 0) {
            prompt += '\n\n--- SELECTED IMAGES (also inside the selection) ---\n' +
                imageMeta.map((img, i) =>
                    `- image ${i + 1}: ${img.width}x${img.height}pt` +
                    (img.altText ? `, alt "${String(img.altText).slice(0, 80)}"` : '')).join('\n') +
                '\n(Visual content is not attached here; image instructions can view and edit these pictures.)';
        }
    } else if (imageMeta.length > 0) {
        // Image-only selection reaching QA (defense in depth — routing sends
        // image-only selections to the image tool session; planned qa tasks
        // may still land here): object references, not bytes.
        prompt += '\n\n--- SELECTED IMAGES (the user is asking about these pictures) ---\n' +
            imageMeta.map((img, i) =>
                `- image ${i + 1}: ${img.width}x${img.height}pt` +
                (img.altText ? `, alt "${String(img.altText).slice(0, 80)}"` : '')).join('\n') +
            '\n(Visual content is not attached here; image instructions can view and edit these pictures.)';
    } else {
        // Bare caret: inject where the user is working so document-scope
        // answers can weight the current section (no tool-call path exists;
        // all context is injected before the request).
        const cursorContext = await readCursorContext(deps);
        if (cursorContext) {
            prompt += '\n\n--- CURSOR LOCATION (where the user is in the document) ---\n' + cursorContext.contextText;
        }
    }
    prompt += '\n\n--- DOCUMENT ---\n' + documentText;

    const backendConfig = getActiveBackendConfig(appState);
    log(`Asking [${backendConfig.model}]...`, 'info');
    if (onStatus) onStatus(`Waiting for ${backendConfig.model}...`);
    // Long documents + slow backends can exceed the 120s client default;
    // chat answers get 5 minutes (the doc pipeline uses the same per-chunk).
    return sendPromptStream(backendConfig, prompt, { onContent: onToken, onReasoning }, log, signal, 300000);
}

/**
 * Scrolls the Word document to the first occurrence of a text snippet.
 * Used by citation pills on completed document-scope runs.
 *
 * @param {object} deps - { log }
 * @param {string} searchText - Distinctive text near the start of the chunk
 * @returns {Promise<boolean>} True when a match was found and selected
 */
export async function revealTextSnippet(deps, searchText) {
    const { log } = deps;
    const needle = (searchText || '').trim().slice(0, 200);
    if (!needle) return false;

    try {
        let found = false;
        await Word.run(async (context) => {
            const results = context.document.body.search(needle, { matchCase: false, matchWholeWord: false });
            results.load('items');
            await context.sync();
            if (results.items.length > 0) {
                results.items[0].select();
                await context.sync();
                found = true;
            }
        });
        if (!found) {
            log('Could not locate the chunk text in the document (it may have changed).', 'warning');
        }
        return found;
    } catch (error) {
        log(`Jump to section failed: ${error.message}`, 'warning');
        return false;
    }
}
