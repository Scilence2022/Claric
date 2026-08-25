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
import { supportsTrackedRowOps } from '../lib/platform.js';
import { fireCommentRequest } from '../lib/comment-request.js';
import { extractAllComments, extractDocumentStructured, estimateTokenCount, extractTrackedChanges, extractCommentsOnRange } from '../lib/comment-extractor.js';
import { formatSelectionWithComments } from '../lib/selection-with-comments.js';
import { createSummaryDocument, buildSummaryHtml } from '../lib/document-generator.js';
import { parseDelimitedResponse, buildFallbackClassificationPrompt } from '../lib/response-parser.js';
import { parseDocument } from '../lib/document-parser.js';
import { chunkDocument } from '../lib/document-chunker.js';
import { extractContext } from '../lib/context-extractor.js';
import { processChunksParallel } from '../lib/orchestrator.js';
import { bookmarkChunkRanges, applyChunkResults, cleanupBookmarks } from '../lib/reassembler.js';
import { buildFormatPrompt, parseFormatOps } from '../lib/format-ops.js';
import {
    buildIllustrationPrompt, parseIllustration, sanitizeSvg,
    ensureSvgDimensions, svgDimensions, illustrationPositionFromInstruction,
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
 * compares, no row-collection iteration).
 *
 * @param {object} deps - { appState, log }
 * @returns {Promise<null | {rowCount: number, colCount: number,
 *   cells: Array<{row: number, col: number, text: string}>}>}
 *   Row/col in cells are 1-based absolute table coordinates.
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
        startCell.load('isNullObject,rowIndex,columnIndex');
        endCell.load('isNullObject,rowIndex,columnIndex');
        table.load('rowCount,values');
        await context.sync();

        if (startCell.isNullObject || endCell.isNullObject) {
            throw new Error('Could not locate the selected cells — re-select the table region.');
        }

        const startRow = Math.min(startCell.rowIndex, endCell.rowIndex);
        const endRow = Math.max(startCell.rowIndex, endCell.rowIndex);
        const startCol = Math.min(startCell.columnIndex, endCell.columnIndex);
        const endCol = Math.max(startCell.columnIndex, endCell.columnIndex);

        const values = table.values || [];
        const colCount = values[0] ? values[0].length : 0;
        const cells = [];
        for (let r = startRow; r <= endRow; r++) {
            for (let c = startCol; c <= endCol; c++) {
                cells.push({ row: r + 1, col: c + 1, text: (values[r] && values[r][c]) || '' });
            }
        }
        region = { rowCount: table.rowCount, colCount, cells };
    });
    if (region) {
        const first = region.cells[0];
        const last = region.cells[region.cells.length - 1];
        log(`Table selection: ${region.cells.length} cell(s), R${first.row}C${first.col} → R${last.row}C${last.col}`, 'info');
    }
    return region;
}

/**
 * Returns whether the document currently has a non-collapsed selection.
 * Resolves to false when the Word API is unavailable.
 *
 * @returns {Promise<boolean>}
 */
export async function hasNonEmptySelection() {
    try {
        let has = false;
        await Word.run(async (context) => {
            const selection = context.document.getSelection();
            selection.load('text');
            await context.sync();
            has = !!(selection.text && selection.text.trim());
        });
        return has;
    } catch (_err) {
        return false;
    }
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
 * Watches the Word selection and invokes callback with the current selected
 * plain text ('' when nothing is selected). Events are debounced so a drag
 * selection fires one Word.run at the end.
 *
 * @param {function(string)} callback - Receives the selection text
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
            callback(await readSelectionSnippet());
        }, debounceMs);
    };

    Office.context.document.addHandlerAsync(
        Office.EventType.DocumentSelectionChanged,
        onChange,
        () => {} // registration failure is non-fatal: the preview just stays static
    );

    // Emit the initial state (the user may have selected text before opening the pane)
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
 * @returns {Promise<{ selectionText: string, amendedText: string|null, commentText: string|null, model: string }>}
 */
export async function prepareSelectionAmendment(deps, { promptTemplate, commentInstructions, onToken, onReasoning } = {}) {
    const { appState, log } = deps;
    const backendConfig = getActiveBackendConfig(appState);

    // Multi-cell table selections take the coordinate patch protocol — the
    // flat-text pipelines below cannot represent cell boundaries.
    const tableRegion = await readSelectionTableRegion(deps);
    if (tableRegion) {
        return _prepareTableAmendment(deps, {
            tableRegion, promptTemplate, commentInstructions, backendConfig, onToken, onReasoning,
        });
    }

    const { selectionText } = await readSelectionText(deps);
    log(`Processing selection (${selectionText.length} chars) via ${backendConfig.model}...`, 'info');

    const merged = !!(commentInstructions && commentInstructions.trim());
    const messages = merged
        ? appState.promptManager.composeMergedMessages(selectionText, commentInstructions, promptTemplate)
        : appState.promptManager.composeMessages(selectionText, 'amendment', promptTemplate);

    if (messages.length === 0) {
        throw new Error('No prompt composed — check the skill template');
    }

    const promptText = _flattenMessages(messages);
    const rawResponse = (onToken || onReasoning)
        ? await sendPromptStream(backendConfig, promptText, { onContent: onToken, onReasoning }, log)
        : await sendPrompt(backendConfig, promptText, log);
    log(`LLM response received [${backendConfig.model}]`, 'success');

    if (!merged) {
        return {
            selectionText,
            amendedText: stripMarkdown(rawResponse, log),
            commentText: null,
            model: backendConfig.model,
        };
    }

    // Merged mode: parse the ===AMENDMENT=== / ===COMMENT=== protocol,
    // with the fallback classification call preserved from the old handler.
    let parsed = parseDelimitedResponse(rawResponse);

    if (parsed.amendment === null) {
        log('Response missing delimiters, attempting to classify...', 'info');
        const fallbackMessages = buildFallbackClassificationPrompt(rawResponse, selectionText);
        try {
            const fallbackResponse = await sendPrompt(backendConfig, _flattenMessages(fallbackMessages), log);
            parsed = parseDelimitedResponse(fallbackResponse);
            if (parsed.amendment === null) {
                log('Could not split response into amendment and comment', 'warning');
                parsed = { amendment: rawResponse.trim(), comment: null, raw: rawResponse };
            }
        } catch (fallbackError) {
            log(`Fallback classification failed: ${fallbackError.message}`, 'warning');
            parsed = { amendment: rawResponse.trim(), comment: null, raw: rawResponse };
        }
    }

    return {
        selectionText,
        amendedText: parsed.amendment ? stripMarkdown(parsed.amendment, log) : null,
        commentText: parsed.comment || null,
        model: backendConfig.model,
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
 * @returns {Promise<object>} Proposal with tablePatch + tableItems
 * @private
 */
async function _prepareTableAmendment(deps, { tableRegion, promptTemplate, commentInstructions, backendConfig, onToken, onReasoning }) {
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

    const messages = [];
    const contextPrompt = appState.promptManager.getActivePrompt('context');
    if (contextPrompt) {
        messages.push({ role: 'system', content: contextPrompt.template });
    }
    messages.push({ role: 'user', content: buildTableUserPrompt(instruction, cells, { rowCount, colCount }) });

    const promptText = _flattenMessages(messages);
    const rawResponse = (onToken || onReasoning)
        ? await sendPromptStream(backendConfig, promptText, { onContent: onToken, onReasoning }, log)
        : await sendPrompt(backendConfig, promptText, log);
    log(`LLM response received [${backendConfig.model}]`, 'success');

    // Full-table grid for no-op detection and per-cell "before" text.
    const originals = Array.from({ length: rowCount }, () => Array(colCount).fill(''));
    for (const c of cells) originals[c.row - 1][c.col - 1] = c.text;

    const patch = parseTablePatchResponse(rawResponse, { rowCount, colCount, originals });
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
        tablePatch: { rowCount, colCount, cells: patch.cells, rowOps: patch.rowOps },
        tableItems,
    };
}

/**
 * Applies a prepared selection amendment to the document as tracked changes
 * (per config.trackChangesEnabled), then inserts the optional comment.
 *
 * @param {object} deps - { appState, log }
 * @param {object} proposal - Result of prepareSelectionAmendment
 */
export async function applySelectionAmendment(deps, proposal) {
    const { appState, log } = deps;
    const { selectionText, amendedText, commentText } = proposal;

    // Table route: per-cell tracked revisions plus row-level structure ops.
    // Table proposals never carry a comment (merged mode is skipped).
    if (proposal.tablePatch) {
        await _applyTablePatch(deps, proposal);
        return;
    }

    if (amendedText) {
        log('Applying changes...', 'info');
        const trackChanges = !!appState.config.trackChangesEnabled;
        await Word.run(async (context) => {
            const selection = context.document.getSelection();
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
        log('Changes applied successfully', 'success');
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
 * @private
 */
async function _applyTablePatch(deps, proposal) {
    const { appState, log } = deps;
    const { tablePatch } = proposal;
    const trackChanges = !!appState.config.trackChangesEnabled;
    const rowTracking = trackChanges && supportsTrackedRowOps(appState.platform);

    if (tablePatch.rowOps.length > 0 && trackChanges && !rowTracking) {
        log(`Row insertions/deletions cannot be tracked as revisions on this host (${appState.platform}) — applying them directly. Cell text edits are still tracked.`, 'warning');
    }

    log('Applying table patch...', 'info');
    let cellsApplied = 0;
    let cellsSkipped = 0;
    let rowOpsApplied = 0;

    await Word.run(async (context) => {
        const selection = context.document.getSelection();
        const table = selection.parentTableOrNullObject;
        table.load('isNullObject,rowCount');
        await context.sync();
        if (table.isNullObject) {
            throw new Error('The selection is no longer inside the table — re-select the region and apply again.');
        }
        if (table.rowCount !== tablePatch.rowCount) {
            throw new Error(`Table changed since this proposal was drafted (${tablePatch.rowCount} → ${table.rowCount} rows). Draft a new edit instead.`);
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
                    const applied = await _patchCell(context, table, cellPatch, log);
                    if (applied) cellsApplied++; else cellsSkipped++;
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
                    const row = table.getCell(op.row - 1, 0).parentRow;
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
        } finally {
            // Later turns and comments must not inherit tracking state.
            if (Word.ChangeTrackingMode) {
                context.document.changeTrackingMode = Word.ChangeTrackingMode.off;
                await context.sync();
            }
        }
    });

    log(`Table patch applied: ${cellsApplied} cell(s) revised, ${rowOpsApplied} row op(s)` +
        (cellsSkipped ? `, ${cellsSkipped} cell(s) already up to date` : ''), 'success');
}

/**
 * Revises one cell's text. Single-paragraph cells go through the granular
 * diff strategies (tracked in-cell edits are native to Word); the rare
 * multi-paragraph cell falls back to per-paragraph replacement when the line
 * count matches, else a coarse whole-content replace. Returns true when the
 * cell was written, false when its text already matched.
 *
 * @private
 */
async function _patchCell(context, table, cellPatch, log) {
    const label = `R${cellPatch.row}C${cellPatch.col}`;
    const cell = table.getCell(cellPatch.row - 1, cellPatch.col - 1);
    const paragraphs = cell.body.paragraphs;
    paragraphs.load('items');
    await context.sync();
    const items = paragraphs.items;
    if (items.length === 0) return false; // a cell always holds ≥1 paragraph

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
export async function prepareDocumentAppend(deps, { instruction, selectionText, onToken, onReasoning } = {}) {
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
        ? await sendPromptStream(backendConfig, prompt, { onContent: onToken, onReasoning }, log)
        : await sendPrompt(backendConfig, prompt, log);
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
export async function planDocumentTasks(deps, { instruction, hasSelection = false, onToken, onReasoning } = {}) {
    const { appState, log } = deps;

    const prompt = buildPlanPrompt(instruction, hasSelection);
    const backendConfig = getActiveBackendConfig(appState);
    log(`Planning tasks [${backendConfig.model}]...`, 'info');
    const rawResponse = (onToken || onReasoning)
        ? await sendPromptStream(backendConfig, prompt, { onContent: onToken, onReasoning }, log)
        : await sendPrompt(backendConfig, prompt, log);

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
export async function prepareFormatProposal(deps, { instruction, scope = 'selection', selectionText, onToken, onReasoning } = {}) {
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
        ? await sendPromptStream(backendConfig, prompt, { onContent: onToken, onReasoning }, log)
        : await sendPrompt(backendConfig, prompt, log);

    const ops = parseFormatOps(rawResponse, log);
    log(`Parsed ${ops.length} formatting op(s) from the model response.`, 'info');
    return { instruction, scope, ops, model: backendConfig.model };
}

/**
 * Applies a prepared format proposal. Insert ops add their paragraph(s) at
 * the scope start/end; other ops' targets are resolved inside the scope
 * range (whole scope, substring matches, or paragraphs of a given built-in
 * style), then font/paragraph properties are set with change tracking per
 * config (Word records them as Formatted revisions).
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
                inserted += _applyInsertOp(scopeRange, op, log);
                await context.sync();
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
 * the text. Only queues commands — the caller syncs. Returns the number of
 * paragraphs inserted.
 * @private
 */
function _applyInsertOp(scopeRange, op, log) {
    const paragraphs = op.insert.text.split(/\n+/).map((p) => p.trim()).filter(Boolean);
    const atStart = op.insert.position === 'start';
    const location = atStart ? Word.InsertLocation.start : Word.InsertLocation.end;
    const ordered = atStart ? [...paragraphs].reverse() : paragraphs;
    for (const text of ordered) {
        const paragraph = scopeRange.insertParagraph(text, location);
        if (op.font) _applyFontOps(paragraph.font, op.font, log);
        if (op.paragraph) _applyParagraphOps(paragraph, op.paragraph, log);
    }
    return paragraphs.length;
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
 * 'dark blue' -> Word.HighlightColor.darkBlue). Returns undefined on miss.
 * @private
 */
function _enumValue(enumObj, name) {
    if (!enumObj || name === undefined || name === null) return undefined;
    const wanted = String(name).replace(/[\s_-]+/g, '').toLowerCase();
    const key = Object.keys(enumObj).find((k) => k.toLowerCase() === wanted);
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
 * Applies validated paragraph ops to a Word.Paragraph object.
 * @private
 */
function _applyParagraphOps(paragraph, ops, log) {
    for (const [key, value] of Object.entries(ops)) {
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
 * @returns {Promise<{ instruction: string, svg: string|null, position: string, model: string }>}
 *   svg is null when the model returned nothing usable.
 */
export async function prepareIllustrationProposal(deps, { instruction, onToken, onReasoning } = {}) {
    const { appState, log } = deps;

    const richness = (appState.config.docExtraction || {}).richness || 'structured';
    log('Extracting document text for context...', 'info');
    const documentText = await extractDocumentStructured({ richness });

    const prompt = buildIllustrationPrompt(instruction, documentText);
    const backendConfig = getActiveBackendConfig(appState);
    log(`Designing illustration [${backendConfig.model}]...`, 'info');
    const rawResponse = (onToken || onReasoning)
        ? await sendPromptStream(backendConfig, prompt, { onContent: onToken, onReasoning }, log)
        : await sendPrompt(backendConfig, prompt, log);

    const parsed = parseIllustration(rawResponse, log);
    const svg = parsed ? ensureSvgDimensions(sanitizeSvg(parsed.svg)) : null;
    if (svg) log(`Illustration SVG received (${(svg.length / 1024).toFixed(1)} KB).`, 'success');
    return {
        instruction: (instruction || '').trim(),
        svg,
        position: illustrationPositionFromInstruction(instruction),
        model: backendConfig.model,
    };
}

/**
 * Applies a prepared illustration proposal: rasterizes the sanitized SVG to
 * PNG (Word's insertInlinePictureFromBase64 takes PNG/JPEG/GIF/BMP base64,
 * not SVG) and inserts it as a centered inline picture in its own paragraph
 * at the document start or end, as a tracked change per config.
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
    const { base64, width, height } = await _svgToPngBase64(svg);
    const atStart = proposal.position === 'start';

    await Word.run(async (context) => {
        if (Word.ChangeTrackingMode) {
            context.document.changeTrackingMode = appState.config.trackChangesEnabled
                ? Word.ChangeTrackingMode.trackAll
                : Word.ChangeTrackingMode.off;
        }
        const body = context.document.body;
        const location = atStart ? Word.InsertLocation.start : Word.InsertLocation.end;
        // A dedicated paragraph keeps the image standalone and centerable.
        const paragraph = body.insertParagraph('', location);
        const picture = paragraph.getRange(Word.RangeLocation.start)
            .insertInlinePictureFromBase64(base64, Word.InsertLocation.start);
        paragraph.alignment = Word.Alignment.centered;
        picture.load('width,height');
        await context.sync();
        // Inline picture dimensions are points; scale oversized images down
        // to the content width, keeping the aspect ratio.
        if (picture.width > MAX_ILLUSTRATION_WIDTH_PT) {
            picture.height = picture.height * (MAX_ILLUSTRATION_WIDTH_PT / picture.width);
            picture.width = MAX_ILLUSTRATION_WIDTH_PT;
        }
        try {
            picture.altTextDescription = (proposal.instruction || 'Illustration').slice(0, 200);
        } catch (_e) {
            // Alt text is best-effort (not critical for the insertion).
        }
        await context.sync();
        if (Word.ChangeTrackingMode) {
            context.document.changeTrackingMode = Word.ChangeTrackingMode.off;
            await context.sync();
        }
        log(`Inserted illustration at the document ${atStart ? 'start' : 'end'} (${width}x${height}px PNG).`, 'success');
    });
    return { inserted: true };
}

/**
 * Rasterizes an SVG string to PNG base64 at 1600px wide via an offscreen
 * canvas (browser/WebView2 only — not exercised under node tests). The SVG
 * is vector, so rendering up to 1600px keeps the inserted image crisp.
 * @private
 */
async function _svgToPngBase64(svg) {
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
export async function runDocumentSkill(deps, { category, promptTemplate, commentInstructions = '', onProgress, onChunkToken, gateApply = false } = {}) {
    const { appState, log, logWithRetry } = deps;
    const signal = appState.processDocController ? appState.processDocController.signal : undefined;
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
     * @returns {Promise<object>} applicationResult from applyChunkResults
     */
    const apply = async (chunkIds) => {
        const selected = Array.isArray(chunkIds)
            ? results.filter((r) => !r.amendment || (r.chunk && chunkIds.includes(r.chunk.id)))
            : results;
        log(`Applying changes to document${Array.isArray(chunkIds) ? ` (${selected.length} of ${results.length} section(s))` : ''}...`, 'info');
        const applicationResult = await applyChunkResults(selected, bookmarkMap, {
            trackChangesEnabled: appState.config.trackChangesEnabled,
            lineDiffEnabled: appState.config.lineDiffEnabled,
            log,
            commentGranularity: appState.config.commentGranularity,
        });

        await cleanupBookmarks(bookmarkMap);

        log(
            `Document processed: ${chunks.length} chunks, ` +
            `${applicationResult.amendmentsApplied} amendments applied, ` +
            `${applicationResult.commentsInserted} comments inserted` +
            (failed > 0 ? `, ${failed} chunks failed` : '') +
            (cancelled > 0 ? `, ${cancelled} chunks cancelled` : ''),
            failed > 0 ? 'warning' : 'success'
        );

        if (failed > 0 && logWithRetry) {
            const failedChunks = results.filter(r => r.status === 'rejected');
            logWithRetry(
                `${failed} chunk(s) failed. Click to retry failed chunks.`,
                'warning',
                () => retryFailedChunks(deps, { failedResults: failedChunks, bookmarkMap, backendConfig, promptShim, onProgress })
            );
        }

        return applicationResult;
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
    // stage a proposal card for user confirmation.
    if (gateApply) {
        return { staged: true, results, chunks, apply, discard, failedCount: failed, cancelledCount: cancelled };
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
        const retryChunks = failedResults.map(r => ({
            id: r.chunkId,
            text: r.originalText || '',
            tokenCount: r.originalText ? Math.ceil(r.originalText.length / 4) : 0,
            overlapText: '',
        }));

        const results = await processChunksParallel(retryChunks, {
            config: backendConfig,
            promptManager: promptShim,
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
            commentGranularity: appState.config.commentGranularity,
            chunkOriginals,
        });

        const stillFailed = results.filter(r => r.status === 'rejected').length;
        log(
            `Retry complete: ${applicationResult.amendmentsApplied} amendments, ` +
            `${applicationResult.commentsInserted} comments` +
            (stillFailed > 0 ? `, ${stillFailed} still failed` : ''),
            stillFailed > 0 ? 'warning' : 'success'
        );
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
export async function runSummarySkill(deps, { promptTemplate, onToken, onReasoning } = {}) {
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
        ? await sendPromptStream(backendConfig, promptText, { onContent: onToken, onReasoning }, log)
        : await sendPrompt(backendConfig, promptText, log);
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
 *   focused excerpt before the full document context
 * @param {function} [args.onToken] - Called with each streamed token
 * @param {function} [args.onReasoning] - Called with each streamed thinking token
 * @param {function} [args.onStatus] - Called with stage updates ("Reading the document...", "Waiting for model...")
 * @param {AbortSignal} [args.signal] - Cancellation signal
 * @returns {Promise<string>} The full answer text
 */
export async function answerQuestion(deps, { question, skillTemplate, selectionText, onToken, onReasoning, onStatus, signal } = {}) {
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
    if (selectionText && selectionText.trim()) {
        prompt += '\n\n--- SELECTED TEXT (the user is asking about this excerpt) ---\n' + selectionText.trim();
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
