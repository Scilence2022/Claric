/**
 * Agent Actions
 *
 * Word-side glue for the tool-calling stack (lib/tool-registry.js,
 * lib/tool-loop.js, lib/table-model.js, lib/image-model.js). Two prepare
 * halves, both staging proposals for the existing card + Apply flow:
 *
 * - prepareTableToolEdit: multi-step table edits. The selection's table
 *   region seeds a table draft model; the loop's translated patch reuses
 *   the existing tablePatch proposal shape, so applySelectionAmendment and
 *   the table proposal card serve unchanged.
 * - prepareImageToolEdit: image management (design/insert, replace,
 *   delete, resize, alt text) plus read_image content inspection. A
 *   snapshot of inline-picture metadata seeds an image draft model;
 *   applyImageOps executes the staged ops. Selection metadata maps onto
 *   snapshot indexes so the loop knows which image the user selected.
 *
 * Nested LLM calls: design_illustration / replace_illustration run the
 * existing illustration design prompt host-side inside tool execution.
 *
 * Dependencies flow in via `deps` like word-actions; module-level code is
 * side-effect free.
 *
 * @module agent-actions
 */

import { buildToolLoopSystemPrompt } from '../lib/tool-registry.js';
import { runToolLoop } from '../lib/tool-loop.js';
import { createTableModel, executeTableTool, TABLE_TOOL_SPECS } from '../lib/table-model.js';
import { createImageModel, IMAGE_TOOL_SPECS } from '../lib/image-model.js';
import { sendMessages, sendPrompt } from '../lib/llm-client.js';
import {
    buildIllustrationPrompt, parseIllustration, sanitizeSvg, ensureSvgDimensions,
} from '../lib/illustration.js';
import { extractDocumentStructured } from '../lib/comment-extractor.js';
import { readSelectionTableRegion, svgToPngBase64, insertPngPicture, finalizeInsertedPicture, imageDataUrl } from './word-actions.js';
import { getActiveBackendConfig } from './app-state.js';

/** Step budgets per loop kind. Tables get more steps (per-cell work). */
const STEP_BUDGETS = Object.freeze({ table: 14, image: 8 });

/** Per-step request timeout (loops make several calls; keep each bounded). */
const STEP_TIMEOUT_MS = 180000;

/** Max base64 chars for one read_image attachment (≈4.5MB binary). */
const MAX_READ_IMAGE_CHARS = 6 * 1024 * 1024;

/**
 * Sends one loop turn. When the history carries image attachments (image_url
 * parts from read_image observations) and the backend rejects the request
 * with an HTTP 4xx — typical for text-only models — retries once with the
 * attachments stripped, so the loop continues text-only instead of erroring
 * the whole turn. Abort/timeout/5xx propagate untouched.
 *
 * @private
 */
async function _sendLoopMessages(deps, backendConfig, messages, signal) {
    const { log } = deps;
    try {
        return await sendMessages(backendConfig, messages, log, signal, STEP_TIMEOUT_MS);
    } catch (err) {
        if (err.name === 'AbortError' || err.name === 'TimeoutError' || !/^HTTP 4\d\d/.test(err.message || '')) {
            throw err;
        }
        const carriesImages = messages.some((m) => Array.isArray(m.content));
        if (!carriesImages) throw err;
        log(`Backend rejected image inputs (${err.message}); retrying without image attachments.`, 'warning');
        const stripped = messages.map((m) => {
            if (!Array.isArray(m.content)) return m;
            const textParts = m.content.filter((p) => p && p.type !== 'image_url');
            return {
                role: m.role,
                content: textParts.length ? textParts : [{ type: 'text', text: '(image attachment removed)' }],
            };
        });
        return sendMessages(backendConfig, stripped, log, signal, STEP_TIMEOUT_MS);
    }
}

/**
 * Runs one tool loop with the standard send wiring.
 *
 * @private
 */
async function _runLoop(deps, { systemPrompt, taskPrompt, tools, execute, maxSteps, signal, onStep }) {
    const backendConfig = getActiveBackendConfig(deps.appState);
    return runToolLoop({
        systemPrompt,
        taskPrompt,
        tools,
        execute,
        maxSteps,
        signal,
        onStep,
        send: (messages) => _sendLoopMessages(deps, backendConfig, messages, signal),
    });
}

/**
 * Prepares a multi-step table edit via the tool loop. Returns a proposal
 * shaped exactly like the single-shot table route of
 * prepareSelectionAmendment ({ tablePatch, tableItems, ... }), so the
 * selection-edit turn runner, proposal card, and applySelectionAmendment
 * are reused unchanged.
 *
 * @param {object} deps - { appState, log }
 * @param {object} args
 * @param {string} args.instruction - The (chained) edit instruction
 * @param {AbortSignal} [args.signal]
 * @param {function} [args.onStep] - Loop activity hook
 * @returns {Promise<object|null>} Table-route proposal, or null when the
 *   selection is not a multi-cell table region (caller falls back).
 */
export async function prepareTableToolEdit(deps, { instruction, signal, onStep } = {}) {
    const { log } = deps;
    const tableRegion = await readSelectionTableRegion(deps);
    if (!tableRegion) return null;

    const model = createTableModel(tableRegion);
    log(`Table tool loop: ${tableRegion.rowCount}×${tableRegion.colCount} region` +
        `${tableRegion.merged ? ' (merged cells — row ops disabled)' : ''}.`, 'info');

    const state = model.getState().result;
    const taskPrompt =
        `USER TASK: ${(instruction || '').trim()}\n\n` +
        'The selection covers this table region (1-based absolute coordinates):\n\n' +
        `${state.grid}\n\n` +
        `Covered region: ${state.coveredRegion}. Row operations allowed: ${state.rowOpsAllowed ? 'yes' : 'no'}.\n` +
        (tableRegion.mergedUnknown
            ? 'The table contains merged cells whose layout could not be mapped — edits to merge-covered coordinates are skipped at apply time.\n'
            : '') +
        'Work through the task with the table tools. Verify with get_state when unsure.';

    const loop = await _runLoop(deps, {
        systemPrompt: buildToolLoopSystemPrompt(TABLE_TOOL_SPECS, { maxSteps: STEP_BUDGETS.table }),
        taskPrompt,
        tools: TABLE_TOOL_SPECS,
        execute: (name, args) => executeTableTool(model, name, args),
        maxSteps: STEP_BUDGETS.table,
        signal,
        onStep,
    });

    if (!loop.finished) {
        log(`Table tool loop hit the ${STEP_BUDGETS.table}-step budget — applying the ops recorded so far.`, 'warning');
    }
    if (loop.summary) log(`Tool loop summary: ${loop.summary}`, 'info');

    const patch = model.toTablePatch();
    if (patch.cells.length === 0 && patch.rowOps.length === 0) {
        const err = new Error('The tool loop produced no table changes. Try rephrasing the instruction.');
        err.noChanges = true;
        throw err;
    }

    // Same review-item derivation as the single-shot table route.
    const originals = patch.originals;
    const tableItems = [
        ...patch.cells.map((c) => ({
            label: `Cell R${c.row}C${c.col}`,
            before: (originals[c.row - 1] || [])[c.col - 1] || '',
            after: c.text,
            searchText: ((originals[c.row - 1] || [])[c.col - 1] || '').trim().slice(0, 60) || undefined,
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
        selectionText: tableRegion.cells.map((c) => c.text).join('\n'),
        amendedText: null,
        commentText: null,
        model: getActiveBackendConfig(deps.appState).model,
        tablePatch: patch,
        tableItems,
        toolLoop: { steps: loop.steps, finished: loop.finished },
    };
}

/**
 * Snapshots the document's inline pictures for the image draft model.
 *
 * @private
 */
async function _snapshotImages() {
    const snapshot = [];
    await Word.run(async (context) => {
        const pictures = context.document.body.inlinePictures;
        pictures.load('items');
        await context.sync();
        const items = pictures.items || [];
        for (const pic of items) pic.load('width,height,altTextDescription');
        await context.sync();
        items.forEach((pic, i) => {
            snapshot.push({
                index: i + 1,
                width: pic.width,
                height: pic.height,
                altText: pic.altTextDescription || '',
            });
        });
    });
    return snapshot;
}

/**
 * Runs the nested illustration design call for design_/replace_ tools.
 *
 * @private
 */
async function _designSvg(deps, { instruction, documentText, signal }) {
    const { log } = deps;
    const backendConfig = getActiveBackendConfig(deps.appState);
    log(`Designing illustration via tool call [${backendConfig.model}]...`, 'info');
    const raw = await sendPrompt(
        backendConfig, buildIllustrationPrompt(instruction, documentText), log, signal
    );
    const parsed = parseIllustration(raw, log);
    return parsed ? ensureSvgDimensions(sanitizeSvg(parsed.svg)) : null;
}

/**
 * Reads one snapshot-indexed picture's visual content for the read_image
 * tool: base64 bytes as a data URL plus its dimensions.
 *
 * @param {number} index - 1-based document-order snapshot index
 * @returns {Promise<{dataUrl: string, width: number, height: number}>}
 * @throws {Error} When the index is out of range or the image is too large
 * @private
 */
async function _readImageAttachment(index) {
    if (!Number.isInteger(index) || index < 1) {
        throw new Error('"index" must be a 1-based snapshot index.');
    }
    let out = null;
    await Word.run(async (context) => {
        const pictures = context.document.body.inlinePictures;
        pictures.load('items');
        await context.sync();
        const items = pictures.items || [];
        const pic = items[index - 1];
        if (!pic) {
            throw new Error(`Image ${index} does not exist in this snapshot (${items.length} picture(s)).`);
        }
        pic.load('width,height');
        const b64 = pic.getBase64ImageSrc();
        await context.sync();
        if (!b64.value) {
            throw new Error(`Image ${index} returned no image data.`);
        }
        if (b64.value.length > MAX_READ_IMAGE_CHARS) {
            throw new Error(`Image ${index} is too large to attach (${(b64.value.length / 1048576).toFixed(1)}MB base64).`);
        }
        out = { dataUrl: imageDataUrl(b64.value), width: pic.width, height: pic.height };
    });
    return out;
}

/**
 * Prepares image operations via the tool loop: design/insert, replace,
 * delete, resize, alt text. Returns a proposal with card items; apply
 * happens only via applyImageOps when the user clicks Apply.
 *
 * A read-only loop (no ops, but a finish summary — e.g. the model inspected
 * the selected image with read_image and answered a question) resolves to a
 * {noOps: true, answer} proposal instead of staging a card.
 *
 * @param {object} deps - { appState, log }
 * @param {object} args
 * @param {string} args.instruction - The image-management instruction
 * @param {Array<{width: number, height: number, altText: string}>} [args.selectionImages] -
 *   Metadata of the pictures inside the CURRENT selection (from
 *   readSelectionContent); matched onto snapshot indexes so the task prompt
 *   can tell the model which image the user is pointing at
 * @param {AbortSignal} [args.signal]
 * @param {function} [args.onStep] - Loop activity hook
 * @returns {Promise<{ instruction: string, ops: Array<object>, items: Array<object>,
 *   snapshotCount: number, model: string, toolLoop: object }>}
 * @throws {Error} When the loop records no ops and produces no answer
 */
export async function prepareImageToolEdit(deps, { instruction, selectionImages, signal, onStep } = {}) {
    const { appState, log } = deps;

    log('Reading document images...', 'info');
    const snapshot = await _snapshotImages();
    log(`Found ${snapshot.length} inline picture(s).`, 'info');

    // Document context for the nested design calls (extracted once).
    const richness = (appState.config.docExtraction || {}).richness || 'structured';
    const documentText = await extractDocumentStructured({ richness });

    const model = createImageModel(snapshot);

    const execute = async (name, args) => {
        switch (name) {
            case 'list_images':
                return model.listImages();
            case 'read_image': {
                // Throws on bad index/oversized image — the loop turns the
                // throw into an error observation the model can react to.
                const img = await _readImageAttachment(args.index);
                return {
                    ok: true,
                    result: {
                        index: args.index,
                        widthPt: img.width,
                        heightPt: img.height,
                        note: 'the image is attached to this observation as an image input — look at it',
                    },
                    attachments: [{ dataUrl: img.dataUrl }],
                };
            }
            case 'design_illustration': {
                const svg = await _designSvg(deps, { instruction: args.instruction, documentText, signal });
                return model.recordInsert({ position: args.position, instruction: args.instruction, svg });
            }
            case 'replace_illustration': {
                const svg = await _designSvg(deps, { instruction: args.instruction, documentText, signal });
                return model.recordReplace({ index: args.index, instruction: args.instruction, svg });
            }
            case 'delete_image':
                return model.recordDelete(args.index);
            case 'resize_image':
                return model.recordResize(args.index, args.widthPt);
            case 'set_alt_text':
                return model.recordAltText(args.index, args.text);
            default:
                return { ok: false, error: `Unknown image tool "${name}".` };
        }
    };

    // Selection focus: map the selected pictures (metadata) onto snapshot
    // indexes, first-match-wins — duplicates beyond the first stay unmapped.
    // Advisory only: every tool remains index-addressed against the snapshot.
    const focusLines = [];
    if (Array.isArray(selectionImages) && selectionImages.length > 0) {
        const used = new Set();
        for (const sel of selectionImages) {
            const hit = snapshot.find((img, i) => !used.has(i)
                && img.width === sel.width
                && img.height === sel.height
                && (img.altText || '') === (sel.altText || ''));
            if (hit) {
                used.add(hit.index - 1);
                focusLines.push(`- image ${hit.index} (SELECTED by the user right now)`);
            }
        }
        if (focusLines.length === 0) {
            focusLines.push('- (the selected picture(s) could not be matched to a snapshot index)');
        }
    }

    const taskPrompt =
        `USER TASK: ${(instruction || '').trim()}\n\n` +
        `The document has ${snapshot.length} inline picture(s):\n` +
        (snapshot.length
            ? snapshot.map((img) =>
                `- image ${img.index}: ${img.width}x${img.height}pt${img.altText ? `, alt "${img.altText.slice(0, 60)}"` : ''}`
            ).join('\n')
            : '(none)') +
        (focusLines.length > 0
            ? '\n\nThe user\'s current selection in the document:\n' +
              focusLines.join('\n') +
              '\nWhen the task says "this/that image" (这张/此图), it means the selected one(s); use read_image on it before answering questions about its content.'
            : '') +
        '\n\nWork through the task with the image tools. Indexes refer to this snapshot.';

    const loop = await _runLoop(deps, {
        systemPrompt: buildToolLoopSystemPrompt(IMAGE_TOOL_SPECS, { maxSteps: STEP_BUDGETS.image }),
        taskPrompt,
        tools: IMAGE_TOOL_SPECS,
        execute,
        maxSteps: STEP_BUDGETS.image,
        signal,
        onStep,
    });

    if (!loop.finished) {
        log(`Image tool loop hit the ${STEP_BUDGETS.image}-step budget — staging the ops recorded so far.`, 'warning');
    }
    if (loop.summary) log(`Tool loop summary: ${loop.summary}`, 'info');

    if (model.ops.length === 0) {
        if (loop.finished && loop.summary && loop.summary.trim()) {
            return {
                noOps: true,
                answer: loop.summary.trim(),
                instruction: (instruction || '').trim(),
                ops: [],
                items: [],
                snapshotCount: snapshot.length,
                model: getActiveBackendConfig(appState).model,
                toolLoop: { steps: loop.steps, finished: loop.finished },
            };
        }
        const err = new Error('The tool loop proposed no image changes.');
        err.noChanges = true;
        throw err;
    }

    return {
        instruction: (instruction || '').trim(),
        ops: model.ops,
        items: model.describeOps(),
        snapshotCount: snapshot.length,
        model: getActiveBackendConfig(appState).model,
        toolLoop: { steps: loop.steps, finished: loop.finished },
    };
}

/**
 * Applies staged image ops. Index-addressed ops (replace/delete/resize/
 * altText) resolve against a fresh picture list first — a snapshot-count
 * mismatch means the document changed since staging and the apply aborts;
 * inserts run last so they never shift the resolved indexes.
 *
 * @param {object} deps - { appState, log }
 * @param {object} proposal - Result of prepareImageToolEdit (ops may be a
 *   user-filtered subset — unchecked card items are skipped)
 * @returns {Promise<{ applied: number, warnings: string[] }>}
 */
export async function applyImageOps(deps, proposal) {
    const { appState, log } = deps;
    const ops = (proposal && proposal.ops) || [];
    if (ops.length === 0) throw new Error('No image operations to apply.');

    const indexOps = ops.filter((op) => op.type !== 'insert');
    const insertOps = ops.filter((op) => op.type === 'insert');

    let applied = 0;
    const warnings = [];

    await Word.run(async (context) => {
        const pictures = context.document.body.inlinePictures;
        pictures.load('items');
        await context.sync();
        const items = pictures.items || [];
        if (items.length !== (proposal.snapshotCount || 0)) {
            throw new Error(
                `The document's images changed since this proposal was drafted ` +
                `(${proposal.snapshotCount} → ${items.length}). Draft a new edit instead.`
            );
        }
        for (const pic of items) pic.load('width,height');
        await context.sync();

        if (Word.ChangeTrackingMode) {
            context.document.changeTrackingMode = appState.config.trackChangesEnabled
                ? Word.ChangeTrackingMode.trackAll
                : Word.ChangeTrackingMode.off;
        }

        try {
            // Phase 1: index-addressed ops against the snapshot mapping.
            for (const op of indexOps) {
                const pic = items[op.index - 1];
                if (!pic) {
                    warnings.push(`Image ${op.index} no longer exists — op skipped.`);
                    continue;
                }
                if (op.type === 'delete') {
                    pic.delete();
                } else if (op.type === 'resize') {
                    const ratio = (pic.height || 1) / (pic.width || 1);
                    pic.width = op.widthPt;
                    pic.height = Math.round(op.widthPt * ratio);
                } else if (op.type === 'altText') {
                    pic.altTextDescription = op.text;
                } else if (op.type === 'replace') {
                    const { base64 } = await svgToPngBase64(op.svg);
                    // Insert the replacement at the old picture's position,
                    // then delete the old one — under tracking this records
                    // as a delete + insert pair, individually rejectable.
                    const range = pic.getRange(Word.RangeLocation.start);
                    const newPic = range.insertInlinePictureFromBase64(base64, Word.InsertLocation.before);
                    newPic.load('width,height');
                    pic.delete();
                }
                applied++;
            }
            await context.sync();

            // Phase 2: inserts (never shift the snapshot indexes above).
            for (const op of insertOps) {
                const { base64 } = await svgToPngBase64(op.svg);
                const pic = insertPngPicture(context, {
                    base64,
                    position: op.position,
                });
                await context.sync();
                finalizeInsertedPicture(context, pic, op.instruction);
                await context.sync();
                applied++;
            }
        } finally {
            if (Word.ChangeTrackingMode) {
                context.document.changeTrackingMode = Word.ChangeTrackingMode.off;
                await context.sync();
            }
        }
    });

    log(`Applied ${applied} image operation(s)` +
        (warnings.length ? ` (${warnings.length} skipped)` : '') + '.', 'success');
    return { applied, warnings };
}
