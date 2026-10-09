/** Native formatting contracts are loaded before any formatting reads or writes. */
import { buildFormatPrompt, parseFormatOps } from '../lib/format-ops.js';
import { requestsBodyFormatting, explicitBodyAlignmentOp } from '../lib/format-targets.js';
import { readFormatInventory, compileFormatTargetPlan, validateFormatTargetPlan, applyVerifiedFormatTargets } from './format-target-runtime.js';
import { requestsEmptyParagraphCleanup, inspectEmptyParagraphXml } from '../lib/empty-paragraphs.js';
import { rangeStructureFingerprint, rangeFingerprintDifference } from '../lib/ooxml-fingerprint.js';
import { extractJsonArray } from '../lib/json-utils.js';
import { getActiveBackendConfig } from './app-state.js';
import { sendMessagesStream } from '../lib/llm-client.js';
import { canReadWordVisuals, createWordVisualTools } from './word-render-tools.js';
import { loadFormatPlanning } from './task-module-loader.js';
import { enumValue as _enumValue, applyFontOps as _applyFontOps } from './word-format-helpers.js';

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
function checkOperationSignal(signal) {
    if (signal?.aborted) {
        const error = new Error('Operation cancelled.');
        error.name = 'AbortError';
        throw error;
    }
}

let formatAnchorSequence = 0;

/** The document body is already a stable native target across Word.run calls. */
function _documentFormatRange(context) {
    const body = context.document.body;
    if (!body || typeof body.getRange !== 'function') {
        throw new Error('This Word host cannot read the document formatting scope safely. No changes were applied.');
    }
    const range = body.getRange('Whole');
    if (!range || typeof range.load !== 'function' || typeof range.getOoxml !== 'function') {
        throw new Error('This Word host cannot read the document formatting baseline safely. No changes were applied.');
    }
    return range;
}

/** Native scope containment and XML evidence are required before deletion. */
async function _collectFormatEmptyParagraphs(context, scopeRange, signal, log) {
    const paragraphs = scopeRange.paragraphs;
    paragraphs.load('items/text');
    await context.sync();
    checkOperationSignal(signal);
    const lastRange = context.document.body.paragraphs.getLast().getRange('Whole');
    const indexes = [];
    const reasons = {};
    const preserve = (reason) => { reasons[reason] = (reasons[reason] || 0) + 1; };
    let candidates = 0;
    let unverifiable = 0;
    for (let index = 0; index < paragraphs.items.length; index++) {
        const paragraph = paragraphs.items[index];
        if ((paragraph.text || '').trim()) continue;
        candidates++;
        checkOperationSignal(signal);
        try {
            const range = paragraph.getRange('Whole');
            const relation = range.compareLocationWith(scopeRange);
            const finalRelation = range.compareLocationWith(lastRange);
            const table = paragraph.parentTableOrNullObject;
            table.load('isNullObject');
            // Paragraph XML identifies one native paragraph. A Whole RANGE
            // export may include additional block/container serialization.
            const xml = typeof paragraph.getOoxml === 'function' ? paragraph.getOoxml() : range.getOoxml();
            await context.sync();
            checkOperationSignal(signal);
            if (!['Inside', 'InsideStart', 'InsideEnd', 'Equal'].includes(relation.value)) {
                preserve('outside captured scope or partial paragraph');
            } else if (finalRelation.value === 'Equal') {
                preserve('final document paragraph');
            } else if (table.isNullObject !== true) {
                preserve('table cell');
            } else {
                const evidence = inspectEmptyParagraphXml(xml.value);
                if (evidence.deletable) indexes.push(index);
                else preserve(`${evidence.reason}${evidence.element ? ` (${evidence.element})` : ''}`);
            }
        } catch (error) {
            if (error.name === 'AbortError') throw error;
            unverifiable++;
            preserve('Word read failure');
            log(`Empty paragraph ${index + 1} could not be verified and will be preserved: ${error.message}`, 'warning');
        }
    }
    return { paragraphs, indexes, summary: { candidates, verified: indexes.length,
        preserved: candidates - indexes.length, unverifiable, reasons } };
}

export async function prepareFormatProposal(deps, { instruction, originalInstruction, scope = 'selection', selectionText, cleanupOnly = false,
    cleanupRequested = requestsEmptyParagraphCleanup(instruction), onToken, onReasoning, signal } = {}) {
    const { appState, log } = deps;
    checkOperationSignal(signal);
    const anchor = { bookmark: null };
    const originalAlignment = explicitBodyAlignmentOp(originalInstruction);
    const bodyOnly = !!originalAlignment || requestsBodyFormatting(instruction);
    let inventory = null;
    try {
        await Word.run(async (context) => {
            const documentScope = scope === 'document';
            const range = documentScope ? _documentFormatRange(context) : context.document.getSelection();
            if (!documentScope && (typeof range.insertBookmark !== 'function' || typeof context.document.getBookmarkRangeOrNullObject !== 'function')) {
                throw new Error('This Word host cannot anchor formatting safely. No changes were applied.');
            }
            range.load('text');
            await context.sync();
            checkOperationSignal(signal);
            if (scope === 'selection' && selectionText && range.text.trim() !== selectionText.trim()) {
                throw new Error('The selection changed before formatting was prepared. Draft a new proposal.');
            }
            let anchoredRange = range;
            if (documentScope) {
                // A bookmark over the whole body may omit table-cell/end-of-
                // body markers on Word for Mac. Whole-document authorization
                // has no selection boundary to recover: use the same native
                // body-range API at prepare/apply and verify its full baseline.
                anchor.kind = 'document-body';
            } else {
                const bookmark = `_claric_fmt_${Date.now().toString(36)}_${++formatAnchorSequence}`;
                // Retain ownership even if a later baseline read fails.
                anchor.bookmark = bookmark;
                range.insertBookmark(bookmark);
                await context.sync();
                checkOperationSignal(signal);
                // Selection/bookmark exports may have different native XML
                // boundaries. Use the SAME handle kind at prepare and apply.
                anchoredRange = context.document.getBookmarkRangeOrNullObject(bookmark);
                anchoredRange.load('isNullObject,text');
                await context.sync();
                checkOperationSignal(signal);
                if (anchoredRange.isNullObject || anchoredRange.text !== range.text) {
                    throw new Error('Word could not recover the exact formatting scope. No changes were applied.');
                }
            }
            const baseline = typeof anchoredRange.getOoxml === 'function' ? anchoredRange.getOoxml() : null;
            if (baseline) await context.sync();
            checkOperationSignal(signal);
            const structureFingerprint = rangeStructureFingerprint(baseline?.value);
            if (typeof anchoredRange.text !== 'string' || (documentScope && !structureFingerprint)) {
                throw new Error('Word could not verify the document formatting baseline. No changes were applied.');
            }
            Object.assign(anchor, { text: anchoredRange.text, ooxml: baseline?.value, structureFingerprint });
            log(`Formatting scope captured from ${documentScope ? 'document body' : 'bookmark'} (${scope}, baseline ${documentScope ? 'v3' : 'v2'}).`, 'info');
            if (!cleanupOnly) inventory = await readFormatInventory(context, anchoredRange, { scope, signal, log });
            if (cleanupRequested) {
                const { indexes, summary } = await _collectFormatEmptyParagraphs(context, anchoredRange, signal, log);
                anchor.cleanupIndexes = indexes;
                anchor.cleanupSummary = summary;
                log(`Found ${indexes.length} verified empty paragraph(s) in ${scope} scope.`, 'info');
                if (summary.preserved) {
                    log(`${summary.preserved} empty paragraph(s) will be preserved (${summary.unverifiable} read failures).`, 'warning');
                    for (const [reason, count] of Object.entries(summary.reasons)) {
                        log(`Empty paragraph cleanup: ${count} preserved — ${reason}.`, 'warning');
                    }
                }
            }
        });
        checkOperationSignal(signal);
        const backendConfig = getActiveBackendConfig(appState);
        let rawResponse = '[]';
        let rendering = null;
        const explicitAlignment = !cleanupOnly && !cleanupRequested ? originalAlignment || explicitBodyAlignmentOp(instruction) : null;
        if (explicitAlignment) {
            rawResponse = JSON.stringify([explicitAlignment]);
            log('Preparing explicit body alignment from verified Word paragraph evidence.', 'info');
        } else if (!cleanupOnly) {
            const prompt = buildFormatPrompt(instruction, anchor.text, scope, inventory?.descriptors);
            log(`Planning formatting ops [${backendConfig.model}]...`, 'info');
            if (canReadWordVisuals()) {
                const { planFormatWithRendering } = await loadFormatPlanning({ signal });
                const renderer = createWordVisualTools({ signal, log, scopeText: anchor.text });
                try {
                    const result = await planFormatWithRendering({ prompt, scopeText: anchor.text, renderer, signal, log,
                        validateOps: (ops) => {
                            const plan = compileFormatTargetPlan(inventory, ops, { bodyOnly });
                            return plan?.summary;
                        },
                        send: async (messages) => (await sendMessagesStream(backendConfig, messages, { onReasoning }, log, signal, 300000)).content,
                        onStep: (step) => { if (step.text) onToken?.(`${step.text}\n`); },
                    });
                    rawResponse = JSON.stringify(result.ops);
                    rendering = result.rendering;
                } finally { await renderer.dispose(); }
            } else {
                log('Native Word rendering is unavailable; planning formatting from text only.', 'info');
                rawResponse = await deps.sendActionRequest(deps, backendConfig, prompt, { onToken, onReasoning, signal });
            }
        }
        checkOperationSignal(signal);
        // Deletion is authorized by the original instruction, never by model
        // output. Host evidence supplies the count, not an invented LLM count.
        const ops = parseFormatOps(rawResponse, log).filter((op) => !op.cleanup);
        const parsedFormat = extractJsonArray(rawResponse);
        if (parsedFormat.error || !Array.isArray(parsedFormat.value)
            || parseFormatOps(rawResponse).length !== parsedFormat.value.length) {
            throw new Error('The formatting plan contains invalid operations or targets. No changes were applied.');
        }
        if (bodyOnly && !cleanupOnly && !ops.length) {
            throw new Error('No verified body-formatting plan was produced. No changes were applied.');
        }
        if (cleanupRequested && anchor.cleanupIndexes.length) {
            ops.push({ cleanup: { emptyParagraphs: true, emptyCount: anchor.cleanupIndexes.length } });
        }
        if (ops.length) anchor.targetPlan = compileFormatTargetPlan(inventory, ops, { bodyOnly });
        const targetSummary = anchor.targetPlan?.summary;
        if (targetSummary) {
            const uncertain = targetSummary.exclusions.filter((entry) => ['unknown', 'protected'].includes(entry.role)).length;
            targetSummary.uncertainParagraphs = bodyOnly ? uncertain : 0;
            log(`Formatting targets verified: ${targetSummary.verifiedParagraphs} paragraph(s); ${targetSummary.excludedParagraphs} excluded; ${targetSummary.uncertainParagraphs} uncertain.`, 'info');
        }
        if (!ops.length) await discardFormatProposal(deps, { anchor });
        return { instruction, scope, ops, anchor, rendering, targetSummary, cleanupSummary: anchor.cleanupSummary, model: backendConfig.model };
    } catch (error) {
        try { await discardFormatProposal(deps, { anchor }); }
        catch (cleanupError) { log(`Formatting anchor cleanup failed: ${cleanupError.message}`, 'warning'); }
        throw error;
    }
}

export async function discardFormatProposal(deps, proposal) {
    if (proposal?.anchor?.kind === 'document-body') {
        proposal.anchor.cleaned = true;
        return;
    }
    if (!proposal?.anchor?.bookmark || proposal.anchor.cleaned) return;
    await Word.run(async (context) => {
        context.document.deleteBookmark(proposal.anchor.bookmark);
        await context.sync();
    });
    proposal.anchor.cleaned = true;
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
export async function applyFormatProposal(deps, proposal, { signal } = {}) {
    const { appState, log } = deps;
    const { ops, anchor } = proposal || {};
    if (!Array.isArray(ops) || !ops.length) throw new Error('No formatting ops to apply.');
    const documentScope = anchor?.kind === 'document-body';
    if (typeof anchor?.text !== 'string' || (documentScope
        ? proposal.scope !== 'document' || !anchor.structureFingerprint || typeof anchor.ooxml !== 'string'
        : !anchor?.bookmark)) throw new Error('Formatting target is not anchored. Draft a new proposal.');
    if (anchor.attempted) throw new Error('This formatting proposal has already been attempted. Review the document and draft a new proposal.');
    if (anchor.cleaned) throw new Error('This formatting proposal has been discarded. Draft a new proposal.');
    checkOperationSignal(signal);
    let appliedRanges = 0;
    let insertedParagraphs = 0;
    let deletedParagraphs = 0;
    let interrupted = false;
    let partial = false;
    let verifiedResult = null;
    await Word.run(async (context) => {
        const scopeRange = documentScope ? _documentFormatRange(context) : context.document.getBookmarkRangeOrNullObject(anchor.bookmark);
        scopeRange.load(documentScope ? 'text' : 'isNullObject,text');
        if (Word.ChangeTrackingMode) context.document.load('changeTrackingMode');
        await context.sync();
        checkOperationSignal(signal);
        const targeted = anchor.targetPlan?.mode === 'targets';
        if ((!documentScope && scopeRange.isNullObject) || (!targeted && scopeRange.text !== anchor.text)) {
            throw new Error('The anchored formatting target changed or disappeared. Draft a new proposal.');
        }
        let currentInventory = null;
        let currentPlan = null;
        if (anchor.targetPlan) {
            currentInventory = await readFormatInventory(context, scopeRange, { scope: proposal.scope, signal, log });
            currentPlan = validateFormatTargetPlan(anchor.targetPlan, currentInventory, ops);
        }
        if (!targeted && anchor.ooxml !== undefined) {
            const current = scopeRange.getOoxml();
            await context.sync();
            checkOperationSignal(signal);
            const currentFingerprint = rangeStructureFingerprint(current.value);
            const unchanged = anchor.structureFingerprint
                ? currentFingerprint === anchor.structureFingerprint
                : current.value === anchor.ooxml;
            if (!unchanged) {
                const location = rangeFingerprintDifference(anchor.structureFingerprint, currentFingerprint);
                log(`Formatting baseline ${documentScope ? 'v3' : 'v2'} mismatch (${proposal.scope}): ${location}.`, 'warning');
                throw new Error('The anchored formatting baseline changed. Draft a new proposal.');
            }
        }
        const cleanupOps = ops.filter((op) => op.cleanup);
        let cleanupTargets;
        if (cleanupOps.length) {
            if (cleanupOps.length !== 1 || cleanupOps[0].cleanup.emptyParagraphs !== true || !Array.isArray(anchor.cleanupIndexes)) {
                throw new Error('Empty paragraph cleanup is not verified. Draft a new proposal.');
            }
            cleanupTargets = await _collectFormatEmptyParagraphs(context, scopeRange, signal, log);
            if (JSON.stringify(cleanupTargets.indexes) !== JSON.stringify(anchor.cleanupIndexes)) {
                throw new Error('Empty paragraph targets changed. Draft a new proposal.');
            }
        }
        const previousMode = context.document.changeTrackingMode;
        if (Word.ChangeTrackingMode) {
            context.document.changeTrackingMode = appState.config.trackChangesEnabled
                ? Word.ChangeTrackingMode.trackAll : Word.ChangeTrackingMode.off;
        }
        try {
            if (targeted) {
                verifiedResult = await applyVerifiedFormatTargets(context, scopeRange, currentPlan, currentInventory, anchor, { signal });
                appliedRanges = verifiedResult.appliedRanges;
                log(`Formatting read-back verified ${verifiedResult.verifiedParagraphs} paragraph(s); ${verifiedResult.noopParagraphs} already satisfied.`, 'success');
                return;
            }
            // Complete formatting before structural deletion changes ranges.
            for (const op of ops.filter((item) => !item.cleanup)) {
                checkOperationSignal(signal);
                if (op.insert) {
                    anchor.attempted = true;
                    insertedParagraphs += await _applyInsertOp(context, scopeRange, op, log);
                    await context.sync();
                    continue;
                }
                let targets;
                const compiled = currentPlan?.entries.find((entry) => entry.signature === JSON.stringify(op));
                if (compiled) {
                    const paragraphRanges = compiled.ids.map((id) => {
                        const descriptor = currentInventory.descriptors.find((d) => (currentPlan.identityMode === 'local-id' ? d.nativeId : d.id) === id);
                        return scopeRange.paragraphs.items[descriptor.index].getRange('Content');
                    });
                    if (op.match && op.font) {
                        targets = [];
                        for (const range of paragraphRanges) targets.push(...await _resolveFormatTargets(context, range, op));
                    } else targets = paragraphRanges;
                } else targets = await _resolveFormatTargets(context, scopeRange, op);
                for (const target of targets) {
                    checkOperationSignal(signal);
                    let paragraphs;
                    if (op.paragraph) {
                        paragraphs = target.paragraphs;
                        paragraphs.load('items');
                        await context.sync();
                        checkOperationSignal(signal);
                    }
                    anchor.attempted = true;
                    if (op.font) _applyFontOps(target.font, op.font, log);
                    if (paragraphs) {
                        for (const paragraph of paragraphs.items) _applyParagraphOps(paragraph, op.paragraph, log);
                        if (_hasListOps(op.paragraph)) await _applyListOps(context, paragraphs.items, op.paragraph, log);
                    }
                    await context.sync();
                    appliedRanges++;
                }
            }
            if (cleanupTargets) {
                for (const index of [...cleanupTargets.indexes].reverse()) {
                    checkOperationSignal(signal);
                    anchor.attempted = true;
                    cleanupTargets.paragraphs.items[index].delete();
                    await context.sync();
                    deletedParagraphs++;
                }
                log(`Deleted ${deletedParagraphs} verified empty paragraph(s) in ${proposal.scope} scope.`, 'success');
            }
            interrupted = !!signal?.aborted;
        } catch (error) {
            interrupted = error.name === 'AbortError';
            if (!anchor.attempted && !interrupted) throw error;
            partial = !!anchor.attempted;
            log(`Formatting stopped: ${error.message}. Review the document before drafting another proposal.`, 'warning');
        } finally {
            if (Word.ChangeTrackingMode) {
                context.document.changeTrackingMode = previousMode;
                await context.sync();
            }
        }
    });
    if (anchor.attempted) {
        try { await discardFormatProposal(deps, proposal); }
        catch (error) { log(`Formatting anchor cleanup failed: ${error.message}`, 'warning'); }
    }
    return { applied: appliedRanges > 0 || insertedParagraphs > 0 || deletedParagraphs > 0, appliedRanges, insertedParagraphs, deletedParagraphs,
        ...verifiedResult, interrupted, partial: partial || (interrupted && !!anchor.attempted) };
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
        if (op.match.length > 255) throw new Error('Formatting match exceeds Word search limits. Use a paragraph ID instead.');
        const results = scopeRange.search(op.match, { matchCase: true, matchWholeWord: false });
        results.load('items');
        await context.sync();
        return results.items;
    }
    if (op.paragraphStyle) {
        const paragraphs = scopeRange.paragraphs;
        paragraphs.load('items/styleBuiltIn');
        await context.sync();
        const enumValue = _enumValue(Word.BuiltInStyleName || Word.Style, op.paragraphStyle);
        const normalize = (value) => String(value || '').toLowerCase().replace(/[\s_-]+/g, '');
        const wanted = normalize(op.paragraphStyle);
        return paragraphs.items
            .filter((p) => normalize(p.styleBuiltIn) === wanted
                || (enumValue && p.styleBuiltIn === enumValue))
            .map((p) => p.getRange());
    }
    return [scopeRange];
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
                const v = _enumValue(Word.BuiltInStyleName || Word.Style, value);
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
