/**
 * Task Planner Module
 *
 * Compound instructions ("增加标题，并深度润色修改") hit several intent
 * families at once. Routed to any single pipeline, the other parts are
 * dropped — or worse, refused outright (the format contract answers []
 * whenever rewriting is requested). The planner decomposes such an
 * instruction into an ordered list of atomic tasks, one per specialized
 * pipeline:
 *
 *   [
 *     { "type": "insert",       "instruction": "为文章拟一个标题，插入文首并套用 Title 样式" },
 *     { "type": "edit",         "instruction": "深度润色全文" },
 *     { "type": "illustration", "instruction": "为文章配一张插图" }
 *   ]
 *
 * The planner only classifies intent — it never sees the document text, so
 * the planning call is cheap. conversation.js executes the tasks through
 * the existing per-pipeline turn runners.
 *
 * Pure module — no DOM, no Word API. Safe to import under Jest/node.
 *
 * @module task-planner
 */

import { extractJsonArray, extractJsonObject } from './json-utils.js';
import { normalizeCompound } from './task-runtime/task-model.js';
import { validateTaskGraph } from './task-runtime/task-graph.js';
import { CAPABILITY_BY_TYPE, capabilityPrompt, capabilityTypeList } from './capability-catalog.js';

/** Pipeline task types the planner may emit (allowlist for parsePlan). */
const TASK_TYPES = capabilityTypeList();

/** A compound instruction decomposes into at most this many tasks. */
const MAX_TASKS = 6;

/** Per-task instruction length cap — planner output should be terse. */
const MAX_TASK_INSTRUCTION_CHARS = 4000;

/**
 * Builds the LLM prompt that decomposes a compound instruction into an
 * ordered task list.
 *
 * @param {string} instruction - The user's compound instruction
 * @param {boolean|object} hasSelection - Whether the document has a non-empty selection,
 *   or selection facts ({ hasSelection, hasImageSelection, hasTextSelection,
 *   hasMultiCellTableRegion })
 * @returns {string}
 */
export function buildPlanPrompt(instruction, hasSelection) {
    const facts = typeof hasSelection === 'object' && hasSelection !== null
        ? hasSelection
        : { hasSelection: !!hasSelection };
    const selectionLabel = facts.hasSelection
        ? (facts.hasImageSelection ? 'an image selection in the document' : 'a text selection in the document')
        : 'NO text selection';
    const selectionKind = facts.hasImageSelection
        ? (facts.hasTextSelection ? 'The selection contains image(s) and text.' : 'The selection contains image(s) only.')
        : (facts.hasTextSelection ? 'The selection contains text only.' : 'No image or text selection is active.');
    return (
        'Plan the complete user outcome for a Microsoft Word add-in. Treat the request as open-ended: the capabilities below are executable, not an exhaustive list of possible requests. Do not map an unsupported action to an approximate capability.\n\n' +
        'CAPABILITIES (task "type"):\n' + capabilityPrompt() + '\n\n' +
        'OUTPUT CONTRACT (strict): Return ONLY one JSON object: {"requirements":[{"id":"r1","kind":"action|constraint","outcome":"document|answer" for actions,"text":"exact user need"}],"tasks":[{"taskId":"t1","type":"' + TASK_TYPES.join('|') + '","instruction":"self-contained subtask in user language","covers":["r1"],"dependsOn":[]}],"unsupported":[{"requirementId":"r2","reason":"specific missing Word action"}]}.\n' +
        '- Identify every requested action and preservation/scope/source constraint separately. Each action must be covered by exactly one executable task or named in unsupported. Constraints may cover several tasks. No missing or invented requirements.\n' +
        '- Use one document_edit task for interdependent prose insertion, transition edits, and bold/italic formatting of its NEW paragraphs. Other pipelines cannot read another unapplied proposal. Mark write-after-write dependencies with dependsOn; do not pretend they share a draft.\n' +
        '- Ask for no unavailable ability: e.g. deleting footnotes or editing reference fields is unsupported. Existing table cell edits use table_management. A question uses qa only when the user wants an answer in chat.\n' +
        '- At most 6 tasks and 16 requirements. Preserve scope and constraints in each task instruction. Dependencies use taskId values and must be acyclic. If everything is unsupported, tasks may be empty.\n' +
        `CONTEXT: the user currently has ${selectionLabel}. ${selectionKind}\n` +
        (facts.hasMultiCellTableRegion ? 'The selection covers a multi-cell table region.\n' : '') +
        '\nUSER INSTRUCTION:\n' + (instruction || '').trim()
    );
}

/** Parse the auditable plan used for open-ended and compound requests. */
export function parseCapabilityPlan(raw, log = (_message, _level) => {}) {
    let value;
    try { value = extractJsonObject(raw); } catch (error) { log(`Task planner: ${error.message}`, 'warning'); return null; }
    if (!value || typeof value !== 'object' || Array.isArray(value)
        || !Array.isArray(value.requirements) || !value.requirements.length || value.requirements.length > 16
        || !Array.isArray(value.tasks) || value.tasks.length > MAX_TASKS
        || !Array.isArray(value.unsupported) || value.unsupported.length > 16) return null;
    const ids = new Set();
    const requirements = [];
    for (const item of value.requirements) {
        if (!item || typeof item !== 'object' || typeof item.id !== 'string'
            || !/^[A-Za-z][A-Za-z0-9_-]{0,39}$/.test(item.id) || ids.has(item.id)
            || !['action', 'constraint'].includes(item.kind) || typeof item.text !== 'string'
            || !item.text.trim() || item.text.length > 1000
            || (item.kind === 'action' && !['document', 'answer'].includes(item.outcome))) return null;
        ids.add(item.id);
        requirements.push({ id: item.id, kind: item.kind, ...(item.kind === 'action' ? { outcome: item.outcome } : {}), text: item.text.trim() });
    }
    const unsupported = [];
    const unsupportedIds = new Set();
    for (const item of value.unsupported) {
        if (!item || !ids.has(item.requirementId) || unsupportedIds.has(item.requirementId)
            || typeof item.reason !== 'string' || !item.reason.trim() || item.reason.length > 1000) return null;
        unsupportedIds.add(item.requirementId);
        unsupported.push({ requirementId: item.requirementId, reason: item.reason.trim() });
    }
    if (value.tasks.some((item) => !item || typeof item.taskId !== 'string')) return null;
    /** @type {Array<any>|null} */
    const tasks = parsePlan(JSON.stringify(value.tasks), log) || (value.tasks.length === 0 ? [] : null);
    if (!tasks) return null;
    const covered = new Map(requirements.map((item) => [item.id, []]));
    for (const task of tasks) {
        const source = value.tasks.find((item) => item.taskId === task.taskId);
        if (!source || !Array.isArray(source.covers) || !source.covers.length
            || source.covers.some((id) => !ids.has(id) || unsupportedIds.has(id))) return null;
        task.covers = [...new Set(source.covers)];
        if (task.covers.length !== source.covers.length) return null;
        const capability = CAPABILITY_BY_TYPE.get(task.type);
        for (const id of task.covers) {
            const requirement = requirements.find((item) => item.id === id);
            if (requirement.kind === 'action' && requirement.outcome !== capability.effect) return null;
            covered.get(id).push(task.taskId);
        }
    }
    if (requirements.some((item) => !unsupportedIds.has(item.id)
        && (covered.get(item.id).length === 0 || (item.kind === 'action' && covered.get(item.id).length !== 1)))) return null;
    if (!tasks.length && !unsupported.length) return null;
    return { requirements, tasks, unsupported };
}

/** A separate model checks whether the plan represents the original request. */
export function parsePlanReview(raw, requirementIds) {
    let value;
    try { value = extractJsonObject(raw); } catch (_error) { return null; }
    if (value?.complete !== true || value.unsupportedAccurate !== true
        || !Array.isArray(value.checks) || value.checks.length !== requirementIds.length
        || !Array.isArray(value.missing) || value.missing.length
        || !Array.isArray(value.invented) || value.invented.length) return null;
    const checks = new Map(value.checks.map((item) => [item?.requirementId, item]));
    if (checks.size !== requirementIds.length || requirementIds.some((id) => checks.get(id)?.represented !== true)) return null;
    return { complete: true, summary: typeof value.summary === 'string' ? value.summary.slice(0, 1000) : '' };
}

/**
 * Re-applies {@link normalizeCompound} to an already-parsed task list so it
 * carries the canonical per-task fields (`taskId`, `attemptId`, normalized
 * `type` / `instruction`, de-duplicated `dependsOn` / `resources`,
 * default `state`). Used after {@link parsePlan} hands the model output to
 * the executor — the planner's own output uses raw string `instruction`
 * values which the task runtime normalizes on first use.
 *
 * @param {Array<object>} tasks - Tasks produced by parsePlan
 * @param {string} [graphId] - Optional graph id; auto-generated when absent
 * @returns {Array<object>} Normalized task list
 */
export function normalizePlan(tasks, graphId) {
    return normalizeCompound({ graphId, tasks }).tasks;
}

/**
 * Parses and validates the planner's JSON task list. Tolerates code fences
 * and surrounding prose. Rejects the whole plan on invalid entries or
 * limits instead of silently dropping user requirements.
 *
 * @param {string} raw - Raw model output
 * @param {function} [log] - Logging callback
 * @returns {Array<{ type: string, instruction: string }> | null}
 */
export function parsePlan(raw, log = () => {}) {
    if (!raw) return null;

    const { value: parsed, error } = extractJsonArray(raw);
    if (error) {
        log(`Task planner: ${error}`, 'warning');
        return null;
    }
    if (!Array.isArray(parsed)) return null;
    if (parsed.length > MAX_TASKS) {
        log(`Task planner: plan exceeds ${MAX_TASKS} tasks; rejected without truncation`, 'warning');
        return null;
    }

    const tasks = [];
    for (const entry of parsed) {
        if (!entry || typeof entry !== 'object' || Array.isArray(entry)) {
            log('Task planner: invalid task entry', 'warning');
            return null;
        }
        if (!TASK_TYPES.includes(entry.type)) {
            log(`Task planner: unknown type "${entry.type}"; plan rejected`, 'warning');
            return null;
        }
        const instruction = typeof entry.instruction === 'string' ? entry.instruction.trim() : '';
        if (!instruction) {
            log(`Task planner: "${entry.type}" task has an empty instruction`, 'warning');
            return null;
        }
        if (instruction.length > MAX_TASK_INSTRUCTION_CHARS) {
            log(`Task planner: instruction exceeds ${MAX_TASK_INSTRUCTION_CHARS} chars; plan rejected`, 'warning');
            return null;
        }
        const task = { type: entry.type, instruction };
        const id = entry.taskId ?? entry.id;
        if (id !== undefined) {
            if (typeof id !== 'string' || !/^[A-Za-z0-9_-]{1,80}$/.test(id)) return null;
            task.taskId = id;
        }
        for (const field of ['dependsOn', 'resources', 'inputRefs']) {
            if (entry[field] === undefined) continue;
            if (!Array.isArray(entry[field]) || entry[field].length > 24
                || entry[field].some((value) => typeof value !== 'string' || !value.trim() || value.length > 200)) return null;
            task[field] = [...new Set(entry[field])];
        }
        tasks.push(task);
    }

    if (tasks.length === 0) {
        log('Task planner: no valid tasks in the model response', 'warning');
        return null;
    }
    const checked = validateTaskGraph({ tasks });
    if (!checked.valid) {
        log(`Task planner: ${checked.errors.join('; ')}`, 'warning');
        return null;
    }
    return tasks;
}
