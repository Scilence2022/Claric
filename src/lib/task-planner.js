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

import { extractJsonArray } from './json-utils.js';
import { normalizeCompound } from './task-runtime/task-model.js';
import { validateTaskGraph } from './task-runtime/task-graph.js';

/** Pipeline task types the planner may emit (allowlist for parsePlan). */
const TASK_TYPES = [
    'insert', 'format', 'edit', 'append', 'table', 'illustration', 'qa',
    'image_management', 'table_management', 'document_edit',
];

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
        'You are the task planner of a Microsoft Word add-in. The user instruction below may mix several ' +
        'request types, or may be ambiguous about which pipeline it belongs to. Split it into an ordered ' +
        'list of atomic tasks, one per specialized pipeline.\n\n' +
        'CAPABILITIES (task "type"):\n' +
        '- "document_edit": inspect the article, choose relevant locations, insert or integrate prose, and revise nearby transitions in ONE shared draft with verification. Use for content insertion anywhere except an explicitly requested append, semantic restructuring, source integration, or interdependent prose edits. Keep locating, drafting, transitions, and checking together in one task.\n' +
        '- "insert": add a short NEW structural element that does not exist yet (e.g. an article title, a heading).\n' +
        '- "format": change the FORMATTING of existing text (font, size, color, highlight, paragraph style ' +
        'incl. headings, bulleted/numbered lists, alignment, spacing, indentation) without rewriting it.\n' +
        '- "edit": rewrite, polish, or otherwise change the CONTENT of existing text.\n' +
        '- "append": generate NEW long-form content appended at the document end.\n' +
        '- "table": create a NEW native Word table (with or without generated cell content). Editing the ' +
        'content of an EXISTING table stays on "edit".\n' +
        '- "illustration": design and insert an illustration (SVG artwork).\n' +
        '- "qa": answer a question in chat (no document change).\n' +
        '- "image_management": modify IMAGES anywhere in the document — size, alignment, alt text, ' +
        'hyperlink, delete, replace, or a visible Figure legend/caption. Figure caption work must inspect ' +
        'the selected image pixels and nearby Word context; do not treat selected text as visual evidence. ' +
        'Editing the visual CONTENT of an image (designed replacement) stays on "illustration".\n' +
        '- "table_management": modify an EXISTING table anywhere in the document — cell text, row ops, ' +
        'merges, AND visual styling (table style, borders incl. three-line tables, cell shading/alignment, ' +
        'fonts, header rows, layout, column widths). Creating a NEW table stays on "table".\n\n' +
        'OUTPUT CONTRACT (strict):\n' +
        '- Output ONLY a JSON array. No markdown, no code fences, no explanations, no commentary.\n' +
        '- Each item: { "taskId": "t1", "type": "insert|format|edit|append|table|illustration|qa|image_management|table_management|document_edit", "instruction": "self-contained ' +
        'sub-instruction in the user\'s language", "dependsOn": [] }. Dependencies reference existing taskId values.\n' +
        '- One task per distinct request, in the user\'s original order; each instruction must stand alone ' +
        '(include needed context, e.g. which paragraph to edit).\n' +
        '- If the instruction is really a single request, output a single-task array.\n' +
        '- Cover everything the user asked for; add nothing they did not ask for.\n\n' +
        '- Preservation phrases (keep formatting/headings) are constraints, not separate formatting tasks. A polite question such as "can you insert ..." can be an action request. Explicit document scope overrides an incidental selection.\n' +
        '- Related prose operations must be one document_edit task so later edits see the draft. Other write pipelines produce unapplied proposals; do not assume they are already in Word. A dependent qa task may discuss their proposed results.\n' +
        `CONTEXT: the user currently has ${selectionLabel}. ${selectionKind}\n` +
        (facts.hasMultiCellTableRegion ? 'The selection covers a multi-cell table region.\n' : '') +
        '\nUSER INSTRUCTION:\n' + (instruction || '').trim()
    );
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
