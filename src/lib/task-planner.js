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

/** Pipeline task types the planner may emit (allowlist for parsePlan). */
const TASK_TYPES = [
    'insert', 'format', 'edit', 'append', 'table', 'illustration', 'qa',
    'image_management', 'table_management',
];

/** A compound instruction decomposes into at most this many tasks. */
const MAX_TASKS = 6;

/** Per-task instruction length cap — planner output should be terse. */
const MAX_TASK_INSTRUCTION_CHARS = 500;

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
        '- Each item: { "type": "insert|format|edit|append|table|illustration|qa|image_management|table_management", "instruction": "self-contained ' +
        'sub-instruction in the user\'s language" }.\n' +
        '- One task per distinct request, in the user\'s original order; each instruction must stand alone ' +
        '(include needed context, e.g. which paragraph to edit).\n' +
        '- If the instruction is really a single request, output a single-task array.\n' +
        '- Cover everything the user asked for; add nothing they did not ask for.\n\n' +
        `CONTEXT: the user currently has ${selectionLabel}. ${selectionKind}\n` +
        (facts.hasMultiCellTableRegion ? 'The selection covers a multi-cell table region.\n' : '') +
        '\nUSER INSTRUCTION:\n' + (instruction || '').trim()
    );
}

/**
 * Parses and validates the planner's JSON task list. Tolerates code fences
 * and surrounding prose; drops malformed entries and unknown types, caps
 * the list at MAX_TASKS. Returns null when nothing usable remains — the
 * caller then falls back to single-intent routing.
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

    const tasks = [];
    for (const entry of parsed) {
        if (!entry || typeof entry !== 'object' || Array.isArray(entry)) continue;
        if (!TASK_TYPES.includes(entry.type)) {
            log(`Task planner: dropped a task with unknown type "${entry.type}"`, 'warning');
            continue;
        }
        const instruction = String(entry.instruction === undefined || entry.instruction === null ? '' : entry.instruction).trim();
        if (!instruction) {
            log(`Task planner: dropped a "${entry.type}" task with an empty instruction`, 'warning');
            continue;
        }
        if (instruction.length > MAX_TASK_INSTRUCTION_CHARS) {
            log(`Task planner: task instruction truncated to ${MAX_TASK_INSTRUCTION_CHARS} chars`, 'warning');
        }
        tasks.push({ type: entry.type, instruction: instruction.slice(0, MAX_TASK_INSTRUCTION_CHARS) });
    }

    if (tasks.length === 0) {
        log('Task planner: no valid tasks in the model response', 'warning');
        return null;
    }
    if (tasks.length > MAX_TASKS) {
        log(`Task planner: plan capped at ${MAX_TASKS} tasks (got ${tasks.length})`, 'warning');
        return tasks.slice(0, MAX_TASKS);
    }
    return tasks;
}
