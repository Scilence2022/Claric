/**
 * Tool Registry
 *
 * L1 of the tool-calling stack: normalized tool specifications plus the
 * system prompt that teaches an LLM the (prompt-simulated) tool protocol.
 *
 * The project targets arbitrary OpenAI-compatible backends (vLLM, Ollama,
 * custom endpoints), so native function-calling cannot be assumed — the
 * loop is ReAct-style over plain messages: the model replies with ONE
 * JSON tool call per turn, the host executes it and replies with a JSON
 * observation, until the model calls `finish` or the step budget runs out.
 *
 * Pure module — no Office.js, no network. Hermetic-testable.
 *
 * @module tool-registry
 */

/**
 * Hard limits for the tool loop. Frozen — STYLE.md "Enums for Fixed Values".
 */
export const TOOL_LOOP_LIMITS = Object.freeze({
    /** Default max model↔host round trips per loop. */
    MAX_STEPS_DEFAULT: 12,
    /** Max characters accepted for one model reply (guard against rambles). */
    MAX_RESPONSE_CHARS: 64 * 1024,
});

/** The reserved tool name that ends a loop. */
export const FINISH_TOOL = 'finish';

/**
 * Normalizes one tool specification.
 *
 * @param {object} spec
 * @param {string} spec.name - Tool name (snake_case, called by the model)
 * @param {string} spec.description - One-paragraph contract for the model,
 *   including constraints and error behavior
 * @param {object} [spec.argsExample] - Example args object; rendered into the
 *   prompt as the canonical call shape (examples teach small models better
 *   than JSON schemas)
 * @returns {{name: string, description: string, argsExample: object}} Frozen spec
 */
export function defineTool({ name, description, argsExample = {} }) {
    if (!name || typeof name !== 'string' || !/^[a-z][a-z0-9_]*$/.test(name)) {
        throw new Error(`Invalid tool name: ${name}`);
    }
    return Object.freeze({
        name,
        description: String(description || ''),
        argsExample: Object.freeze({ ...argsExample }),
    });
}

/**
 * Builds the tool-loop system prompt: the strict one-call-per-turn JSON
 * protocol, the observation contract, and the per-tool reference rendered
 * from the specs.
 *
 * @param {Array<{name: string, description: string, argsExample: object}>} tools
 * @param {object} [options]
 * @param {number} [options.maxSteps=TOOL_LOOP_LIMITS.MAX_STEPS_DEFAULT]
 * @returns {string}
 */
export function buildToolLoopSystemPrompt(tools, { maxSteps = TOOL_LOOP_LIMITS.MAX_STEPS_DEFAULT } = {}) {
    const list = (Array.isArray(tools) ? tools : []).map((t) => defineTool(t));
    const toolBlocks = list.map((t) =>
        `### ${t.name}\n${t.description}\nExample call: {"tool":"${t.name}","args":${JSON.stringify(t.argsExample)}}`
    ).join('\n\n');

    return (
        'You are an agent that operates document-editing TOOLS in Microsoft Word through a host program.\n\n' +
        'PROTOCOL (strict):\n' +
        `- Each reply must be EXACTLY ONE JSON object and nothing else — no prose, no markdown fences:\n` +
        `  {"tool":"<name>","args":{...}}  to call a tool\n` +
        `  {"tool":"${FINISH_TOOL}","args":{"summary":"one-line report of what you did"}}  when the task is complete\n` +
        '- The host executes your call and replies with an observation:\n' +
        '  {"ok":true,"result":...} on success, {"ok":false,"error":"..."} on failure.\n' +
        '- When a call fails, read the error, correct the arguments, and call again. Do not give up after one failure.\n' +
        '- Call exactly ONE tool per reply. Plan multi-step work step by step and verify results with tool output as you go.\n' +
        `- You have at most ${maxSteps} tool calls. Finish before the limit; do not repeat identical calls.\n` +
        '- Never claim a change you did not make through a tool call.\n\n' +
        `TOOLS:\n${toolBlocks}\n\n` +
        `### ${FINISH_TOOL}\nEnds the loop. args: {"summary":"..."} — required.`
    );
}
