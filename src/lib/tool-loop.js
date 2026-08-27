/**
 * Tool Loop
 *
 * L3 of the tool-calling stack: the ReAct-style execution loop. Alternates
 * model replies (one JSON tool call each) with host observations until the
 * model calls `finish` or the step budget runs out.
 *
 * Pure module — the LLM transport (`send`) and the tool dispatch
 * (`execute`) are injected, so the loop is hermetic-testable. No Office.js,
 * no network of its own.
 *
 * @module tool-loop
 */

import { FINISH_TOOL, TOOL_LOOP_LIMITS } from './tool-registry.js';

/**
 * Extracts the first JSON object from a model reply, tolerating code fences
 * and surrounding prose (both common with smaller models). Trailing commas
 * are cleaned like the table-patch parser.
 *
 * @param {string} raw
 * @returns {object} Parsed object
 * @throws {Error} When no JSON object can be located or parsed
 * @private
 */
function _extractJsonObject(raw) {
    const text = String(raw || '');
    const start = text.indexOf('{');
    const end = text.lastIndexOf('}');
    if (start === -1 || end <= start) {
        throw new Error('reply contains no JSON object');
    }
    const cleaned = text.slice(start, end + 1).replace(/,\s*([}\]])/g, '$1');
    try {
        const parsed = JSON.parse(cleaned);
        if (!parsed || typeof parsed !== 'object' || Array.isArray(parsed)) {
            throw new Error('not an object');
        }
        return parsed;
    } catch (err) {
        throw new Error(`JSON parse failed: ${err.message}`);
    }
}

/**
 * Builds the user-message for one observation. Plain observations serialize
 * to a JSON string as before; observations carrying `attachments`
 * ({dataUrl}) become OpenAI-compatible multimodal content arrays — the JSON
 * body (attachments stripped) as the text part plus one image_url part per
 * attachment, so vision-capable backends actually see e.g. a read_image
 * result. Text-only backends reject the array (HTTP 4xx) — degradation is
 * the send wrapper's job, not the loop's.
 *
 * @param {{ok?: boolean, result?: *, error?: string, attachments?: Array<{dataUrl: string}>}} observation
 * @returns {{role: string, content: string|Array<object>}}
 * @private
 */
function _observationMessage(observation) {
    const attachments = Array.isArray(observation.attachments) ? observation.attachments : [];
    if (attachments.length === 0) {
        return { role: 'user', content: JSON.stringify(observation) };
    }
    const { attachments: _drop, ...body } = observation;
    /** @type {Array<{type: string, text?: string, image_url?: {url: string}}>} */
    const parts = [{ type: 'text', text: JSON.stringify(body) }];
    for (const att of attachments) {
        if (att && typeof att.dataUrl === 'string' && att.dataUrl) {
            parts.push({ type: 'image_url', image_url: { url: att.dataUrl } });
        }
    }
    return { role: 'user', content: parts };
}

/**
 * Runs the tool loop.
 *
 * Message history grows as [system, user(task), assistant(call),
 * user(observation), ...] — alternating roles per the chat-completions
 * contract; `send` receives the array verbatim so callers keep control of
 * transport (streaming vs plain), config, logging, and signals.
 *
 * Protocol failures (non-JSON reply, unknown tool, non-object args) are NOT
 * thrown — they become error observations so the model can correct itself
 * within the step budget. Abort and transport errors propagate to the
 * caller (turn runners render them as Cancelled/error).
 *
 * @param {object} args
 * @param {string} args.systemPrompt - Protocol prompt (buildToolLoopSystemPrompt)
 * @param {string} args.taskPrompt - Task + initial state for the first user message
 * @param {Array<{name: string}>} args.tools - Registered tool specs (names validate calls)
 * @param {function(string, object): Promise<{ok: boolean, result?: *, error?: string,
 *   attachments?: Array<{dataUrl: string}>}>} args.execute -
 *   Host-side dispatch: (toolName, args) → observation; attachments ride the
 *   next user message as image inputs (see _observationMessage)
 * @param {function(Array<{role: string, content: string|Array<object>}>): Promise<string>} args.send -
 *   LLM transport: messages → assistant reply text (content may be a
 *   multimodal parts array on attachment-bearing observations)
 * @param {number} [args.maxSteps=TOOL_LOOP_LIMITS.MAX_STEPS_DEFAULT]
 * @param {AbortSignal} [args.signal]
 * @param {function({step: number, call: {tool: string, args: object}|null,
 *   ok: boolean|null, text: string}): void} [args.onStep] - Per-event hook
 *   (model reply, each observation) for UI activity
 * @returns {Promise<{finished: boolean, reason: 'finish'|'step-limit',
 *   steps: number, summary: string|null, calls: Array<{tool: string, ok: boolean}>}>}
 * @throws {DOMException} AbortError when the signal fires
 * @throws {Error} Transport errors from `send`
 */
export async function runToolLoop({
    systemPrompt, taskPrompt, tools, execute, send,
    maxSteps = TOOL_LOOP_LIMITS.MAX_STEPS_DEFAULT, signal, onStep,
}) {
    const known = new Set((Array.isArray(tools) ? tools : []).map((t) => t.name));
    /** @type {Array<{role: string, content: string | Array<object>}>} */
    const messages = [
        { role: 'system', content: systemPrompt },
        { role: 'user', content: `${taskPrompt}\n\nBegin.` },
    ];
    const calls = [];
    let steps = 0;
    let summary = null;

    while (steps < maxSteps) {
        if (signal && signal.aborted) {
            throw new DOMException('The operation was aborted.', 'AbortError');
        }
        const rawReply = await send(messages);
        const reply = rawReply.length > TOOL_LOOP_LIMITS.MAX_RESPONSE_CHARS
            ? rawReply.slice(0, TOOL_LOOP_LIMITS.MAX_RESPONSE_CHARS)
            : rawReply;
        steps++;
        if (onStep) onStep({ step: steps, call: null, ok: null, text: reply });
        messages.push({ role: 'assistant', content: reply });

        let call;
        try {
            call = _extractJsonObject(reply);
        } catch (err) {
            const observation = { ok: false, error: `Protocol violation: ${err.message}. Reply with exactly one JSON object: {"tool":"<name>","args":{...}}` };
            messages.push({ role: 'user', content: JSON.stringify(observation) });
            if (onStep) onStep({ step: steps, call: null, ok: false, text: observation.error });
            continue;
        }

        const name = call.tool;
        const toolArgs = (call.args && typeof call.args === 'object' && !Array.isArray(call.args)) ? call.args : {};
        if (typeof name !== 'string' || name.length === 0) {
            const observation = { ok: false, error: 'Protocol violation: missing "tool" string.' };
            messages.push({ role: 'user', content: JSON.stringify(observation) });
            if (onStep) onStep({ step: steps, call: null, ok: false, text: observation.error });
            continue;
        }
        if (name === FINISH_TOOL) {
            summary = typeof toolArgs.summary === 'string' ? toolArgs.summary : '';
            if (onStep) onStep({ step: steps, call: { tool: name, args: toolArgs }, ok: true, text: summary });
            return { finished: true, reason: 'finish', steps, summary, calls };
        }
        if (!known.has(name)) {
            const observation = { ok: false, error: `Unknown tool "${name}". Available: ${[...known, FINISH_TOOL].join(', ')}.` };
            messages.push({ role: 'user', content: JSON.stringify(observation) });
            if (onStep) onStep({ step: steps, call: { tool: name, args: toolArgs }, ok: false, text: observation.error });
            continue;
        }

        let observation;
        try {
            observation = await execute(name, toolArgs);
            if (!observation || typeof observation !== 'object') {
                observation = { ok: false, error: 'Tool returned no observation.' };
            }
        } catch (err) {
            observation = { ok: false, error: `Tool execution failed: ${err.message}` };
        }
        calls.push({ tool: name, ok: observation.ok !== false });
        messages.push(_observationMessage(observation));
        if (onStep) {
            onStep({
                step: steps,
                call: { tool: name, args: toolArgs },
                ok: observation.ok !== false,
                text: observation.ok === false ? observation.error : JSON.stringify(observation.result ?? {}),
            });
        }
    }

    return { finished: false, reason: 'step-limit', steps, summary: null, calls };
}
