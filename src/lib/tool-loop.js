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
import { extractJsonObject } from './json-utils.js';

/**
 * Extracts the first JSON object from a model reply, tolerating code fences
 * and surrounding prose (both common with smaller models). Shared
 * implementation (json-utils): balanced-candidate scanning plus
 * string-aware trailing-comma recovery.
 *
 * @param {string} raw
 * @returns {object} Parsed object
 * @throws {Error} When no JSON object can be located or parsed
 * @private
 */
function _extractJsonObject(raw) {
    return extractJsonObject(raw, {
        noObjectMessage: 'reply contains no JSON object',
        parseFailedPrefix: 'JSON parse failed: ',
    });
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
 * @param {{ok?: boolean, result?: *, error?: string, attachmentWarning?: string,
 *   attachments?: Array<{dataUrl: string}>}} observation
 * @returns {{role: string, content: string|Array<object>}}
 * @private
 */
function _observationMessage(observation) {
    const rawAttachments = Array.isArray(observation.attachments) ? observation.attachments : [];
    const attachments = [];
    let attachmentChars = 0;
    let omitted = 0;
    for (const att of rawAttachments) {
        const dataUrl = att && typeof att.dataUrl === 'string' ? att.dataUrl : '';
        if (!dataUrl) continue;
        if (attachmentChars + dataUrl.length > TOOL_LOOP_LIMITS.MAX_ATTACHMENT_CHARS) {
            omitted++;
            continue;
        }
        attachmentChars += dataUrl.length;
        attachments.push(dataUrl);
    }

    const { attachments: _drop, ...body } = observation;
    if (omitted > 0) {
        body.attachmentWarning = `${omitted} image attachment(s) were omitted because they exceeded the request budget.`;
    }
    const text = _boundedObservationJson(body);
    if (attachments.length === 0) {
        return { role: 'user', content: text };
    }

    /** @type {Array<{type: string, text?: string, image_url?: {url: string}}>} */
    const parts = [{ type: 'text', text }];
    for (const dataUrl of attachments) {
        parts.push({ type: 'image_url', image_url: { url: dataUrl } });
    }
    return { role: 'user', content: parts };
}

/**
 * Serializes an observation as valid JSON under the per-message limit. Slicing
 * raw JSON would produce an invalid observation that smaller models cannot
 * recover from, so oversized values are replaced by a valid preview envelope.
 *
 * @param {object} value
 * @returns {string}
 * @private
 */
function _boundedObservationJson(value) {
    let raw;
    try {
        raw = JSON.stringify(value);
    } catch (error) {
        raw = JSON.stringify({ ok: false, error: `Tool observation could not be serialized: ${error.message}` });
    }
    if (raw.length <= TOOL_LOOP_LIMITS.MAX_OBSERVATION_CHARS) return raw;

    const envelope = {
        ok: value && value.ok !== false,
        truncated: true,
        preview: '',
        note: 'Tool observation truncated to stay within the model context budget.',
    };
    const overhead = JSON.stringify(envelope).length;
    envelope.preview = raw.slice(0, Math.max(0, TOOL_LOOP_LIMITS.MAX_OBSERVATION_CHARS - overhead - 8));
    let bounded = JSON.stringify(envelope);
    if (bounded.length > TOOL_LOOP_LIMITS.MAX_OBSERVATION_CHARS) {
        envelope.preview = envelope.preview.slice(0, -(bounded.length - TOOL_LOOP_LIMITS.MAX_OBSERVATION_CHARS));
        bounded = JSON.stringify(envelope);
    }
    return bounded;
}

/**
 * Truncates one immutable initial prompt while retaining a visible suffix.
 *
 * @param {string} value
 * @param {string} suffix
 * @returns {string}
 * @private
 */
function _boundedInitial(value, suffix) {
    const text = String(value || '');
    const limit = TOOL_LOOP_LIMITS.MAX_INITIAL_MESSAGE_CHARS;
    if (text.length <= limit) return text;
    const tail = `\n...[${suffix} truncated to fit context budget]`;
    return text.slice(0, Math.max(0, limit - tail.length)) + tail;
}

/** Text-only character cost (image URLs use a separate wire budget). */
function _messageTextChars(message) {
    const content = message && message.content;
    if (typeof content === 'string') return content.length;
    if (!Array.isArray(content)) return 0;
    return content.reduce((total, part) =>
        total + (part && typeof part.text === 'string' ? part.text.length : 0), 0);
}

/** Actual message-content characters sent over the wire. */
function _messageWireChars(message) {
    const content = message && message.content;
    if (typeof content === 'string') return content.length;
    if (!Array.isArray(content)) return 0;
    return content.reduce((total, part) => {
        if (!part || typeof part !== 'object') return total;
        const textChars = typeof part.text === 'string' ? part.text.length : 0;
        const imageChars = part.image_url && typeof part.image_url.url === 'string'
            ? part.image_url.url.length
            : 0;
        return total + textChars + imageChars;
    }, 0);
}

const BUDGET_NOTE_FRAGMENT = 'dropped to stay within the context budget';
const BUDGET_NOTE_START = '[TOOL-LOOP-HISTORY-NOTE]';
const BUDGET_NOTE_END = '[/TOOL-LOOP-HISTORY-NOTE]';

/** Builds the compact history note inserted after the preserved task. */
function _budgetPlaceholder(totalEvicted) {
    return {
        role: 'user',
        content: JSON.stringify({
            ok: true,
            note: `${BUDGET_NOTE_START}[${totalEvicted} earlier message(s) ${BUDGET_NOTE_FRAGMENT}. `
                + 'Any earlier tool calls in this run DID run; their effects are already staged. '
                + 'Do not repeat them; call list_* tools if you need current state.'
                + BUDGET_NOTE_END,
        }),
    };
}

/**
 * Evicts oldest conversation turns, then complete assistant/observation pairs
 * until both text and wire budgets fit. The system, current task and latest
 * tool exchange remain. The history notice is included in the budget check.
 *
 * @param {Array<{role: string, content: string|Array<object>}>} messages
 * @param {object} state - Current task identity, previous notice and eviction count
 * @private
 */
function _fitMessagesForSend(messages, state) {
    const previousNoteIndex = messages.indexOf(state.placeholder);
    if (previousNoteIndex >= 0) messages.splice(previousNoteIndex, 1);
    state.placeholder = null;

    const totals = () => ({
        text: messages.reduce((sum, m) => sum + _messageTextChars(m), 0),
        wire: messages.reduce((sum, m) => sum + _messageWireChars(m), 0),
    });
    const overBudget = (current) => current.text > TOOL_LOOP_LIMITS.MAX_TRANSCRIPT_CHARS
        || current.wire > TOOL_LOOP_LIMITS.MAX_REQUEST_CHARS;
    let current = totals();
    const evictOldest = () => {
        const taskIndex = messages.indexOf(state.task);
        if (taskIndex > 1) {
            let end = 2;
            while (end < taskIndex && messages[end].role !== 'user') end++;
            state.evicted += messages.splice(1, end - 1).length;
            return true;
        }
        if (messages.length > taskIndex + 3) {
            messages.splice(taskIndex + 1, 2);
            state.evicted += 2;
            return true;
        }
        return false;
    };

    while (overBudget(current) && evictOldest()) {
        current = totals();
    }

    while (state.evicted > 0) {
        const placeholder = _budgetPlaceholder(state.evicted);
        const noteIndex = messages.indexOf(state.task) + 1;
        messages.splice(noteIndex, 0, placeholder);
        current = totals();
        if (!overBudget(current)) {
            state.placeholder = placeholder;
            break;
        }
        messages.splice(noteIndex, 1);
        if (!evictOldest()) break;
    }

    current = totals();
    if (overBudget(current)) {
        throw new Error('Tool-loop request exceeds the configured context budget after trimming.');
    }
}

/** Recursively sorts object keys without filtering nested properties. */
function _stableJsonValue(value) {
    if (Array.isArray(value)) return value.map(_stableJsonValue);
    if (!value || typeof value !== 'object') return value;
    const out = {};
    for (const key of Object.keys(value).sort()) {
        out[key] = _stableJsonValue(value[key]);
    }
    return out;
}

/** Stable fingerprint for consecutive-repeat detection. */
function _callFingerprint(name, args) {
    let argPart;
    try {
        argPart = JSON.stringify(_stableJsonValue(args));
    } catch (_err) {
        argPart = '';
    }
    return `${name}::${argPart}`;
}

/**
 * Runs the tool loop.
 *
 * Message history grows as [system, ...conversationHistory, user(task),
 * assistant(call), user(observation), ...]. `send` receives the array verbatim
 * so callers keep control of transport, config, logging, and signals.
 *
 * Protocol failures (non-JSON reply, unknown tool, non-object args) are NOT
 * thrown — they become error observations so the model can correct itself
 * within the step budget. Abort and transport errors propagate to the
 * caller (turn runners render them as Cancelled/error).
 *
 * Budget guards bound a run beyond `maxSteps`: immutable initial prompts and
 * each observation are capped, complete historical exchanges are evicted until
 * both the text-context and actual wire-size budgets fit, and
 * MAX_REPEATED_CALLS consecutive identical calls end the loop with reason
 * 'repeat-limit' instead of spending every remaining step.
 *
 * @param {object} args
 * @param {string} args.systemPrompt - Protocol prompt (buildToolLoopSystemPrompt)
 * @param {string} args.taskPrompt - Task + initial state for the first user message
 * @param {Array<{role: string, content: string}>} [args.conversationHistory] - Prior user/assistant turns
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
 * @returns {Promise<{finished: boolean, reason: 'finish'|'step-limit'|'repeat-limit',
 *   steps: number, summary: string|null, calls: Array<{tool: string, ok: boolean}>}>}
 *   'repeat-limit' means the model re-issued one identical call
 *   MAX_REPEATED_CALLS times and the loop stopped early; callers treat it like
 *   'step-limit' (no summary, staged work preserved)
 * @throws {DOMException} AbortError when the signal fires
 * @throws {Error} Transport errors from `send`
 */
export async function runToolLoop({
    systemPrompt, taskPrompt, tools, execute, send, conversationHistory = [],
    maxSteps = TOOL_LOOP_LIMITS.MAX_STEPS_DEFAULT, signal, onStep,
}) {
    const known = new Set((Array.isArray(tools) ? tools : []).map((t) => t.name));
    /** @type {Array<{role: string, content: string | Array<object>}>} */
    const messages = [
        { role: 'system', content: _boundedInitial(systemPrompt, 'system prompt') },
        ...(Array.isArray(conversationHistory) ? conversationHistory : [])
            .filter((message) => message && ['user', 'assistant'].includes(message.role)
                && typeof message.content === 'string' && message.content.trim())
            .map(({ role, content }) => ({ role, content })),
        { role: 'user', content: `${_boundedInitial(taskPrompt, 'task prompt')}\n\nBegin.` },
    ];
    const budgetState = { task: messages[messages.length - 1], placeholder: null, evicted: 0 };
    const calls = [];
    let steps = 0;
    let summary = null;
    let lastFingerprint = null;
    let repeatCount = 0;
    const checkAbort = () => {
        if (signal && signal.aborted) {
            throw new DOMException('The operation was aborted.', 'AbortError');
        }
    };

    while (steps < maxSteps) {
        checkAbort();
        // Fit BEFORE sending, including the observation just added. This is a
        // hard invariant over the actual request, not an accounting estimate.
        _fitMessagesForSend(messages, budgetState);
        const rawReply = await send(messages);
        checkAbort();
        const reply = rawReply.length > TOOL_LOOP_LIMITS.MAX_RESPONSE_CHARS
            ? rawReply.slice(0, TOOL_LOOP_LIMITS.MAX_RESPONSE_CHARS)
            : rawReply;
        steps++;
        if (onStep) onStep({ step: steps, call: null, ok: null, text: reply });
        checkAbort();
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
            checkAbort();
            return { finished: true, reason: 'finish', steps, summary, calls };
        }
        if (!known.has(name)) {
            const observation = { ok: false, error: `Unknown tool "${name}". Available: ${[...known, FINISH_TOOL].join(', ')}.` };
            messages.push({ role: 'user', content: JSON.stringify(observation) });
            if (onStep) onStep({ step: steps, call: { tool: name, args: toolArgs }, ok: false, text: observation.error });
            continue;
        }

        // Consecutive-repeat guard: a model stuck on one call (typically a
        // failing one it keeps retrying verbatim) otherwise burns the whole
        // step budget and reports 'step-limit', hiding the real reason.
        const fingerprint = _callFingerprint(name, toolArgs);
        if (fingerprint === lastFingerprint) {
            repeatCount++;
            if (repeatCount >= TOOL_LOOP_LIMITS.MAX_REPEATED_CALLS) {
                const error = `Repeated the identical call to "${name}" ${repeatCount} times in a row. `
                    + 'Stopping to avoid a loop — change the arguments, try a different tool, '
                    + 'or call finish with what you have.';
                if (onStep) onStep({ step: steps, call: { tool: name, args: toolArgs }, ok: false, text: error });
                checkAbort();
                return { finished: false, reason: 'repeat-limit', steps, summary: null, calls };
            }
        } else {
            lastFingerprint = fingerprint;
            repeatCount = 0;
        }

        let observation;
        checkAbort();
        try {
            observation = await execute(name, toolArgs);
            checkAbort();
            if (!observation || typeof observation !== 'object') {
                observation = { ok: false, error: 'Tool returned no observation.' };
            }
        } catch (err) {
            if (err && err.name === 'AbortError') throw err;
            checkAbort();
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

    checkAbort();
    return { finished: false, reason: 'step-limit', steps, summary: null, calls };
}
