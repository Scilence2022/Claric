/**
 * Specs for the tool-calling stack's pure layers:
 *   - tool-registry.js: defineTool + buildToolLoopSystemPrompt
 *   - tool-loop.js: runToolLoop protocol behavior
 *
 * The loop's send/execute are injected fakes — no LLM, no Word.
 */

const {
    defineTool, buildToolLoopSystemPrompt, FINISH_TOOL, TOOL_LOOP_LIMITS,
} = require('../src/lib/tool-registry.js');
const { runToolLoop } = require('../src/lib/tool-loop.js');

const TOOLS = [
    defineTool({ name: 'set_cell', description: 'Set one cell.', argsExample: { row: 1, col: 1, text: 'x' } }),
    defineTool({ name: 'get_state', description: 'Read the grid.' }),
];

describe('defineTool', () => {
    test('normalizes and freezes a spec', () => {
        const spec = defineTool({ name: 'do_thing', description: 'Does a thing.', argsExample: { a: 1 } });
        expect(spec.name).toBe('do_thing');
        expect(Object.isFrozen(spec)).toBe(true);
    });

    test('rejects names that are not snake_case identifiers', () => {
        expect(() => defineTool({ name: 'Do Thing' })).toThrow(/Invalid tool name/);
        expect(() => defineTool({ name: '' })).toThrow(/Invalid tool name/);
        expect(() => defineTool({ name: '1bad' })).toThrow(/Invalid tool name/);
    });
});

describe('buildToolLoopSystemPrompt', () => {
    test('embeds the protocol, tools, examples, and step budget', () => {
        const prompt = buildToolLoopSystemPrompt(TOOLS, { maxSteps: 7 });
        expect(prompt).toContain('EXACTLY ONE JSON object');
        expect(prompt).toContain(`{"tool":"${FINISH_TOOL}","args":{"summary":"`);
        expect(prompt).toContain('### set_cell');
        expect(prompt).toContain('### get_state');
        expect(prompt).toContain('{"tool":"set_cell","args":{"row":1,"col":1,"text":"x"}}');
        expect(prompt).toContain('at most 7 tool calls');
    });

    test('defaults to the standard step budget', () => {
        expect(buildToolLoopSystemPrompt(TOOLS)).toContain(
            `at most ${TOOL_LOOP_LIMITS.MAX_STEPS_DEFAULT} tool calls`);
    });
});

describe('runToolLoop', () => {
    function makeLoop({ replies, execute } = {}) {
        const sent = [];
        const send = async (messages) => {
            sent.push(messages.map((m) => ({ role: m.role, content: m.content })));
            const reply = replies[Math.min(sent.length - 1, replies.length - 1)];
            if (reply instanceof Error) throw reply;
            return reply;
        };
        return { send, sent };
    }

    test('happy path: one tool call, then finish; observations appended', async () => {
        const { send, sent } = makeLoop({
            replies: [
                '{"tool":"set_cell","args":{"row":1,"col":1,"text":"x"}}',
                '{"tool":"finish","args":{"summary":"set one cell"}}',
            ],
        });
        const execute = jest.fn(async () => ({ ok: true, result: { set: 'R1C1' } }));

        const out = await runToolLoop({
            systemPrompt: 'SYS', taskPrompt: 'TASK', tools: TOOLS, execute, send,
        });

        expect(out.finished).toBe(true);
        expect(out.reason).toBe('finish');
        expect(out.steps).toBe(2);
        expect(out.summary).toBe('set one cell');
        expect(out.calls).toEqual([{ tool: 'set_cell', ok: true }]);
        expect(execute).toHaveBeenCalledWith('set_cell', { row: 1, col: 1, text: 'x' });

        // History alternates and carries the observation between the calls.
        expect(sent[1]).toEqual([
            { role: 'system', content: 'SYS' },
            { role: 'user', content: 'TASK\n\nBegin.' },
            { role: 'assistant', content: '{"tool":"set_cell","args":{"row":1,"col":1,"text":"x"}}' },
            { role: 'user', content: JSON.stringify({ ok: true, result: { set: 'R1C1' } }) },
        ]);
    });

    test('non-JSON reply becomes an error observation and the model recovers', async () => {
        const { send, sent } = makeLoop({
            replies: [
                'Sure! I will set the cell now.',  // protocol violation
                '{"tool":"set_cell","args":{"row":1,"col":1,"text":"x"}}',
                '{"tool":"finish","args":{"summary":"done"}}',
            ],
        });
        const out = await runToolLoop({
            systemPrompt: 'SYS', taskPrompt: 'TASK', tools: TOOLS,
            execute: async () => ({ ok: true, result: {} }), send,
        });

        expect(out.finished).toBe(true);
        expect(out.steps).toBe(3);
        const observation = JSON.parse(sent[1][3].content);
        expect(observation.ok).toBe(false);
        expect(observation.error).toMatch(/Protocol violation/);
    });

    test('unknown tool and bad args shape become error observations', async () => {
        const { send, sent } = makeLoop({
            replies: [
                '{"tool":"nope","args":{}}',
                '{"tool":"set_cell","args":"not an object"}',
                '{"tool":"finish","args":{"summary":"ok"}}',
            ],
        });
        const out = await runToolLoop({
            systemPrompt: 'SYS', taskPrompt: 'TASK', tools: TOOLS,
            execute: async () => ({ ok: true, result: {} }), send,
        });

        expect(out.finished).toBe(true);
        const unknownObs = JSON.parse(sent[1][3].content);
        expect(unknownObs.error).toMatch(/Unknown tool "nope"/);
        // Non-object args are coerced to {} — execute still ran.
        expect(out.calls).toEqual([{ tool: 'set_cell', ok: true }]);
    });

    test('execute throwing becomes an error observation, not a loop crash', async () => {
        const { send, sent } = makeLoop({
            replies: [
                '{"tool":"set_cell","args":{}}',
                '{"tool":"finish","args":{"summary":"recovered"}}',
            ],
        });
        const out = await runToolLoop({
            systemPrompt: 'SYS', taskPrompt: 'TASK', tools: TOOLS,
            execute: async () => { throw new Error('boom'); }, send,
        });

        expect(out.finished).toBe(true);
        expect(out.calls).toEqual([{ tool: 'set_cell', ok: false }]);
        const obs = JSON.parse(sent[1][3].content);
        expect(obs.error).toMatch(/boom/);
    });

    test('step limit stops the loop unfinished', async () => {
        const { send } = makeLoop({
            replies: ['{"tool":"get_state","args":{}}'],  // repeated forever
        });
        const onStep = jest.fn();
        const out = await runToolLoop({
            systemPrompt: 'SYS', taskPrompt: 'TASK', tools: TOOLS,
            execute: async () => ({ ok: true, result: {} }), send,
            maxSteps: 3, onStep,
        });

        expect(out.finished).toBe(false);
        expect(out.reason).toBe('step-limit');
        expect(out.steps).toBe(3);
        expect(onStep).toHaveBeenCalledTimes(6); // reply + observation per step
    });

    test('pre-aborted signal rejects with AbortError', async () => {
        const controller = new AbortController();
        controller.abort();
        const { send } = makeLoop({ replies: ['{}'] });
        await expect(runToolLoop({
            systemPrompt: 'SYS', taskPrompt: 'TASK', tools: TOOLS,
            execute: async () => ({ ok: true }), send, signal: controller.signal,
        })).rejects.toMatchObject({ name: 'AbortError' });
    });

    test('transport errors propagate to the caller', async () => {
        const { send } = makeLoop({ replies: [new Error('HTTP 500: boom')] });
        await expect(runToolLoop({
            systemPrompt: 'SYS', taskPrompt: 'TASK', tools: TOOLS,
            execute: async () => ({ ok: true }), send,
        })).rejects.toThrow(/HTTP 500/);
    });

    test('markdown-fenced tool calls parse (small-model tolerance)', async () => {
        const { send } = makeLoop({
            replies: [
                '```json\n{"tool":"set_cell","args":{"row":2,"col":1,"text":"y"}}\n```',
                '{"tool":"finish","args":{"summary":"ok"}}',
            ],
        });
        const out = await runToolLoop({
            systemPrompt: 'SYS', taskPrompt: 'TASK', tools: TOOLS,
            execute: async () => ({ ok: true, result: {} }), send,
        });
        expect(out.calls).toEqual([{ tool: 'set_cell', ok: true }]);
    });

    test('observations with attachments become multimodal content arrays', async () => {
        const { send, sent } = makeLoop({
            replies: [
                '{"tool":"get_state","args":{}}',
                '{"tool":"finish","args":{"summary":"seen"}}',
            ],
        });
        const dataUrl = 'data:image/png;base64,iVBORw0KGgo=';
        const out = await runToolLoop({
            systemPrompt: 'SYS', taskPrompt: 'TASK', tools: TOOLS,
            execute: async () => ({
                ok: true,
                result: { seen: true },
                attachments: [{ dataUrl }, { dataUrl: '' }],  // empty entries dropped
            }),
            send,
        });

        expect(out.finished).toBe(true);
        const observation = sent[1][3];
        expect(Array.isArray(observation.content)).toBe(true);
        // Text part carries the JSON body WITHOUT the attachment bytes.
        const textPart = observation.content[0];
        expect(textPart.type).toBe('text');
        expect(JSON.parse(textPart.text)).toEqual({ ok: true, result: { seen: true } });
        // One image_url part per valid attachment.
        expect(observation.content[1]).toEqual({ type: 'image_url', image_url: { url: dataUrl } });
        expect(observation.content).toHaveLength(2);
    });

    test('plain observations stay JSON strings (no behavior change)', async () => {
        const { send, sent } = makeLoop({
            replies: [
                '{"tool":"get_state","args":{}}',
                '{"tool":"finish","args":{"summary":"ok"}}',
            ],
        });
        await runToolLoop({
            systemPrompt: 'SYS', taskPrompt: 'TASK', tools: TOOLS,
            execute: async () => ({ ok: true, result: {}, attachments: [] }),
            send,
        });
        expect(typeof sent[1][3].content).toBe('string');
        expect(JSON.parse(sent[1][3].content).ok).toBe(true);
    });
});
