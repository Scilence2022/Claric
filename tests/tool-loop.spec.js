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

    test('one oversized image attachment is omitted before the next send', async () => {
        const { send, sent } = makeLoop({
            replies: [
                '{"tool":"get_state","args":{}}',
                '{"tool":"finish","args":{"summary":"continued without image"}}',
            ],
        });
        const oversized = 'data:image/png;base64,' + 'A'.repeat(
            TOOL_LOOP_LIMITS.MAX_ATTACHMENT_CHARS + 1);

        const out = await runToolLoop({
            systemPrompt: 'SYS', taskPrompt: 'TASK', tools: TOOLS,
            execute: async () => ({ ok: true, result: { seen: false }, attachments: [{ dataUrl: oversized }] }),
            send,
        });

        expect(out.reason).toBe('finish');
        const observation = sent[1][3];
        expect(typeof observation.content).toBe('string');
        const body = JSON.parse(observation.content);
        expect(body.attachmentWarning).toMatch(/omitted/);
        expect(observation.content.length).toBeLessThan(TOOL_LOOP_LIMITS.MAX_OBSERVATION_CHARS);
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

    // A model stuck re-issuing one call (typically a failing one) otherwise
    // burns the whole step budget and reports 'step-limit', hiding the reason.
    describe('consecutive-repeat guard', () => {
        test('MAX_REPEATED_CALLS identical calls end the loop early with repeat-limit', async () => {
            const { send } = makeLoop({
                replies: ['{"tool":"get_state","args":{}}'],  // same call forever
            });
            const execute = jest.fn(async () => ({ ok: false, error: 'still failing' }));
            const maxSteps = 12;

            const out = await runToolLoop({
                systemPrompt: 'SYS', taskPrompt: 'TASK', tools: TOOLS,
                execute, send, maxSteps,
            });

            expect(out.finished).toBe(false);
            expect(out.reason).toBe('repeat-limit');
            expect(out.summary).toBeNull();
            // Stopping early is the whole point: the step budget is not spent.
            expect(out.steps).toBeLessThan(maxSteps);
            expect(execute.mock.calls.length).toBeLessThan(maxSteps);
        });

        test('argument key order does not defeat the fingerprint', async () => {
            // Same tool, same args, different key order — the fingerprint sorts
            // keys, so these count as repeats.
            const { send } = makeLoop({
                replies: [
                    '{"tool":"set_cell","args":{"row":1,"col":2,"text":"x"}}',
                    '{"tool":"set_cell","args":{"text":"x","row":1,"col":2}}',
                    '{"tool":"set_cell","args":{"col":2,"text":"x","row":1}}',
                    '{"tool":"set_cell","args":{"row":1,"col":2,"text":"x"}}',
                ],
            });

            const out = await runToolLoop({
                systemPrompt: 'SYS', taskPrompt: 'TASK', tools: TOOLS,
                execute: async () => ({ ok: true, result: {} }), send, maxSteps: 12,
            });

            expect(out.reason).toBe('repeat-limit');
            expect(out.steps).toBeLessThan(12);
        });

        test('the repeat is reported through onStep before the loop returns', async () => {
            const { send } = makeLoop({ replies: ['{"tool":"get_state","args":{}}'] });
            const onStep = jest.fn();

            await runToolLoop({
                systemPrompt: 'SYS', taskPrompt: 'TASK', tools: TOOLS,
                execute: async () => ({ ok: true, result: {} }), send,
                maxSteps: 12, onStep,
            });

            const failures = onStep.mock.calls
                .map(([event]) => event)
                .filter((event) => event.ok === false);
            expect(failures.length).toBeGreaterThan(0);
            expect(failures[failures.length - 1].text).toMatch(/identical call/i);
        });

        test('varying arguments never trip the guard', async () => {
            const { send } = makeLoop({
                replies: [
                    '{"tool":"set_cell","args":{"row":1,"col":1,"text":"a"}}',
                    '{"tool":"set_cell","args":{"row":2,"col":1,"text":"b"}}',
                    '{"tool":"set_cell","args":{"row":3,"col":1,"text":"c"}}',
                    '{"tool":"set_cell","args":{"row":4,"col":1,"text":"d"}}',
                    '{"tool":"finish","args":{"summary":"filled four cells"}}',
                ],
            });

            const out = await runToolLoop({
                systemPrompt: 'SYS', taskPrompt: 'TASK', tools: TOOLS,
                execute: async () => ({ ok: true, result: {} }), send, maxSteps: 12,
            });

            expect(out.reason).toBe('finish');
            expect(out.calls).toHaveLength(4);
        });

        test('different nested arguments are not collapsed into one fingerprint', async () => {
            const nestedTool = defineTool({
                name: 'set_style', description: 'Set nested style.', argsExample: {},
            });
            const replies = [1, 2, 3, 4].map((width) => JSON.stringify({
                tool: 'set_style',
                args: { borders: { top: { type: 'solid', width } } },
            }));
            replies.push('{"tool":"finish","args":{"summary":"four styles"}}');
            const { send } = makeLoop({ replies });

            const out = await runToolLoop({
                systemPrompt: 'SYS', taskPrompt: 'TASK', tools: [nestedTool],
                execute: async () => ({ ok: true, result: {} }), send, maxSteps: 8,
            });

            expect(out.reason).toBe('finish');
            expect(out.calls).toHaveLength(4);
        });

        test('a repeat broken by a different call resets the counter', async () => {
            const { send } = makeLoop({
                replies: [
                    '{"tool":"get_state","args":{}}',
                    '{"tool":"get_state","args":{}}',
                    '{"tool":"set_cell","args":{"row":1,"col":1,"text":"x"}}',  // breaks the run
                    '{"tool":"get_state","args":{}}',
                    '{"tool":"get_state","args":{}}',
                    '{"tool":"finish","args":{"summary":"done"}}',
                ],
            });

            const out = await runToolLoop({
                systemPrompt: 'SYS', taskPrompt: 'TASK', tools: TOOLS,
                execute: async () => ({ ok: true, result: {} }), send, maxSteps: 12,
            });

            // Neither run of two reaches MAX_REPEATED_CALLS, so the loop finishes.
            expect(out.reason).toBe('finish');
            expect(out.calls).toHaveLength(5);
        });
    });

    // The transcript is trimmed before each send so the request about to go out
    // stays within budget. System prompt and task are never evicted.
    describe('transcript eviction', () => {
        /** Total characters across a sent messages array. */
        function totalChars(messages) {
            return messages.reduce((sum, m) => sum + (typeof m.content === 'string'
                ? m.content.length
                : 0), 0);
        }

        /** An observation big enough that a few of them bust the budget. */
        function makeHugeExecute() {
            const filler = 'z'.repeat(Math.ceil(TOOL_LOOP_LIMITS.MAX_TRANSCRIPT_CHARS / 3));
            return async () => ({ ok: true, result: { blob: filler } });
        }

        function varyingReplies(count = 8) {
            return Array.from({ length: count }, (_, i) =>
                `{"tool":"set_cell","args":{"row":${i + 1},"col":1,"text":"x"}}`);
        }

        test('every send stays within the transcript budget despite huge observations', async () => {
            const { send, sent } = makeLoop({ replies: varyingReplies() });

            await runToolLoop({
                systemPrompt: 'SYS', taskPrompt: 'TASK', tools: TOOLS,
                execute: makeHugeExecute(), send, maxSteps: 8,
            });

            expect(sent.length).toBeGreaterThan(3);
            for (const messages of sent) {
                expect(totalChars(messages)).toBeLessThanOrEqual(TOOL_LOOP_LIMITS.MAX_TRANSCRIPT_CHARS);
            }
        });

        test('the system prompt and task survive every eviction', async () => {
            const { send, sent } = makeLoop({ replies: varyingReplies() });

            await runToolLoop({
                systemPrompt: 'SYS', taskPrompt: 'TASK', tools: TOOLS,
                execute: makeHugeExecute(), send, maxSteps: 8,
            });

            for (const messages of sent) {
                expect(messages[0]).toEqual({ role: 'system', content: 'SYS' });
                expect(messages[1]).toEqual({ role: 'user', content: 'TASK\n\nBegin.' });
            }
        });

        test('a placeholder tells the model its early history was dropped', async () => {
            const { send, sent } = makeLoop({ replies: varyingReplies() });

            await runToolLoop({
                systemPrompt: 'SYS', taskPrompt: 'TASK', tools: TOOLS,
                execute: makeHugeExecute(), send, maxSteps: 8,
            });

            const lastSent = sent[sent.length - 1];
            const placeholder = lastSent.find((m) => typeof m.content === 'string'
                && m.content.includes('dropped to stay within the context budget'));
            expect(placeholder).toBeDefined();
            expect(placeholder.role).toBe('user');
            // It sits right after the preserved head, where the old turns were.
            expect(lastSent.indexOf(placeholder)).toBe(2);
            // It must not imply the dropped calls never ran.
            const note = JSON.parse(placeholder.content).note;
            expect(note).toMatch(/DID run/);
        });

        test('a small transcript is never evicted and carries no placeholder', async () => {
            const { send, sent } = makeLoop({
                replies: [
                    '{"tool":"get_state","args":{}}',
                    '{"tool":"finish","args":{"summary":"done"}}',
                ],
            });

            await runToolLoop({
                systemPrompt: 'SYS', taskPrompt: 'TASK', tools: TOOLS,
                execute: async () => ({ ok: true, result: { small: true } }), send,
            });

            const lastSent = sent[sent.length - 1];
            expect(lastSent).toHaveLength(4);  // system, task, assistant, observation
            for (const m of lastSent) {
                expect(m.content).not.toContain('context budget');
            }
        });
    });
});
