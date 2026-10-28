/** @jest-environment jsdom */
jest.mock('../src/lib/llm-client.js', () => ({ sendPrompt: jest.fn(), sendPromptStream: jest.fn(), sendMessages: jest.fn(), sendMessagesStream: jest.fn() }));
const { sendPrompt, sendMessages, sendMessagesStream } = require('../src/lib/llm-client.js');
const { planDocumentTasks } = require('../src/taskpane/word-actions.js');
const { COMPOSITE_FORMAT_REQUEST, compositeFormatPlan, approvingFormatReview } = require('./fixtures/composite-format-plan.js');

const deps = () => ({ appState: { config: { backend: 'mock', providers: { mock: { model: 'm' } } } }, log: jest.fn() });
const planReply = jest.fn();
const reviewReply = jest.fn();
beforeEach(() => {
    jest.resetAllMocks();
    sendMessages.mockImplementation((config, messages, ...args) =>
        (messages[0].content.startsWith('Independently audit') ? reviewReply : planReply)(config, messages, ...args));
    sendMessagesStream.mockImplementation(async (config, messages, handlers, ...args) => {
        const content = await planReply(config, messages, ...args);
        handlers.onContent?.(content);
        return { content };
    });
});

test('the reported compound request passes with shared coverage and explicit document scope', async () => {
    planReply.mockResolvedValue(JSON.stringify(compositeFormatPlan()));
    reviewReply.mockResolvedValue(JSON.stringify(approvingFormatReview()));
    const d = deps();
    const result = await planDocumentTasks(d, { instruction: COMPOSITE_FORMAT_REQUEST, hasSelection: true, hasTextSelection: true });
    expect(result.tasks).toHaveLength(3);
    expect(result.tasks.map(({ type, scope, covers }) => ({ type, scope, covers }))).toEqual(compositeFormatPlan().tasks
        .map(({ type, scope, covers }) => ({ type, scope, covers })));
    expect(planReply).toHaveBeenCalledTimes(1);
    const reviewMessages = reviewReply.mock.calls[0][1];
    expect(reviewMessages[0].role).toBe('system');
    expect(reviewMessages[0].content).toContain('native table/row borders');
    expect(reviewMessages[0].content).toContain('redundant spaces in body text');
    expect(JSON.parse(reviewMessages[1].content)).toMatchObject({ originalRequest: COMPOSITE_FORMAT_REQUEST,
        selectionFacts: { hasSelection: true }, plan: { tasks: compositeFormatPlan().tasks } });
});

test('repair receives actual rejected-review findings and can correct an invented blank-line operation', async () => {
    const first = compositeFormatPlan();
    first.tasks[0].instruction += '同时删除空段落。';
    planReply.mockResolvedValueOnce(JSON.stringify(first)).mockResolvedValueOnce(JSON.stringify(compositeFormatPlan()));
    const negative = { ...approvingFormatReview(), complete: false, invented: ['Do not delete blank paragraphs; only spaces were requested'],
        missing: ['Preserve meaningful spaces in equations'], summary: 'Repair cleanup scope without adding paragraph deletion' };
    reviewReply.mockResolvedValueOnce(JSON.stringify(negative)).mockResolvedValueOnce(JSON.stringify(approvingFormatReview()));
    const d = deps();
    const result = await planDocumentTasks(d, { instruction: COMPOSITE_FORMAT_REQUEST });
    expect(result.tasks).toHaveLength(3);
    const repairPrompt = JSON.stringify(planReply.mock.calls[1][1]);
    expect(repairPrompt).toContain(negative.summary);
    expect(repairPrompt).toContain(negative.missing[0]);
    expect(repairPrompt).toContain(negative.invented[0]);
    expect(result.tasks[0].instruction).not.toContain('删除空段落');
    expect(d.log).toHaveBeenCalledWith(expect.stringContaining(negative.summary), 'warning');
    expect(reviewReply).toHaveBeenCalledTimes(2);
});

test('malformed coverage gives the repair a concrete contract reason instead of a blind retry', async () => {
    const first = compositeFormatPlan(); first.tasks[2].covers = ['r3', 'r5'];
    planReply.mockResolvedValueOnce(JSON.stringify(first)).mockResolvedValueOnce(JSON.stringify(compositeFormatPlan()));
    reviewReply.mockResolvedValue(JSON.stringify(approvingFormatReview()));
    expect((await planDocumentTasks(deps(), { instruction: COMPOSITE_FORMAT_REQUEST })).tasks).toHaveLength(3);
    expect(JSON.stringify(planReply.mock.calls[1][1])).toContain('a requirement has no task or unsupported explanation');
    expect(reviewReply).toHaveBeenCalledTimes(1);
});

test('persistent rejection stops execution and exposes the final concrete reason', async () => {
    planReply.mockResolvedValue(JSON.stringify(compositeFormatPlan()));
    reviewReply.mockResolvedValue(JSON.stringify({ ...approvingFormatReview(), complete: false,
        missing: ['Keep meaningful mathematical spacing'], summary: 'Whitespace cleanup must preserve equations' }));
    const result = await planDocumentTasks(deps(), { instruction: COMPOSITE_FORMAT_REQUEST });
    expect(result.tasks).toBeNull();
    expect(result.failure).toMatchObject({ phase: 'coverage-review', reason: expect.stringContaining('preserve equations') });
    expect(planReply).toHaveBeenCalledTimes(2);
    expect(reviewReply).toHaveBeenCalledTimes(2);
});

test('an unavailable selected target cannot silently widen to document scope', async () => {
    const plan = compositeFormatPlan(); plan.tasks[0].scope = 'selection';
    planReply.mockResolvedValue(JSON.stringify(plan));
    const result = await planDocumentTasks(deps(), { instruction: '整理选区格式' });
    expect(result.tasks).toBeNull();
    expect(result.failure).toMatchObject({ phase: 'scope', reason: expect.stringContaining('no selected target') });
    expect(reviewReply).not.toHaveBeenCalled();
});

test('legacy plans inherit explicit whole-document scope even when a selection is active', async () => {
    const plan = compositeFormatPlan(); plan.tasks.forEach((task) => { delete task.scope; task.instruction = '执行本任务'; });
    planReply.mockResolvedValue(JSON.stringify(plan));
    reviewReply.mockResolvedValue(JSON.stringify(approvingFormatReview()));
    const result = await planDocumentTasks(deps(), { instruction: COMPOSITE_FORMAT_REQUEST, hasSelection: true, hasTextSelection: true });
    expect(result.tasks.every((task) => task.scope === 'document')).toBe(true);
});

test('cancellation during review prevents a success or repair response', async () => {
    const controller = new AbortController();
    planReply.mockResolvedValue(JSON.stringify(compositeFormatPlan()));
    reviewReply.mockImplementationOnce(async () => { controller.abort(); return JSON.stringify(approvingFormatReview()); });
    await expect(planDocumentTasks(deps(), { instruction: COMPOSITE_FORMAT_REQUEST, signal: controller.signal }))
        .rejects.toMatchObject({ name: 'AbortError' });
    expect(planReply).toHaveBeenCalledTimes(1);
});

const abstractRequest = '根据全文内容重写一份摘要';
const abstractPlan = () => ({
    requirements: [
        { id: 'rewrite', kind: 'action', outcome: 'document', text: '重写摘要' },
        { id: 'source', kind: 'constraint', text: '根据全文内容' },
    ],
    tasks: [{ taskId: 'abstract', type: 'document_edit', scope: 'document',
        instruction: '阅读全文，根据全文内容重写现有摘要，仅修改摘要，保留其他正文。', covers: ['rewrite', 'source'], dependsOn: [] }],
    unsupported: [],
});
const abstractReview = () => ({ complete: true, unsupportedAccurate: true,
    checks: ['rewrite', 'source'].map((requirementId) => ({ requirementId, represented: true })), missing: [], invented: [] });

test.each([false, true])('abstract rewrite recovers from prose with history and selection, stream=%s', async (stream) => {
    const d = deps();
    d.conversationHistory = [
        { role: 'user', content: '全文润色修改' },
        { role: 'assistant', content: 'Applied 4 of 4 changes. Here is an abstract draft.' },
    ];
    planReply.mockResolvedValueOnce('请提供全文，我可以帮你重写摘要。').mockResolvedValueOnce(JSON.stringify(abstractPlan()));
    reviewReply.mockResolvedValue(JSON.stringify(abstractReview()));
    const onToken = stream ? jest.fn() : undefined;
    const result = await planDocumentTasks(d, { instruction: abstractRequest, hasSelection: true, hasTextSelection: true, onToken });
    expect(result.tasks).toMatchObject(abstractPlan().tasks);
    expect(result.review.complete).toBe(true);
    expect(planReply).toHaveBeenCalledTimes(2);
    expect(reviewReply).toHaveBeenCalledTimes(1);
    for (const [, messages] of planReply.mock.calls) {
        expect(messages.map(({ role }) => role)).toEqual(['system', 'user']);
        expect(messages[0].content).toContain('Return ONLY one JSON object');
        expect(messages[0].content).toContain('Document reading and writing happen later');
        expect(messages[0].content).not.toContain(abstractRequest);
        expect(JSON.parse(messages[1].content)).toMatchObject({ originalRequest: abstractRequest, requestContext: d.conversationHistory });
    }
    const repair = JSON.parse(planReply.mock.calls[1][1][1].content);
    expect(repair.previousResponse).toContain('请提供全文');
    expect(repair.validationFindings.reason).toContain('Reply contains no JSON object');
    expect(sendMessagesStream).toHaveBeenCalledTimes(stream ? 1 : 0);
    expect(sendPrompt).not.toHaveBeenCalled();
    if (stream) expect(onToken).toHaveBeenLastCalledWith(JSON.stringify(abstractPlan()));
});

test('empty streamed planner output is repaired through a non-streaming request', async () => {
    planReply.mockResolvedValueOnce('').mockResolvedValueOnce(JSON.stringify(abstractPlan()));
    reviewReply.mockResolvedValue(JSON.stringify(abstractReview()));
    const result = await planDocumentTasks(deps(), { instruction: abstractRequest, onReasoning: jest.fn() });
    expect(result.tasks).toHaveLength(1);
    expect(sendMessagesStream).toHaveBeenCalledTimes(1);
    expect(sendMessages).toHaveBeenCalledTimes(2); // repair and independent review
});

test('persistent non-JSON output stops after one repair without inventing an executable fallback', async () => {
    planReply.mockResolvedValue('Here is the rewritten abstract.');
    const result = await planDocumentTasks(deps(), { instruction: abstractRequest });
    expect(result.tasks).toBeNull();
    expect(result.failure).toEqual({ phase: 'contract', reason: 'Task planner: Reply contains no JSON object' });
    expect(planReply).toHaveBeenCalledTimes(2);
    expect(reviewReply).not.toHaveBeenCalled();
});

test('cancelling a malformed response prevents the repair request', async () => {
    const controller = new AbortController();
    planReply.mockImplementationOnce(async () => { controller.abort(); return ''; });
    await expect(planDocumentTasks(deps(), { instruction: abstractRequest, signal: controller.signal }))
        .rejects.toMatchObject({ name: 'AbortError' });
    expect(planReply).toHaveBeenCalledTimes(1);
    expect(reviewReply).not.toHaveBeenCalled();
});
