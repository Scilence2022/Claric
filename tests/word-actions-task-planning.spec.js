/** @jest-environment jsdom */
jest.mock('../src/lib/llm-client.js', () => ({ sendPrompt: jest.fn(), sendPromptStream: jest.fn(), sendMessages: jest.fn() }));
const { sendPrompt, sendMessages } = require('../src/lib/llm-client.js');
const { planDocumentTasks } = require('../src/taskpane/word-actions.js');
const { COMPOSITE_FORMAT_REQUEST, compositeFormatPlan, approvingFormatReview } = require('./fixtures/composite-format-plan.js');

const deps = () => ({ appState: { config: { backend: 'mock', providers: { mock: { model: 'm' } } } }, log: jest.fn() });
beforeEach(() => jest.clearAllMocks());

test('the reported compound request passes with shared coverage and explicit document scope', async () => {
    sendPrompt.mockResolvedValue(JSON.stringify(compositeFormatPlan()));
    sendMessages.mockResolvedValue(JSON.stringify(approvingFormatReview()));
    const d = deps();
    const result = await planDocumentTasks(d, { instruction: COMPOSITE_FORMAT_REQUEST, hasSelection: true, hasTextSelection: true });
    expect(result.tasks).toHaveLength(3);
    expect(result.tasks.map(({ type, scope, covers }) => ({ type, scope, covers }))).toEqual(compositeFormatPlan().tasks
        .map(({ type, scope, covers }) => ({ type, scope, covers })));
    expect(sendPrompt).toHaveBeenCalledTimes(1);
    const reviewMessages = sendMessages.mock.calls[0][1];
    expect(reviewMessages[0].role).toBe('system');
    expect(reviewMessages[0].content).toContain('native table/row borders');
    expect(reviewMessages[0].content).toContain('redundant spaces in body text');
    expect(JSON.parse(reviewMessages[1].content)).toMatchObject({ originalRequest: COMPOSITE_FORMAT_REQUEST,
        selectionFacts: { hasSelection: true }, plan: { tasks: compositeFormatPlan().tasks } });
});

test('repair receives actual rejected-review findings and can correct an invented blank-line operation', async () => {
    const first = compositeFormatPlan();
    first.tasks[0].instruction += '同时删除空段落。';
    sendPrompt.mockResolvedValueOnce(JSON.stringify(first)).mockResolvedValueOnce(JSON.stringify(compositeFormatPlan()));
    const negative = { ...approvingFormatReview(), complete: false, invented: ['Do not delete blank paragraphs; only spaces were requested'],
        missing: ['Preserve meaningful spaces in equations'], summary: 'Repair cleanup scope without adding paragraph deletion' };
    sendMessages.mockResolvedValueOnce(JSON.stringify(negative)).mockResolvedValueOnce(JSON.stringify(approvingFormatReview()));
    const d = deps();
    const result = await planDocumentTasks(d, { instruction: COMPOSITE_FORMAT_REQUEST });
    expect(result.tasks).toHaveLength(3);
    const repairPrompt = sendPrompt.mock.calls[1][1];
    expect(repairPrompt).toContain(negative.summary);
    expect(repairPrompt).toContain(negative.missing[0]);
    expect(repairPrompt).toContain(negative.invented[0]);
    expect(result.tasks[0].instruction).not.toContain('删除空段落');
    expect(d.log).toHaveBeenCalledWith(expect.stringContaining(negative.summary), 'warning');
    expect(sendMessages).toHaveBeenCalledTimes(2);
});

test('malformed coverage gives the repair a concrete contract reason instead of a blind retry', async () => {
    const first = compositeFormatPlan(); first.tasks[2].covers = ['r3', 'r5'];
    sendPrompt.mockResolvedValueOnce(JSON.stringify(first)).mockResolvedValueOnce(JSON.stringify(compositeFormatPlan()));
    sendMessages.mockResolvedValue(JSON.stringify(approvingFormatReview()));
    expect((await planDocumentTasks(deps(), { instruction: COMPOSITE_FORMAT_REQUEST })).tasks).toHaveLength(3);
    expect(sendPrompt.mock.calls[1][1]).toContain('a requirement has no task or unsupported explanation');
    expect(sendMessages).toHaveBeenCalledTimes(1);
});

test('persistent rejection stops execution and exposes the final concrete reason', async () => {
    sendPrompt.mockResolvedValue(JSON.stringify(compositeFormatPlan()));
    sendMessages.mockResolvedValue(JSON.stringify({ ...approvingFormatReview(), complete: false,
        missing: ['Keep meaningful mathematical spacing'], summary: 'Whitespace cleanup must preserve equations' }));
    const result = await planDocumentTasks(deps(), { instruction: COMPOSITE_FORMAT_REQUEST });
    expect(result.tasks).toBeNull();
    expect(result.failure).toMatchObject({ phase: 'coverage-review', reason: expect.stringContaining('preserve equations') });
    expect(sendPrompt).toHaveBeenCalledTimes(2);
    expect(sendMessages).toHaveBeenCalledTimes(2);
});

test('an unavailable selected target cannot silently widen to document scope', async () => {
    const plan = compositeFormatPlan(); plan.tasks[0].scope = 'selection';
    sendPrompt.mockResolvedValue(JSON.stringify(plan));
    const result = await planDocumentTasks(deps(), { instruction: '整理选区格式' });
    expect(result.tasks).toBeNull();
    expect(result.failure).toMatchObject({ phase: 'scope', reason: expect.stringContaining('no selected target') });
    expect(sendMessages).not.toHaveBeenCalled();
});

test('legacy plans inherit explicit whole-document scope even when a selection is active', async () => {
    const plan = compositeFormatPlan(); plan.tasks.forEach((task) => { delete task.scope; task.instruction = '执行本任务'; });
    sendPrompt.mockResolvedValue(JSON.stringify(plan));
    sendMessages.mockResolvedValue(JSON.stringify(approvingFormatReview()));
    const result = await planDocumentTasks(deps(), { instruction: COMPOSITE_FORMAT_REQUEST, hasSelection: true, hasTextSelection: true });
    expect(result.tasks.every((task) => task.scope === 'document')).toBe(true);
});

test('cancellation during review prevents a success or repair response', async () => {
    const controller = new AbortController();
    sendPrompt.mockResolvedValue(JSON.stringify(compositeFormatPlan()));
    sendMessages.mockImplementationOnce(async () => { controller.abort(); return JSON.stringify(approvingFormatReview()); });
    await expect(planDocumentTasks(deps(), { instruction: COMPOSITE_FORMAT_REQUEST, signal: controller.signal }))
        .rejects.toMatchObject({ name: 'AbortError' });
    expect(sendPrompt).toHaveBeenCalledTimes(1);
});
