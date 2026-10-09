jest.mock('../src/taskpane/task-module-loader.js', () => ({ loadFormatActions: jest.fn() }));
const { loadFormatActions } = require('../src/taskpane/task-module-loader.js');
const { prepareFormatProposal, applyFormatProposal, discardFormatProposal } = require('../src/taskpane/word-actions.js');

beforeEach(() => { jest.clearAllMocks(); global.Word = { run: jest.fn() }; });
afterEach(() => { delete global.Word; });

test('failed formatting download cannot capture a bookmark or run a native action', async () => {
    const failure = new Error('Formatting tools unavailable');
    loadFormatActions.mockRejectedValue(failure);
    const proposal = { anchor: { bookmark: '_test' }, ops: [{ font: { bold: true } }] };
    await expect(prepareFormatProposal({}, { instruction: '正文两端对齐' })).rejects.toBe(failure);
    await expect(applyFormatProposal({}, proposal)).rejects.toBe(failure);
    expect(Word.run).not.toHaveBeenCalled();
    expect(proposal.anchor.attempted).toBeUndefined();
});

test('loaded formatting actions are invoked once, with original request and cancellation intact', async () => {
    const prepared = { ops: [], anchor: {} };
    const native = { prepareFormatProposal: jest.fn(async () => prepared), applyFormatProposal: jest.fn(async () => ({ verifiedParagraphs: 1 })) };
    loadFormatActions.mockResolvedValue(native);
    const signal = new AbortController().signal;
    const deps = { log: jest.fn() };
    const args = { instruction: 'Expanded task', originalInstruction: '正文两端对齐', signal };
    expect(await prepareFormatProposal(deps, args)).toBe(prepared);
    expect(native.prepareFormatProposal).toHaveBeenCalledTimes(1);
    expect(native.prepareFormatProposal).toHaveBeenCalledWith(expect.objectContaining({ log: deps.log, sendActionRequest: expect.any(Function) }), args);
    expect(await applyFormatProposal(deps, prepared, { signal })).toEqual({ verifiedParagraphs: 1 });
    expect(native.applyFormatProposal).toHaveBeenCalledTimes(1);
    expect(native.applyFormatProposal).toHaveBeenCalledWith(deps, prepared, { signal });
    expect(loadFormatActions.mock.calls).toEqual([[{ signal }], [{ signal }]]);
});

test('bookmark cleanup stays available without downloading formatting tools', async () => {
    loadFormatActions.mockRejectedValue(new Error('Download unavailable'));
    const context = { document: { deleteBookmark: jest.fn() }, sync: jest.fn(async () => {}) };
    Word.run.mockImplementation((callback) => callback(context));
    const proposal = { anchor: { bookmark: '_claric_img_test' } };
    await discardFormatProposal({}, proposal);
    await discardFormatProposal({}, proposal);
    expect(context.document.deleteBookmark).toHaveBeenCalledTimes(1);
    expect(proposal.anchor.cleaned).toBe(true);
    expect(loadFormatActions).not.toHaveBeenCalled();
});
