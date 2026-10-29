/** @jest-environment jsdom */
const { clearUndoHistoryAfterChangeSet, clearUndoHistory } = require('../src/taskpane/undo-history.js');
const { createProposalCard } = require('../src/taskpane/ui/proposal-card.js');

const enabled = { clearUndoHistoryAfterApply: true };
const flush = async () => { for (let i = 0; i < 30; i++) await Promise.resolve(); };
let context;
beforeEach(() => {
    context = { document: { undoClear: jest.fn(), changeTrackingMode: 'TrackAll' }, sync: jest.fn(async () => {}) };
    global.Office = { context: { platform: 'Mac', diagnostics: { version: '16.104' },
        requirements: { isSetSupported: jest.fn(() => true) } } };
    global.Word = { run: jest.fn(async (callback) => callback(context)) };
});
afterEach(() => { delete global.Word; delete global.Office; });

function cardWithCleanup(options = {}) {
    return createProposalCard({ title: 'Changes', onApply: async () => {},
        onChangeSetApplied: () => clearUndoHistoryAfterChangeSet(enabled), ...options });
}
const items = ['a', 'b', 'c'].map((id) => ({ id, label: id, before: 'before', after: 'after' }));

test.each([undefined, {}, { clearUndoHistoryAfterApply: false }, { clearUndoHistoryAfterApply: 'true' }])(
    'requires an explicit opt-in: %p', async (config) => {
        expect(await clearUndoHistoryAfterChangeSet(config)).toBeNull();
        expect(Word.run).not.toHaveBeenCalled();
    });

test('Mac capability misreport attempts the native command and waits for sync before claiming success', async () => {
    Office.context.requirements.isSetSupported.mockReturnValue(false);
    let finish;
    context.sync.mockImplementation(() => new Promise((resolve) => { finish = resolve; }));
    const log = jest.fn();
    const pending = clearUndoHistoryAfterChangeSet(enabled, log);
    await flush();
    expect(context.document.undoClear).toHaveBeenCalledTimes(1);
    expect(log).not.toHaveBeenCalledWith(expect.anything(), 'success');
    finish();
    expect(await pending).toMatchObject({ type: 'success' });
    expect(context.document.changeTrackingMode).toBe('TrackAll');
});

test.each(['missing method', 'native error', 'unsupported web'])('%s returns a warning without a false success', async (failure) => {
    Office.context.requirements.isSetSupported.mockReturnValue(false);
    if (failure === 'missing method') delete context.document.undoClear;
    if (failure === 'native error') context.sync.mockRejectedValue(Object.assign(new Error('Not supported'), { code: 'NotImplemented' }));
    if (failure === 'unsupported web') Office.context.platform = 'OfficeOnline';
    const log = jest.fn();
    const result = await clearUndoHistoryAfterChangeSet(enabled, log);
    expect(result.type).toBe('warning');
    expect(result.message).toContain('ChangeSet applied, but undo history was not cleared');
    expect(result.message).toContain('16.104');
    if (failure === 'native error') expect(result.message).toContain('NotImplemented');
    if (failure === 'unsupported web') expect(Word.run).not.toHaveBeenCalled();
    expect(log).not.toHaveBeenCalledWith(expect.anything(), 'success');
});

test('missing runtime rejects a manual command and warns for opted-in cleanup', async () => {
    delete global.Word;
    await expect(clearUndoHistory()).rejects.toThrow('desktop version of Word');
    expect(await clearUndoHistoryAfterChangeSet(enabled)).toMatchObject({ type: 'warning' });
});

test('each verified section clears once before the next section, without a duplicate final cleanup', async () => {
    const card = cardWithCleanup({ items, onApply: async (ids, progress) => {
        for (const [index, id] of ids.entries()) {
            expect(context.sync).toHaveBeenCalledTimes(index);
            await progress.onChunkApplied(id, { applied: true });
            await progress.onChunkApplied(id, { applied: true });
            expect(context.sync).toHaveBeenCalledTimes(index + 1);
        }
    } });
    await card.applyAll();
    expect(context.document.undoClear).toHaveBeenCalledTimes(3);
    expect(card.el.dataset.outcome).toBe('applied');
    expect(card.el.querySelector('.proposal-cleanup-status').textContent).toContain('3 ChangeSet(s)');
    await card.applyAll();
    expect(context.document.undoClear).toHaveBeenCalledTimes(3);
});

test.each([{ noChange: true }, { skipped: true }, { error: true }, { partial: true }, { uncertain: true }])(
    'does not clear for an unverified or unchanged section: %p', async (outcome) => {
        const card = cardWithCleanup({ items: items.slice(0, 1),
            onApply: async (_ids, progress) => progress.onChunkApplied('a', outcome) });
        await card.applyAll();
        expect(Word.run).not.toHaveBeenCalled();
    });

test('a successful single ChangeSet holds the apply lock through native cleanup', async () => {
    let finish;
    context.sync.mockImplementation(() => new Promise((resolve) => { finish = resolve; }));
    const busy = jest.fn();
    const card = cardWithCleanup({ setApplyBusy: busy });
    const pending = card.applyAll();
    await flush();
    const otherApply = jest.fn();
    const other = cardWithCleanup({ onApply: otherApply });
    await other.applyAll();
    expect(otherApply).not.toHaveBeenCalled();
    expect(busy).toHaveBeenLastCalledWith(true);
    finish();
    await pending;
    expect(busy).toHaveBeenLastCalledWith(false);
    expect(context.document.undoClear).toHaveBeenCalledTimes(1);
});

test.each(['warning', 'error', 'no-change'])('single ChangeSet %s does not clear earlier manual edits', async (outcome) => {
    const card = cardWithCleanup({ onApply: async () => {
        if (outcome === 'no-change') card.markApplied('Already satisfied', { noChange: true });
        else if (outcome === 'error') throw new Error('Apply failed');
        else card.markWarning('Unverified');
    } });
    await card.applyAll();
    expect(Word.run).not.toHaveBeenCalled();
});

test('cleanup failure remains visible alongside an applied edit and cannot cause replay', async () => {
    context.sync.mockRejectedValue(new Error('Not supported'));
    const apply = jest.fn(async () => {});
    const card = cardWithCleanup({ onApply: apply });
    await card.applyAll();
    expect(card.el.dataset.outcome).toBe('applied');
    expect(card.el.querySelector('.proposal-cleanup-status').textContent).toContain('0 succeeded, 1 failed');
    expect(card.el.querySelector('.proposal-cleanup-status').textContent).toContain('Not supported');
    await card.applyAll();
    expect(apply).toHaveBeenCalledTimes(1);
});

test('pause, resume and explicit recovery clean only newly verified sections', async () => {
    let pass = 0;
    const card = cardWithCleanup({ items, onApply: async (_ids, progress) => {
        if (pass++ === 0) {
            await progress.onChunkApplied('c', { applied: true });
            await progress.onChunkApplied('b', { error: true });
            card.setPaused('Paused');
        } else await progress.onChunkApplied('a', { applied: true });
    }, onRecover: async (id, _mode, progress) => {
        await progress.onChunkApplied(id, { applied: true });
        return { errors: [] };
    } });
    await card.applyAll();
    expect(context.document.undoClear).toHaveBeenCalledTimes(1);
    await card.applyAll();
    expect(context.document.undoClear).toHaveBeenCalledTimes(2);
    card.el.querySelectorAll('.proposal-card-change')[1].querySelector('[data-recovery="retry"]').click();
    await flush();
    expect(context.document.undoClear).toHaveBeenCalledTimes(3);
    expect(card.el.dataset.outcome).toBe('applied');
});
