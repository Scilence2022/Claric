/** @jest-environment jsdom */
jest.mock('../src/taskpane/ui/status-bar.js', () => ({ addLog: jest.fn() }));
const { initUndoHistoryControl } = require('../src/taskpane/ui/undo-history.js');
const { appState } = require('../src/taskpane/app-state.js');
const { addLog } = require('../src/taskpane/ui/status-bar.js');

const settle = async () => { for (let i = 0; i < 12; i++) await Promise.resolve(); };
let context;
let button;
let status;

beforeEach(() => {
    jest.clearAllMocks();
    Object.assign(appState, { isProcessing: false, isProcessingDoc: false, isProcessingSummary: false,
        processDocController: null, chatController: null });
    document.body.innerHTML = '<button id="clearUndoHistoryBtn">Clear undo history</button>'
        + '<small id="clearUndoHistoryStatus" role="status" hidden></small>';
    context = { document: { undoClear: jest.fn(), changeTrackingMode: 'TrackAll' }, sync: jest.fn(async () => {}) };
    global.Office = { context: { requirements: { isSetSupported: jest.fn(() => true) } } };
    global.Word = { run: jest.fn(async (callback) => callback(context)) };
    button = document.getElementById('clearUndoHistoryBtn');
    status = document.getElementById('clearUndoHistoryStatus');
    initUndoHistoryControl();
});

afterEach(() => { delete global.Word; delete global.Office; });

function confirm() {
    document.querySelector('[data-confirm="clear-undo"]').click();
}

test('requires explicit confirmation, clears only undo history, and reports success after sync', async () => {
    let finishSync;
    context.sync.mockImplementation(() => new Promise((resolve) => { finishSync = resolve; }));
    button.focus();
    button.click();
    expect(Office.context.requirements.isSetSupported).toHaveBeenCalledWith('WordApiDesktop', '1.4');
    expect(Word.run).not.toHaveBeenCalled();
    expect(document.activeElement.dataset.confirm).toBe('cancel');
    expect(document.querySelector('[role="alertdialog"]').textContent).toContain('including edits made before using Claric');
    confirm();
    await settle();
    expect(context.document.undoClear).toHaveBeenCalledTimes(1);
    expect(context.document.changeTrackingMode).toBe('TrackAll');
    expect(context.sync).toHaveBeenCalledTimes(1);
    expect(button.disabled).toBe(true);
    expect(appState.isProcessing).toBe(true);
    expect(status.textContent).toBe('Clearing undo history...');
    finishSync();
    await settle();
    expect(status.hidden).toBe(false);
    expect(status.textContent).toMatch(/Undo history cleared \(\d+ ms\)/);
    expect(button.disabled).toBe(false);
    expect(button.hasAttribute('aria-busy')).toBe(false);
    expect(appState.isProcessing).toBe(false);
    expect(addLog).toHaveBeenCalledWith(status.textContent, 'success');
});

test.each(['cancel', 'escape'])('%s leaves Word untouched and restores focus', async (action) => {
    button.focus();
    button.click();
    const cancel = document.querySelector('[data-confirm="cancel"]');
    if (action === 'cancel') cancel.click();
    else cancel.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    await settle();
    expect(Word.run).not.toHaveBeenCalled();
    expect(document.querySelector('[role="alertdialog"]')).toBeNull();
    expect(document.activeElement).toBe(button);
    expect(appState.isProcessing).toBe(false);
});

test('unsupported hosts explain the requirement without opening a dialog or calling Word', async () => {
    Office.context.requirements.isSetSupported.mockReturnValue(false);
    button.click();
    await settle();
    expect(status.textContent).toContain('WordApiDesktop 1.4');
    expect(document.querySelector('[role="alertdialog"]')).toBeNull();
    expect(Word.run).not.toHaveBeenCalled();
});

test('missing Office runtime produces a visible error', async () => {
    delete global.Office;
    button.click();
    await settle();
    expect(status.hidden).toBe(false);
    expect(status.textContent).toContain('desktop version of Word');
    expect(Word.run).not.toHaveBeenCalled();
});

test.each(['isProcessing', 'isProcessingDoc', 'isProcessingSummary', 'processDocController', 'chatController'])(
    'does not clear history during %s or release another operation’s lock', async (flag) => {
        appState[flag] = true;
        button.click();
        await settle();
        expect(status.textContent).toContain('Wait for Claric');
        expect(Word.run).not.toHaveBeenCalled();
        expect(appState[flag]).toBe(true);
    });

test('rechecks processing after confirmation and prevents duplicate dialogs', async () => {
    button.click();
    button.click();
    expect(document.querySelectorAll('[role="alertdialog"]')).toHaveLength(1);
    appState.isProcessingDoc = true;
    confirm();
    await settle();
    expect(Word.run).not.toHaveBeenCalled();
    expect(status.textContent).toContain('Wait for Claric');
    expect(appState.isProcessingDoc).toBe(true);
});

test.each(['missing method', 'sync failure'])('%s reports failure, unlocks the UI, and permits retry', async (failure) => {
    if (failure === 'missing method') delete context.document.undoClear;
    else context.sync.mockRejectedValueOnce(new Error('Word is unavailable'));
    button.click();
    confirm();
    await settle();
    expect(status.textContent).toContain('Could not clear undo history:');
    expect(addLog).not.toHaveBeenCalledWith(expect.anything(), 'success');
    expect(button.disabled).toBe(false);
    expect(appState.isProcessing).toBe(false);
    context.document.undoClear = jest.fn();
    button.click();
    confirm();
    await settle();
    expect(status.textContent).toContain('Undo history cleared');
});
