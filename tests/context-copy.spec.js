/** @jest-environment jsdom */
import { copyText, copyGeneratedText } from '../src/taskpane/ui/clipboard.js';
import { initContextCopy } from '../src/taskpane/ui/context-copy.js';
import { containFocus } from '../src/taskpane/ui/dialog.js';

let release;
const flush = async () => { for (let i = 0; i < 30; i++) await Promise.resolve(); };
beforeEach(() => {
    document.body.innerHTML = '<section role="dialog" id="aboutPanel"><button id="copyContextBtn">Copy all context</button>'
        + '<p id="contextCopyStatus"></p><div id="contextCopyFallback" hidden><textarea id="contextCopyText" readonly></textarea>'
        + '<button id="retryContextCopyBtn">Copy prepared text</button></div></section>';
    Object.defineProperty(navigator, 'clipboard', { configurable: true, value: {} });
});
afterEach(() => { release?.(); release = null; delete global.ClipboardItem; delete document.execCommand; });

test('legacy copy selects the entire payload inside the active focus trap', async () => {
    const button = document.getElementById('copyContextBtn');
    const panel = document.getElementById('aboutPanel');
    release = containFocus(panel, jest.fn(), button);
    const payload = 'Full report\n' + 'test '.repeat(10000);
    document.execCommand = jest.fn(() => {
        const field = document.activeElement;
        expect(panel.contains(field)).toBe(true);
        expect(field.tagName).toBe('TEXTAREA');
        expect(field.value.slice(field.selectionStart, field.selectionEnd)).toBe(payload);
        return true;
    });
    await copyText(payload);
    expect(document.execCommand).toHaveBeenCalledWith('copy');
    expect(document.activeElement).toBe(button);
    expect(panel.querySelectorAll('textarea')).toHaveLength(1);
});

test('promised clipboard write registers before the asynchronous report starts', async () => {
    let finish;
    const generate = jest.fn(() => new Promise((resolve) => { finish = resolve; }));
    global.ClipboardItem = class { constructor(data) { this.data = data; } };
    navigator.clipboard.write = jest.fn(async ([item]) => { await item.data['text/plain']; });
    const pending = copyGeneratedText(generate);
    expect(navigator.clipboard.write).toHaveBeenCalledTimes(1);
    expect(generate).not.toHaveBeenCalled();
    await flush();
    finish('Prepared after a slow Word read');
    await pending;
    expect(generate).toHaveBeenCalledTimes(1);
});

test('denied async write reuses the same report for dialog-scoped fallback', async () => {
    global.ClipboardItem = class { constructor(data) { this.data = data; } };
    navigator.clipboard.write = jest.fn(async () => { throw new Error('Denied'); });
    const generate = jest.fn(async () => 'Single snapshot');
    document.execCommand = jest.fn(() => true);
    await copyGeneratedText(generate, { container: document.getElementById('aboutPanel') });
    expect(generate).toHaveBeenCalledTimes(1);
    expect(document.execCommand).toHaveBeenCalledTimes(1);
});

test('blocked copy exposes a selected report and a fresh-click retry without re-reading Word', async () => {
    const getReport = jest.fn(async () => 'Ready context');
    document.execCommand = jest.fn(() => false);
    release = containFocus(document.getElementById('aboutPanel'), jest.fn(), document.getElementById('copyContextBtn'));
    initContextCopy({ getReport });
    document.getElementById('copyContextBtn').click();
    await flush();
    const preview = document.getElementById('contextCopyText');
    expect(document.getElementById('contextCopyFallback').hidden).toBe(false);
    expect(document.activeElement).toBe(preview);
    expect(preview.value.slice(preview.selectionStart, preview.selectionEnd)).toBe('Ready context');
    expect(document.getElementById('contextCopyStatus').textContent).not.toContain('All context copied');
    document.execCommand.mockReturnValue(true);
    document.getElementById('retryContextCopyBtn').click();
    // The command runs during click dispatch, before a microtask or host read.
    expect(document.execCommand).toHaveBeenCalledTimes(2);
    await flush();
    expect(getReport).toHaveBeenCalledTimes(1);
    expect(document.getElementById('contextCopyStatus').textContent).toBe('All context copied.');
});

test('report generation failure does not show stale prepared text or claim copy success', async () => {
    initContextCopy({ getReport: async () => { throw new Error('Asset unavailable'); } });
    document.getElementById('contextCopyText').value = 'Stale report';
    document.getElementById('copyContextBtn').click();
    await flush();
    expect(document.getElementById('contextCopyText').value).toBe('');
    expect(document.getElementById('contextCopyFallback').hidden).toBe(true);
    expect(document.getElementById('contextCopyStatus').textContent).toContain('Asset unavailable');
    expect(document.getElementById('copyContextBtn').disabled).toBe(false);
});
