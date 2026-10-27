/** @jest-environment jsdom */
jest.mock('../src/lib/document-parser.js', () => ({ parseDocument: jest.fn() }));
jest.mock('../src/lib/word-revisions.js', () => ({ queueRevisionRead: () => ({}), resolveRevisionRead: () => ({ text: 'Selected current text' }) }));
import { parseDocument } from '../src/lib/document-parser.js';
import { buildContextReport, redactContext } from '../src/taskpane/context-export.js';
import { copyText } from '../src/taskpane/ui/clipboard.js';
import { addLog, addLogWithRetry, clearLogs, getActivityLogText, initStatusBar } from '../src/taskpane/ui/status-bar.js';

beforeEach(() => {
    document.body.innerHTML = '<button id="copyLogsBtn">Copy all</button><span id="logCopyStatus"></span><div id="logs"></div>';
    global.fetch = jest.fn(async () => ({ ok: false }));
    jest.spyOn(console, 'log').mockImplementation(() => {});
    Object.defineProperty(navigator, 'clipboard', { configurable: true, value: { writeText: jest.fn(async () => {}) } });
    clearLogs();
});
afterEach(() => { jest.restoreAllMocks(); delete global.Word; delete document.execCommand; });

test('export includes all session logs, types and retry errors even after the DOM cap', () => {
    for (let i = 0; i < 210; i++) addLog(`entry ${i}`);
    addLogWithRetry('Retryable error', 'error', jest.fn());
    expect(document.getElementById('logs').children).toHaveLength(200);
    expect(getActivityLogText()).toContain('[INFO] entry 0');
    expect(getActivityLogText()).toContain('[ERROR] Retryable error');
    clearLogs();
    expect(getActivityLogText()).toBe('');
});

test('log copy redacts credentials and reports success', async () => {
    addLog('failed with key private-example');
    initStatusBar({ getConfig: () => ({ apiKey: 'private-example' }) });
    document.getElementById('copyLogsBtn').click();
    for (let i = 0; i < 20; i++) await Promise.resolve();
    expect(navigator.clipboard.writeText).toHaveBeenCalledWith(expect.stringContaining('[redacted]'));
    expect(navigator.clipboard.writeText.mock.calls[0][0]).not.toContain('private-example');
    expect(document.getElementById('logCopyStatus').textContent).toContain('copied');
});

test('context includes document, selection, live proposals, session, draft, attachments, prompts and safe settings', async () => {
    parseDocument.mockResolvedValue({ paragraphs: [{ text: 'Current paragraph' }, { text: 'Next paragraph' }] });
    global.Word = { run: async (fn) => fn({ document: { getSelection: () => ({ load: jest.fn() }) }, sync: async () => {} }) };
    const report = await buildContextReport({ state: { platform: 'Mac', config: { apiKey: 'secret-example', contextBudgetTokens: 128000 },
        promptManager: { getState: () => ({ context: 'custom prompt' }) } }, session: { messages: [{ text: 'secret-example', dataUrl: 'data:image/png;base64,abc' }] },
    chatText: 'Live pending proposal', logs: 'Full activity', draft: 'Pending instruction', attachments: 'data.docx', metadata: { build: 'abc123' } });
    for (const text of ['Current paragraph\nNext paragraph', 'Selected current text', 'Live pending proposal', 'Full activity', 'Pending instruction',
        'data.docx', 'custom prompt', '128000', 'abc123', '[binary attachment omitted]']) expect(report).toContain(text);
    expect(report).not.toContain('secret-example');
    expect(report).not.toContain('base64,abc');
});

test('unavailable Word reads are identified while remaining context is still exported', async () => {
    parseDocument.mockRejectedValue(new Error('Document unavailable'));
    const report = await buildContextReport({ state: { config: {} }, chatText: 'Chat survives', logs: 'Logs survive' });
    expect(report).toContain('Chat survives');
    expect(report).toContain('Logs survive');
    expect(report).toContain('Unavailable: Document unavailable');
});

test('redaction covers encoded configured keys, authorization, URL credentials and tokens', () => {
    const text = 'private/key private%2Fkey Bearer unknown-token https://name:pass@example.org/?token=other';
    const safe = redactContext(text, { providers: { custom: { apiKey: 'private/key' } } });
    for (const value of ['private/key', 'private%2Fkey', 'unknown-token', 'name:pass', 'token=other']) expect(safe).not.toContain(value);
});

test('clipboard denial falls back to copy command and restores focus', async () => {
    navigator.clipboard.writeText.mockRejectedValue(new Error('Denied'));
    document.execCommand = jest.fn(() => true);
    const button = document.getElementById('copyLogsBtn');
    button.focus();
    await copyText('full report');
    expect(document.execCommand).toHaveBeenCalledWith('copy');
    expect(document.activeElement).toBe(button);
    expect(document.querySelector('textarea')).toBeNull();
    document.execCommand.mockReturnValue(false);
    await expect(copyText('full report')).rejects.toThrow(/Clipboard unavailable/);
    expect(document.querySelector('textarea')).toBeNull();
});
