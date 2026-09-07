/** @jest-environment jsdom */
jest.mock('../src/lib/file-store.js', () => ({ saveFile: jest.fn() }));
const { saveFile } = require('../src/lib/file-store.js');
const { initInputBar } = require('../src/taskpane/ui/input-bar.js');
const flush = async () => { for (let i = 0; i < 12; i++) await Promise.resolve(); };
const parsed = { name: 'notes.txt', kind: 'text', size: 4, text: 'body', fileId: 'file_1', versionId: 'v1', source: 'upload' };
let bar;
let onSubmit;
let pick;
beforeEach(() => {
    jest.clearAllMocks();
    document.body.innerHTML = '<textarea id="chatInput"></textarea><button id="sendBtn"></button><div id="skillPicker" hidden></div><button id="modelPill"></button><div id="attachmentChips" hidden></div><div id="inputError" hidden></div><button id="attachBtn"></button><input type="file" id="attachmentInput" multiple>';
    onSubmit = jest.fn();
    bar = initInputBar({ onSubmit, onCancel: jest.fn(), getSkills: () => [], onOpenSettings: jest.fn() });
    pick = (files) => {
        const input = document.getElementById('attachmentInput');
        Object.defineProperty(input, 'files', { value: files, configurable: true });
        input.dispatchEvent(new Event('change'));
    };
});

test('addAttachment enforces payload, count, size and busy validation', () => {
    expect(bar.addAttachment(null)).toBe(false);
    expect(bar.addAttachment({ ...parsed, kind: 'unknown' })).toBe(false);
    expect(bar.addAttachment({ ...parsed, size: -1 })).toBe(false);
    expect(bar.addAttachment({ ...parsed, size: 11 * 1024 * 1024 })).toBe(false);
    expect(bar.addAttachment({ ...parsed, pending: true })).toBe(false);
    expect(bar.addAttachment({ ...parsed, kind: 'image', dataUrl: 'javascript:alert(1)' })).toBe(false);
    bar.setProcessing(true);
    expect(bar.addAttachment(parsed)).toBe(false);
    bar.setProcessing(false);
    for (let i = 0; i < 5; i++) expect(bar.addAttachment(parsed)).toBe(true);
    expect(bar.addAttachment(parsed)).toBe(false);
    expect(document.querySelectorAll('.attachment-chip')).toHaveLength(5);
});

test('library attachment metadata and content are a snapshot, without raw objects', () => {
    const attachment = { ...parsed, rawFile: new File(['body'], 'notes.txt'), arbitrary: { value: 1 } };
    expect(bar.addAttachment(attachment)).toBe(true);
    attachment.name = 'changed';
    attachment.text = 'changed';
    attachment.versionId = 'v2';
    document.getElementById('sendBtn').click();
    const sent = onSubmit.mock.calls[0][1][0];
    expect(sent).toEqual(parsed);
    expect(Object.isFrozen(sent)).toBe(true);
    expect(saveFile).not.toHaveBeenCalled();
});

test('pending local parse prevents adding a library file', async () => {
    let resolve;
    pick([{ name: 'slow.txt', size: 4, text: () => new Promise((done) => { resolve = done; }) }]);
    expect(bar.addAttachment(parsed)).toBe(false);
    resolve('body');
    await flush();
    expect(bar.addAttachment(parsed)).toBe(true);
});

test('local uploads save only by explicit per-chip action and retain immutable reference metadata', async () => {
    const file = { name: 'notes.txt', size: 4, text: async () => 'body' };
    pick([file]);
    await flush();
    expect(saveFile).not.toHaveBeenCalled();
    let resolve;
    saveFile.mockImplementationOnce(() => new Promise((done) => { resolve = done; }));
    document.querySelector('.attachment-chip-save').click();
    await flush();
    expect(saveFile).toHaveBeenCalledWith(file);
    expect(document.querySelector('.attachment-chip-save').textContent).toBe('Saving…');
    expect(bar.addAttachment(parsed)).toBe(false);
    document.getElementById('sendBtn').click();
    expect(onSubmit).not.toHaveBeenCalled();
    resolve({ fileId: 'file_1', versionId: 'v1', source: 'upload' });
    await flush();
    expect(document.querySelector('.attachment-chip-save')).toBeNull();
    document.getElementById('sendBtn').click();
    expect(onSubmit.mock.calls[0][1][0]).toEqual(parsed);
});

test('save failure is visible, retryable and never discards the local attachment', async () => {
    pick([{ name: 'notes.txt', size: 4, text: async () => 'body' }]);
    await flush();
    saveFile.mockRejectedValueOnce(new Error('Quota exceeded'));
    document.querySelector('.attachment-chip-save').click();
    await flush();
    expect(document.getElementById('inputError').textContent).toBe('Quota exceeded');
    expect(document.querySelector('.attachment-chip-save').disabled).toBe(false);
    document.getElementById('sendBtn').click();
    expect(onSubmit.mock.calls[0][1][0]).toMatchObject({ name: 'notes.txt', text: 'body' });
    expect(onSubmit.mock.calls[0][1][0].fileId).toBeUndefined();
});

test('clearing or removing an attachment during a save does not restore it later', async () => {
    pick([{ name: 'notes.txt', size: 4, text: async () => 'body' }]);
    await flush();
    let resolve;
    saveFile.mockImplementationOnce(() => new Promise((done) => { resolve = done; }));
    document.querySelector('.attachment-chip-save').click();
    await flush();
    bar.clearAttachments();
    expect(bar.addAttachment(parsed)).toBe(true);
    resolve({ fileId: 'other', versionId: 'v2', source: 'upload' });
    await flush();
    expect(document.querySelectorAll('.attachment-chip')).toHaveLength(1);
    document.getElementById('sendBtn').click();
    expect(onSubmit.mock.calls[0][1][0]).toEqual(parsed);
});
