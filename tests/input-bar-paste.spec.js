/** @jest-environment jsdom */

jest.mock('../src/lib/file-store.js', () => ({ saveFile: jest.fn() }));
const fs = require('fs');
const path = require('path');
const { saveFile } = require('../src/lib/file-store.js');
const { initInputBar } = require('../src/taskpane/ui/input-bar.js');
const { ATTACHMENT_LIMITS } = require('../src/lib/file-attachments.js');
const html = fs.readFileSync(path.join(__dirname, '../src/taskpane/taskpane.html'), 'utf8');
let bar;
let onSubmit;
let onLog;
let textarea;

function paste(clipboardData, target = textarea) {
    const event = new Event('paste', { bubbles: true, cancelable: true });
    Object.defineProperty(event, 'clipboardData', { value: clipboardData });
    target.dispatchEvent(event);
    return event;
}

async function untilParsed() {
    const deadline = Date.now() + 2000;
    do {
        await new Promise((resolve) => setTimeout(resolve, 5));
        if (document.getElementById('sendBtn').getAttribute('aria-busy') !== 'true') return;
    } while (Date.now() < deadline);
    throw new Error('Attachment parsing did not finish.');
}

const names = () => [...document.querySelectorAll('.attachment-chip-name')].map((el) => el.textContent);
const send = () => document.getElementById('sendBtn').click();

beforeEach(() => {
    jest.clearAllMocks();
    document.documentElement.innerHTML = html;
    textarea = document.getElementById('chatInput');
    onSubmit = jest.fn();
    onLog = jest.fn();
    bar = initInputBar({ onSubmit, onLog, onCancel: jest.fn(), getSkills: () => [], onOpenSettings: jest.fn() });
});

afterEach(() => bar.clearAttachments());

test('pasted screenshot shows a pending chip and thumbnail, then sends once with the existing draft', async () => {
    const file = new File(['png'], 'image.png', { type: 'image/png' });
    const getAsFile = jest.fn(() => file);
    textarea.value = 'Explain this screenshot';
    expect(paste({ files: [file], items: [
        { kind: 'string', type: 'text/html' },
        { kind: 'file', type: 'image/png', getAsFile },
    ] }).defaultPrevented).toBe(true);
    // Clipboard data is consumed during the event, before asynchronous parsing.
    expect(getAsFile).toHaveBeenCalledTimes(1);
    expect(document.querySelectorAll('.attachment-chip-pending')).toHaveLength(1);
    send();
    expect(onSubmit).not.toHaveBeenCalled();
    await untilParsed();
    expect(names()).toEqual(['image.png']);
    expect(textarea.value).toBe('Explain this screenshot');
    expect(document.querySelector('.attachment-chip-thumb').src).toBe('data:image/png;base64,cG5n');
    expect(saveFile).not.toHaveBeenCalled();
    send();
    expect(onSubmit).toHaveBeenCalledWith('Explain this screenshot', [
        { name: 'image.png', kind: 'image', size: 3, dataUrl: 'data:image/png;base64,cG5n' },
    ]);
    expect(names()).toEqual([]);
});

test.each(['files', 'items'])('files copied from a file browser attach via clipboard %s and can send without text', async (representation) => {
    const files = [new File(['notes'], 'notes.txt'), new File(['# Heading'], 'draft.md')];
    const transfer = representation === 'files'
        ? { files }
        : { items: files.map((file) => ({ kind: 'file', getAsFile: () => file })) };
    expect(paste(transfer).defaultPrevented).toBe(true);
    await untilParsed();
    send();
    expect(onSubmit).toHaveBeenCalledWith('', [
        { name: 'notes.txt', kind: 'text', size: 5, text: 'notes' },
        { name: 'draft.md', kind: 'text', size: 9, text: '# Heading' },
    ]);
});

test('a missing item File falls back to the clipboard FileList without duplicating files', async () => {
    const files = [new File(['one'], 'one.txt'), new File(['two'], 'two.txt')];
    paste({ files, items: [
        { kind: 'string', type: 'text/plain' },
        { kind: 'file', getAsFile: () => null },
        { kind: 'file', getAsFile: () => files[1] },
    ] });
    await untilParsed();
    expect(names()).toEqual(['one.txt', 'two.txt']);
});

test.each([null, {}, { files: [], items: [{ kind: 'string', type: 'text/plain' }] },
    { files: [], items: [{ kind: 'string', type: 'text/html' }, { kind: 'string', type: 'text/uri-list' }] },
])('ordinary text, rich text and links keep native paste behavior: %j', (transfer) => {
    textarea.value = 'draft';
    expect(paste(transfer).defaultPrevented).toBe(false);
    expect(names()).toEqual([]);
    expect(document.getElementById('inputError').hidden).toBe(true);
    expect(textarea.value).toBe('draft');
});

test('pasting in another input does not attach to the conversation', () => {
    expect(paste({ files: [new File(['x'], 'file.txt')] }, document.getElementById('crossDocumentUrl')).defaultPrevented).toBe(false);
    expect(names()).toEqual([]);
});

test('unsupported, oversized and failed files report errors while valid files still attach', async () => {
    const huge = new File(['x'], 'huge.png', { type: 'image/png' });
    Object.defineProperty(huge, 'size', { value: ATTACHMENT_LIMITS.MAX_IMAGE_FILE_BYTES + 1 });
    const broken = new File(['x'], 'broken.txt');
    broken.text = async () => { throw new Error('read failed'); };
    paste({ files: [new File(['x'], 'archive.zip'), huge, broken, new File(['ok'], 'good.txt')] });
    await untilParsed();
    expect(names()).toEqual(['good.txt']);
    const error = document.getElementById('inputError');
    expect(error.hidden).toBe(false);
    expect(error.textContent).toContain('archive.zip: unsupported file type');
    expect(error.textContent).toContain('huge.png: 4.5 MB exceeds');
    expect(error.textContent).toContain('broken.txt: read failed');
    expect(onLog).toHaveBeenCalledWith('broken.txt: read failed', 'error');
});

test('pasting respects count and total size limits including existing attachments', async () => {
    bar.addAttachment({ name: 'existing.txt', kind: 'text', size: 1, text: 'x' });
    paste({ files: Array.from({ length: 5 }, (_, i) => new File(['x'], `${i}.txt`)) });
    await untilParsed();
    expect(names()).toHaveLength(5);
    expect(document.getElementById('inputError').textContent).toContain('4.txt: at most 5 attachments');
    bar.clearAttachments();
    bar.addAttachment({ name: 'large.txt', kind: 'text', size: ATTACHMENT_LIMITS.MAX_TOTAL_BYTES, text: 'x' });
    paste({ files: [new File(['x'], 'extra.txt')] });
    expect(names()).toEqual(['large.txt']);
    expect(document.getElementById('inputError').textContent).toContain('attachments total would exceed');
});

test('folders and unreadable items are reported without losing remaining readable files', async () => {
    const getFolder = jest.fn();
    const createReader = jest.fn();
    paste({ items: [
        { kind: 'file', webkitGetAsEntry: () => ({ name: 'folder', isDirectory: true, createReader }), getAsFile: getFolder },
        { kind: 'file', getAsFile: () => null },
        { kind: 'file', getAsFile: () => { throw new Error('denied'); } },
        { kind: 'file', getAsFile: () => new File(['ok'], 'good.txt') },
    ] });
    await untilParsed();
    expect(names()).toEqual(['good.txt']);
    expect(getFolder).not.toHaveBeenCalled();
    expect(createReader).not.toHaveBeenCalled();
    expect(document.getElementById('inputError').textContent).toContain('folder: folders are not supported');
    expect(document.getElementById('inputError').textContent).toContain('A pasted item could not be read');
});

test('processing and parsing reject new file pastes visibly while preserving the draft', async () => {
    const transfer = { files: [new File(['x'], 'retry.txt')] };
    textarea.value = 'draft';
    bar.setProcessing(true);
    expect(paste(transfer).defaultPrevented).toBe(true);
    expect(names()).toEqual([]);
    expect(document.getElementById('inputError').textContent).toContain('Files were not attached');
    bar.setProcessing(false);
    paste(transfer);
    expect(paste({ files: [new File(['x'], 'second.txt')] }).defaultPrevented).toBe(true);
    await untilParsed();
    expect(names()).toEqual(['retry.txt']);
    expect(textarea.value).toBe('draft');
    expect(document.getElementById('inputError').textContent).toContain('Files were not attached');
});

test('saving a pasted file blocks another file paste and preserves its library reference', async () => {
    const file = { name: 'notes.txt', size: 4, text: async () => 'body' };
    paste({ files: [file] });
    await untilParsed();
    let finishSave;
    saveFile.mockImplementationOnce(() => new Promise((resolve) => { finishSave = resolve; }));
    document.querySelector('.attachment-chip-save').click();
    await untilParsed();
    expect(saveFile).toHaveBeenCalledWith(file);
    expect(paste({ files: [new File(['x'], 'second.txt')] }).defaultPrevented).toBe(true);
    expect(names()).toEqual(['notes.txt']);
    finishSave({ fileId: 'file_1', versionId: 'v1', source: 'upload' });
    await untilParsed();
    send();
    expect(onSubmit.mock.calls[0][1][0]).toEqual({ name: 'notes.txt', size: 4, kind: 'text', text: 'body', fileId: 'file_1', versionId: 'v1', source: 'upload' });
});

test('removing a pending paste and starting a new chat both discard late parse results', async () => {
    let finishOld;
    paste({ files: [{ name: 'old.txt', size: 3, text: () => new Promise((resolve) => { finishOld = resolve; }) }] });
    document.querySelector('.attachment-chip-remove').click();
    finishOld('old');
    await untilParsed();
    expect(names()).toEqual([]);
    let finishCleared;
    paste({ files: [{ name: 'cleared.txt', size: 3, text: () => new Promise((resolve) => { finishCleared = resolve; }) }] });
    bar.clearAttachments();
    paste({ files: [new File(['new'], 'new.txt')] });
    finishCleared('old');
    await untilParsed();
    send();
    expect(onSubmit).toHaveBeenCalledWith('', [{ name: 'new.txt', kind: 'text', size: 3, text: 'new' }]);
});
