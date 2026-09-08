/** @jest-environment jsdom */

jest.mock('../src/lib/file-store.js', () => ({ saveFile: jest.fn() }));
const fs = require('fs');
const path = require('path');
const { saveFile } = require('../src/lib/file-store.js');
const { initInputBar } = require('../src/taskpane/ui/input-bar.js');
const { ATTACHMENT_LIMITS } = require('../src/lib/file-attachments.js');
const html = fs.readFileSync(path.join(__dirname, '../src/taskpane/taskpane.html'), 'utf8');
const flush = () => new Promise((resolve) => setTimeout(resolve, 30));
const names = () => [...document.querySelectorAll('.attachment-chip-name')].map((el) => el.textContent);
let app;
let bar;
let onSubmit;
let onLog;

function drag(type, target = app, transfer = { types: ['Files'], files: [] }, relatedTarget = null) {
    const event = new Event(type, { bubbles: true, cancelable: true });
    Object.defineProperties(event, { dataTransfer: { value: transfer }, relatedTarget: { value: relatedTarget } });
    target.dispatchEvent(event);
    return event;
}

function drop(files, target = app, extra = {}) {
    return drag('drop', target, { types: ['Files'], files, ...extra });
}

function send(text = '') {
    document.getElementById('chatInput').value = text;
    document.getElementById('sendBtn').click();
}

beforeEach(() => {
    jest.clearAllMocks();
    document.documentElement.innerHTML = html;
    app = document.querySelector('.app');
    onSubmit = jest.fn();
    onLog = jest.fn();
    bar = initInputBar({ onSubmit, onLog, onCancel: jest.fn(), getSkills: () => [], onOpenSettings: jest.fn() });
});

afterEach(() => {
    bar.clearAttachments();
    jest.restoreAllMocks();
});

test('drop uses FileReader, pending chips, real parsing and the final submit payload without saving', async () => {
    const read = jest.spyOn(FileReader.prototype, 'readAsText');
    const files = [new File(['hello world'], 'notes.txt', { type: 'text/plain' }), new File(['abc'], 'photo.png', { type: 'image/png' })];
    expect(drop(files, document.getElementById('chatInput')).defaultPrevented).toBe(true);
    expect(document.querySelectorAll('.attachment-chip-pending')).toHaveLength(2);
    expect(document.getElementById('sendBtn').getAttribute('aria-busy')).toBe('true');
    send('summarize');
    expect(onSubmit).not.toHaveBeenCalled();
    await flush();
    expect(read).toHaveBeenCalledWith(files[0]);
    expect(document.querySelectorAll('.attachment-chip-pending')).toHaveLength(0);
    expect(document.querySelectorAll('.attachment-chip-save')).toHaveLength(2);
    expect(document.querySelector('.attachment-chip-thumb').src).toBe('data:image/png;base64,YWJj');
    expect(saveFile).not.toHaveBeenCalled();
    send('summarize');
    expect(onSubmit).toHaveBeenCalledWith('summarize', [
        { name: 'notes.txt', kind: 'text', size: 11, text: 'hello world' },
        { name: 'photo.png', kind: 'image', size: 3, dataUrl: 'data:image/png;base64,YWJj' },
    ]);
    expect(names()).toEqual([]);
});

test.each(['settingsOverlay', 'historyOverlay', 'filesOverlay', 'savePromptModal'])('drop on %s attaches to the composer even if a child stops bubbling', async (id) => {
    const overlay = document.getElementById(id);
    overlay.hidden = false;
    expect(app.contains(overlay)).toBe(true);
    overlay.addEventListener('drop', (event) => event.stopPropagation());
    expect(drag('dragover', overlay).defaultPrevented).toBe(true);
    expect(drop([new File(['body'], 'overlay.txt')], overlay).defaultPrevented).toBe(true);
    await flush();
    send();
    expect(onSubmit.mock.calls[0][1]).toEqual([{ name: 'overlay.txt', kind: 'text', size: 4, text: 'body' }]);
    expect(saveFile).not.toHaveBeenCalled();
});

test('mixed unsupported, oversized, failed and valid files preserve each error and valid attachment', async () => {
    const huge = new File(['x'], 'huge.txt');
    Object.defineProperty(huge, 'size', { value: ATTACHMENT_LIMITS.MAX_TEXT_FILE_BYTES + 1 });
    const broken = new File(['x'], 'broken.txt');
    broken.text = async () => { throw new Error('read failed'); };
    drop([new File(['zip'], 'bad.zip'), huge, broken, new File(['ok'], 'good.md')]);
    await flush();
    expect(names()).toEqual(['good.md']);
    const error = document.getElementById('inputError');
    expect(error.hidden).toBe(false);
    expect(error.textContent).toContain('bad.zip: unsupported file type');
    expect(error.textContent).toContain('huge.txt: 10.0 MB exceeds');
    expect(error.textContent).toContain('broken.txt: read failed');
    expect(onLog).toHaveBeenCalledWith('broken.txt: read failed', 'error');
    expect(app.classList.contains('file-drop-error')).toBe(true);
    send();
    expect(onSubmit.mock.calls[0][1][0].text).toBe('ok');
});

test('drop respects count and total limits including existing library attachments', async () => {
    bar.addAttachment({ name: 'existing.txt', kind: 'text', size: 1, text: 'a' });
    drop(Array.from({ length: 5 }, (_, i) => new File(['x'], `${i}.txt`)));
    expect(names()).toHaveLength(5);
    await flush();
    expect(document.getElementById('inputError').textContent).toContain('4.txt: at most 5 attachments');
    bar.clearAttachments();
    bar.addAttachment({ name: 'large.txt', kind: 'text', size: ATTACHMENT_LIMITS.MAX_TOTAL_BYTES, text: 'a' });
    drop([new File(['x'], 'extra.txt')]);
    expect(names()).toEqual(['large.txt']);
    expect(document.getElementById('inputError').textContent).toContain('attachments total would exceed');
});

test('file items work without files; directories with supported-looking names are skipped without traversal', async () => {
    const file = new File(['ok'], 'good.txt');
    const folderFile = jest.fn(() => new File([], 'folder.txt'));
    const createReader = jest.fn();
    drop([], app, { items: [
        { kind: 'file', webkitGetAsEntry: () => ({ name: 'folder.txt', isDirectory: true, createReader }), getAsFile: folderFile },
        { kind: 'string', getAsFile: jest.fn() },
        { kind: 'file', webkitGetAsEntry: () => null, getAsFile: () => file },
    ] });
    await flush();
    expect(names()).toEqual(['good.txt']);
    expect(folderFile).not.toHaveBeenCalled();
    expect(createReader).not.toHaveBeenCalled();
    expect(document.getElementById('inputError').textContent).toContain('folder.txt: folders are not supported');
    send();
    expect(onSubmit.mock.calls[0][1][0].text).toBe('ok');
});

test('unreadable and inaccessible items are reported while remaining files attach', async () => {
    drop([], app, { items: [
        { kind: 'file', getAsFile: () => null },
        { kind: 'file', webkitGetAsEntry: () => { throw new Error('denied'); } },
        { kind: 'file', getAsFile: () => new File(['ok'], 'good.txt') },
    ] });
    await flush();
    expect(names()).toEqual(['good.txt']);
    expect(document.getElementById('inputError').textContent).toContain('could not be read');
});

test('processing rejects drops visibly, prevents navigation, retains draft and accepts a retry', async () => {
    document.getElementById('chatInput').value = 'draft';
    drag('dragenter');
    bar.setProcessing(true);
    expect(app.classList.contains('file-drag-over')).toBe(false);
    const transfer = { types: ['Files'], files: [] };
    expect(drag('dragover', app, transfer).defaultPrevented).toBe(true);
    expect(transfer.dropEffect).toBe('none');
    expect(app.classList.contains('file-drag-busy')).toBe(true);
    expect(drop([new File(['x'], 'retry.txt')]).defaultPrevented).toBe(true);
    expect(names()).toEqual([]);
    expect(document.getElementById('inputError').textContent).toContain('Files were not attached');
    expect(document.getElementById('chatInput').value).toBe('draft');
    bar.setProcessing(false);
    drop([new File(['x'], 'retry.txt')]);
    await flush();
    expect(names()).toEqual(['retry.txt']);
});

test('parsing rejects a second drop and the rejection survives the first parse finishing', async () => {
    drop([new File(['first'], 'first.txt')]);
    expect(drop([new File(['second'], 'second.txt')]).defaultPrevented).toBe(true);
    expect(document.getElementById('inputError').textContent).toContain('Files were not attached');
    await flush();
    expect(names()).toEqual(['first.txt']);
    expect(document.getElementById('inputError').hidden).toBe(false);
    send();
    expect(onSubmit.mock.calls[0][1].map((att) => att.name)).toEqual(['first.txt']);
});

test('saving rejects drops and picker changes, retaining raw file and explicit save metadata', async () => {
    const file = new File(['body'], 'save.txt');
    drop([file]);
    await flush();
    let finishSave;
    saveFile.mockImplementationOnce(() => new Promise((resolve) => { finishSave = resolve; }));
    drag('dragenter');
    document.querySelector('.attachment-chip-save').click();
    expect(app.classList.contains('file-drag-over')).toBe(false);
    await flush();
    expect(saveFile).toHaveBeenCalledWith(file);
    expect(drop([new File(['x'], 'lost.txt')]).defaultPrevented).toBe(true);
    const input = document.getElementById('attachmentInput');
    Object.defineProperty(input, 'files', { value: [new File(['x'], 'picked.txt')] });
    input.dispatchEvent(new Event('change'));
    expect(names()).toEqual(['save.txt']);
    expect(document.getElementById('inputError').textContent).toContain('Files were not attached');
    send();
    expect(onSubmit).not.toHaveBeenCalled();
    finishSave({ fileId: 'file_1', versionId: 'v1', source: 'upload' });
    await flush();
    send();
    expect(onSubmit.mock.calls[0][1]).toEqual([{ name: 'save.txt', size: 4, kind: 'text', text: 'body', fileId: 'file_1', versionId: 'v1', source: 'upload' }]);
});

test('nested enter/leave does not flicker; leaving, dragend, blur, drop and clear remove feedback', () => {
    const child = document.getElementById('chatInput');
    const sibling = document.getElementById('sendBtn');
    const transfer = { types: ['Files'], files: [] };
    drag('dragenter', app, transfer);
    expect(transfer.dropEffect).toBe('copy');
    drag('dragenter', child);
    drag('dragleave', app, undefined, child);
    expect(app.classList.contains('file-drag-over')).toBe(true);
    drag('dragenter', sibling);
    drag('dragleave', child, undefined, sibling);
    expect(app.classList.contains('file-drag-over')).toBe(true);
    drag('dragleave', sibling);
    expect(app.classList.contains('file-drag-over')).toBe(false);
    for (const finish of [() => drag('dragend'), () => window.dispatchEvent(new Event('blur')), () => drop([]), () => bar.clearAttachments()]) {
        drag('dragenter');
        expect(app.classList.contains('file-drag-over')).toBe(true);
        finish();
        expect(app.classList.contains('file-drag-over')).toBe(false);
    }
});

test.each([['text/plain'], ['text/uri-list'], ['text/html', 'text/plain']])('non-file %j drag keeps all browser defaults', (...types) => {
    const transfer = { types, files: [], items: [{ kind: 'string' }], dropEffect: 'move' };
    for (const type of ['dragenter', 'dragover', 'dragleave', 'drop']) {
        expect(drag(type, document.getElementById('chatInput'), transfer).defaultPrevented).toBe(false);
    }
    expect(transfer.dropEffect).toBe('move');
    expect(app.classList.contains('file-drag-over')).toBe(false);
    expect(document.getElementById('inputError').hidden).toBe(true);
});

test('clearAttachments discards a late FileReader result without affecting a new drop', async () => {
    drop([new File(['old'], 'old.txt')]);
    expect(document.querySelectorAll('.attachment-chip-pending')).toHaveLength(1);
    bar.clearAttachments();
    expect(names()).toEqual([]);
    drop([new File(['new'], 'new.txt')]);
    await flush();
    expect(names()).toEqual(['new.txt']);
    expect(document.getElementById('inputError').hidden).toBe(true);
    send();
    expect(onSubmit.mock.calls[0][1]).toEqual([{ name: 'new.txt', kind: 'text', size: 3, text: 'new' }]);
});

test('drag listeners belong to the root; temporary window blur listener is removed after each drag', () => {
    const docListen = jest.spyOn(document, 'addEventListener');
    const winListen = jest.spyOn(window, 'addEventListener');
    const winRemove = jest.spyOn(window, 'removeEventListener');
    bar.clearAttachments();
    document.documentElement.innerHTML = html;
    app = document.querySelector('.app');
    bar = initInputBar({ onSubmit, onCancel: jest.fn(), getSkills: () => [], onOpenSettings: jest.fn() });
    expect(docListen.mock.calls.some(([type]) => type.startsWith('drag') || type === 'drop')).toBe(false);
    expect(winListen).not.toHaveBeenCalled();
    drag('dragenter');
    const blurHandler = winListen.mock.calls.find(([type]) => type === 'blur')[1];
    drag('dragleave');
    expect(winRemove).toHaveBeenCalledWith('blur', blurHandler);
    expect(drag('drop', document.body, { types: ['Files'], files: [] }).defaultPrevented).toBe(false);
});
