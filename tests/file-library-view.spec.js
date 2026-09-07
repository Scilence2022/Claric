/** @jest-environment jsdom */

const fs = require('fs');
const path = require('path');
jest.mock('../src/lib/file-store.js', () => ({
    saveFile: jest.fn(), listFiles: jest.fn(), readFile: jest.fn(), getAttachment: jest.fn(),
    renameFile: jest.fn(), deleteFile: jest.fn(),
}));
const store = require('../src/lib/file-store.js');
const { initFileLibraryView } = require('../src/taskpane/ui/file-library-view.js');
const { initInputBar } = require('../src/taskpane/ui/input-bar.js');
const flush = async () => { for (let i = 0; i < 12; i++) await Promise.resolve(); };
const meta = { fileId: 'file_1', versionId: 'version_1', name: 'notes.txt', kind: 'text', size: 4, source: 'upload' };
const click = (text) => {
    const found = [...document.querySelectorAll('#filesOverlay button')].find((node) => node.textContent === text);
    expect(found).toBeDefined();
    found.click();
};

let view;
let onAttach;
beforeEach(() => {
    jest.clearAllMocks();
    document.documentElement.innerHTML = fs.readFileSync(path.join(__dirname, '../src/taskpane/taskpane.html'), 'utf8');
    store.listFiles.mockResolvedValue({ items: [{ ...meta }], total: 1 });
    store.readFile.mockResolvedValue({ file: { ...meta }, text: 'body', offset: 0, hasMore: false });
    store.getAttachment.mockResolvedValue({ ...meta, text: 'body' });
    store.saveFile.mockResolvedValue({ ...meta });
    store.renameFile.mockResolvedValue({ ...meta });
    store.deleteFile.mockResolvedValue({ deleted: true });
    onAttach = jest.fn(() => true);
    view = initFileLibraryView({ onAttach });
});
afterEach(() => { view.close(); });

test('independent overlay lists, filters, pages and restores keyboard focus', async () => {
    const trigger = document.getElementById('filesBtn');
    trigger.focus();
    await view.open();
    expect(document.activeElement.id).toBe('filesSearch');
    expect(store.listFiles).toHaveBeenLastCalledWith({ query: '', kind: undefined, offset: 0, limit: 30 });
    const search = document.getElementById('filesSearch');
    search.value = 'notes';
    search.dispatchEvent(new Event('input'));
    await flush();
    const kind = document.getElementById('filesKind');
    kind.value = 'text';
    store.listFiles.mockResolvedValue({ items: [{ ...meta }], total: 35 });
    kind.dispatchEvent(new Event('change'));
    await flush();
    expect(store.listFiles).toHaveBeenLastCalledWith({ query: 'notes', kind: 'text', offset: 0, limit: 30 });
    click('Next files');
    await flush();
    expect(store.listFiles).toHaveBeenLastCalledWith({ query: 'notes', kind: 'text', offset: 30, limit: 30 });
    document.getElementById('filesUploadBtn').focus();
    document.activeElement.dispatchEvent(new KeyboardEvent('keydown', { key: 'Tab', shiftKey: true, bubbles: true, cancelable: true }));
    expect(document.activeElement.id).toBe('filesPreviousBtn');
    document.activeElement.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    expect(document.getElementById('filesOverlay').hidden).toBe(true);
    expect(document.activeElement).toBe(trigger);
});

test('previews extracted text safely and pages it without native binary editing', async () => {
    const unsafe = '<img src=x onerror=alert(1)>';
    store.listFiles.mockResolvedValue({ items: [{ ...meta, name: unsafe, kind: 'pdf' }], total: 1 });
    store.readFile.mockResolvedValue({ file: { ...meta, name: unsafe }, text: unsafe, offset: 0, hasMore: true });
    await view.open();
    click(unsafe);
    await flush();
    expect(store.readFile).toHaveBeenCalledWith(meta.fileId, { offset: 0, limit: 12000 });
    expect(document.querySelector('.file-library-text').textContent).toBe(unsafe);
    expect(document.querySelector('#filesOverlay img')).toBeNull();
    expect(document.getElementById('filesPreview').textContent).toContain('not edited here');
    click('Next text');
    await flush();
    expect(store.readFile).toHaveBeenLastCalledWith(meta.fileId, { offset: unsafe.length, limit: 12000 });
});

test('image preview uses a data URL and is released on close; unsafe URLs are rejected', async () => {
    store.readFile.mockResolvedValue({ file: { ...meta, kind: 'image' }, attachment: { dataUrl: 'data:image/png;base64,YQ==' } });
    await view.open();
    click(meta.name);
    await flush();
    expect(document.querySelector('.file-library-image').src).toBe('data:image/png;base64,YQ==');
    click('Close preview');
    expect(document.querySelector('.file-library-image')).toBeNull();
    store.readFile.mockResolvedValue({ file: meta, attachment: { dataUrl: 'javascript:alert(1)' } });
    click(meta.name);
    await flush();
    expect(document.querySelector('.file-library-image')).toBeNull();
    expect(document.getElementById('filesError').textContent).toContain('safely');
});

test('rename updates the library only; delete requires confirmation and Escape cancels it', async () => {
    await view.open();
    click('Rename');
    const field = document.getElementById('filesRenameInput');
    expect(document.activeElement).toBe(field);
    field.value = 'renamed.txt';
    field.closest('form').dispatchEvent(new Event('submit', { bubbles: true, cancelable: true }));
    await flush();
    expect(store.renameFile).toHaveBeenCalledWith(meta.fileId, 'renamed.txt');
    click('Delete');
    expect(store.deleteFile).not.toHaveBeenCalled();
    expect(document.activeElement.textContent).toBe('Cancel');
    document.activeElement.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    expect(document.getElementById('filesOverlay').hidden).toBe(false);
    expect(document.getElementById('filesPreview').hidden).toBe(true);
    click('Delete');
    click('Delete file');
    await flush();
    expect(store.deleteFile).toHaveBeenCalledWith(meta.fileId);
});

test('imports every picked file through the store, shows pending and individual errors', async () => {
    let resolve;
    store.saveFile.mockImplementationOnce(() => new Promise((done) => { resolve = done; }))
        .mockRejectedValueOnce(new Error('Quota exceeded'));
    await view.open();
    const files = [new File(['body'], 'a.txt'), new File(['body'], 'b.pdf')];
    const input = document.getElementById('filesInput');
    Object.defineProperty(input, 'files', { value: files });
    input.dispatchEvent(new Event('change'));
    expect(document.getElementById('filesUploadBtn').disabled).toBe(true);
    expect(document.getElementById('filesStatus').textContent).toContain('Importing a.txt');
    resolve({ ...meta });
    await flush();
    expect(store.saveFile.mock.calls.map(([file]) => file)).toEqual(files);
    expect(onAttach).not.toHaveBeenCalled();
    expect(document.getElementById('filesError').textContent).toContain('b.pdf: Quota exceeded');
    expect(document.getElementById('filesUploadBtn').disabled).toBe(false);
});

test('attach goes through getAttachment and composer validation; failed attachment keeps overlay open', async () => {
    onAttach.mockReturnValueOnce(false);
    await view.open();
    click('Attach');
    await flush();
    expect(store.getAttachment).toHaveBeenCalledWith(meta.fileId);
    expect(document.getElementById('filesOverlay').hidden).toBe(false);
    expect(document.getElementById('filesError').textContent).toContain('Could not attach');
    click('Attach');
    await flush();
    expect(onAttach).toHaveBeenCalledWith({ ...meta, text: 'body' });
    expect(document.getElementById('filesOverlay').hidden).toBe(true);
    expect(document.activeElement.id).toBe('chatInput');
});

test('real composer receives an immutable library attachment with no silent save', async () => {
    view.close();
    const onSubmit = jest.fn();
    const bar = initInputBar({ onSubmit, onCancel: jest.fn(), getSkills: () => [], onOpenSettings: jest.fn() });
    view = initFileLibraryView({ onAttach: bar.addAttachment });
    await view.open();
    click('Attach');
    await flush();
    expect(document.querySelector('.attachment-chip-name').textContent).toBe(meta.name);
    document.getElementById('sendBtn').click();
    expect(onSubmit).toHaveBeenCalledWith('', [{ ...meta, text: 'body' }]);
    expect(Object.isFrozen(onSubmit.mock.calls[0][1][0])).toBe(true);
    expect(store.saveFile).not.toHaveBeenCalled();
});

test('stale search and preview responses cannot overwrite newer state or reopen a closed view', async () => {
    let resolveList;
    store.listFiles.mockImplementationOnce(() => new Promise((done) => { resolveList = done; }));
    const opening = view.open();
    document.getElementById('filesSearch').dispatchEvent(new Event('input'));
    await flush();
    resolveList({ items: [], total: 0 });
    await opening;
    expect(document.querySelector('.file-library-name').textContent).toBe(meta.name);
    let resolveRead;
    store.readFile.mockImplementationOnce(() => new Promise((done) => { resolveRead = done; }));
    click(meta.name);
    view.close();
    resolveRead({ file: meta, text: 'late' });
    await flush();
    expect(document.getElementById('filesOverlay').hidden).toBe(true);
    expect(document.getElementById('filesPreview').textContent).toBe('');
});

test('store failures are visible without optimistic deletion', async () => {
    await view.open();
    store.deleteFile.mockRejectedValueOnce(new Error('Storage unavailable'));
    click('Delete');
    click('Delete file');
    await flush();
    expect(document.querySelector('.file-library-name').textContent).toBe(meta.name);
    expect(document.getElementById('filesError').textContent).toBe('Storage unavailable');
    view.close();
    store.listFiles.mockRejectedValueOnce(new Error('Database blocked'));
    await view.open();
    expect(document.getElementById('filesError').textContent).toBe('Database blocked');
});

test('actual store and parser support import, preview, rename, attach, and delete through UI', async () => {
    const { createFileStore, createMemoryFileAdapter } = jest.requireActual('../src/lib/file-store.js');
    const actual = createFileStore({ adapter: createMemoryFileAdapter() });
    view.close();
    document.documentElement.innerHTML = fs.readFileSync(path.join(__dirname, '../src/taskpane/taskpane.html'), 'utf8');
    for (const key of Object.keys(store)) store[key].mockImplementation(actual[key]);
    const inputBar = initInputBar({ onSubmit: jest.fn(), onCancel: jest.fn(), getSkills: () => [], onOpenSettings: jest.fn() });
    view = initFileLibraryView({ onAttach: inputBar.addAttachment });
    await view.open();
    const input = document.getElementById('filesInput');
    Object.defineProperty(input, 'files', { value: [new File(['actual body'], 'actual.txt', { type: 'text/plain' })] });
    input.dispatchEvent(new Event('change'));
    await new Promise((resolve) => setTimeout(resolve, 40));
    const { items } = await actual.listFiles();
    expect(items).toHaveLength(1);
    expect(items[0].parseStatus).toBe('ready');
    click('actual.txt');
    await flush();
    expect(document.querySelector('.file-library-text').textContent).toBe('actual body');
    click('Rename');
    const field = document.getElementById('filesRenameInput');
    field.value = 'new.txt';
    field.closest('form').dispatchEvent(new Event('submit', { bubbles: true, cancelable: true }));
    await flush();
    expect((await actual.getFile(items[0].fileId)).name).toBe('new.txt');
    click('Attach');
    await flush();
    expect(document.querySelector('.attachment-chip-name').textContent).toBe('new.txt');
    await view.open();
    click('Delete');
    click('Delete file');
    await flush();
    expect((await actual.listFiles()).total).toBe(0);
    expect(document.querySelector('.attachment-chip-name').textContent).toBe('new.txt');
});
