/** @jest-environment jsdom */
const { createFileStore, createMemoryFileAdapter, createIndexedDbFileAdapter } = require('../src/lib/file-store.js');

it('uses FileReader fallback in browser-like test environments without Blob.text/arrayBuffer', async () => {
    const store = createFileStore({ adapter: createMemoryFileAdapter() });
    const file = await store.saveFile(new File(['browser text'], 'browser.txt', { type: 'text/plain' }));
    expect(file).toMatchObject({ parseStatus: 'ready', persistence: 'memory' });
    expect((await store.getAttachment(file.fileId)).text).toBe('browser text');
    const image = await store.saveFile(new File([new Uint8Array([137, 80, 78, 71])], 'image.png'));
    expect((await store.readFile(image.fileId)).attachment.dataUrl).toBe('data:image/png;base64,iVBORw==');
    expect(localStorage.length).toBe(0);
});

it('does not pretend unavailable IndexedDB is persistent storage', async () => {
    const store = createFileStore({ adapter: createIndexedDbFileAdapter({ indexedDB: null }) });
    await expect(store.saveFile(new File(['not persisted'], 'file.txt'))).rejects.toThrow('unavailable');
    expect(localStorage.length).toBe(0);
});
