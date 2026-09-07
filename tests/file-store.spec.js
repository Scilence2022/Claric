jest.mock('mammoth/mammoth.browser.min.js', () => ({ extractRawText: jest.fn(async () => ({ value: 'DOCX extracted body' })) }));
jest.mock('pdfjs-dist/legacy/build/pdf.min.mjs', () => ({ GlobalWorkerOptions: {}, getDocument: jest.fn() }));
const pdfjs = require('pdfjs-dist/legacy/build/pdf.min.mjs');
const { createFileStore, createMemoryFileAdapter, createIndexedDbFileAdapter, FILE_STORE_LIMITS } = require('../src/lib/file-store.js');

function memoryStore(options = {}) { return createFileStore({ adapter: createMemoryFileAdapter(), ...options }); }

it('persists original Blob and cached extracted text, never content in metadata; rename retains content version', async () => {
    const adapter = createMemoryFileAdapter();
    const store = createFileStore({ adapter });
    const file = await store.saveFile(new File(['private body'], 'notes.txt'));
    expect(file).toMatchObject({ kind: 'text', parseStatus: 'ready', persistence: 'memory', textLength: 12 });
    expect(file.fileId).toMatch(/^file_/);
    expect(JSON.stringify(file)).not.toContain('private body');
    const record = await adapter.get(file.fileId);
    expect(record.body).toBeInstanceOf(Blob);
    expect(await record.body.text()).toBe('private body');
    expect(record.text).toBe('private body');
    const renamed = await store.renameFile(file.fileId, 'renamed.pdf');
    expect(renamed).toMatchObject({ versionId: file.versionId, kind: 'text', name: 'renamed.pdf' });
    expect(await store.getAttachment(file.fileId)).toMatchObject({ fileId: file.fileId, versionId: file.versionId, text: 'private body', source: 'upload' });
    await expect(store.readFile(file.fileId, { versionId: 'wrong' })).rejects.toThrow('version');
    await store.deleteFile(file.fileId);
    await expect(store.getFile(file.fileId)).rejects.toThrow('deleted');
    await expect(store.deleteFile(file.fileId)).rejects.toThrow('deleted');
});

it('uses mammoth and PDF extraction instead of decoding binary as text', async () => {
    const store = memoryStore();
    const docx = await store.saveFile(new File(['zip bytes'], 'report.docx'));
    expect((await store.readFile(docx.fileId)).text).toBe('DOCX extracted body');
    pdfjs.getDocument.mockReturnValue({
        promise: Promise.resolve({ numPages: 1, getPage: async () => ({ streamTextContent: () => ({ getReader: () => ({
            read: jest.fn().mockResolvedValueOnce({ done: false, value: { items: [{ str: 'PDF extracted body' }] } }).mockResolvedValue({ done: true }),
            releaseLock: jest.fn(),
        }) }) }) }),
        destroy: jest.fn(),
    });
    const pdf = await store.saveFile(new File(['%PDF binary'], 'report.pdf'));
    expect((await store.readFile(pdf.fileId)).text).toBe('PDF extracted body');
});

it('records empty scanned PDFs and parse errors without falsely offering content', async () => {
    const store = memoryStore({ parse: async () => ({ text: '' }) });
    const pdf = await store.saveFile(new File(['pdf'], 'scan.pdf'));
    expect(pdf).toMatchObject({ parseStatus: 'empty', textLength: 0, parseError: expect.stringContaining('OCR') });
    expect(await store.readFile(pdf.fileId)).toMatchObject({ text: '', hasMore: false });
    const broken = memoryStore({ parse: async () => { throw new Error('corrupt archive'); } });
    const docx = await broken.saveFile(new File(['bad'], 'bad.docx'));
    expect(docx.parseStatus).toBe('error');
    await expect(broken.getAttachment(docx.fileId)).rejects.toThrow('corrupt archive');
});

it('bounds reads, attachment context, names, queries, file size and library size', async () => {
    const store = memoryStore();
    const item = await store.saveFile(new File(['a'.repeat(210000)], 'large.txt'));
    expect((await store.readFile(item.fileId, { limit: 1e9 })).text).toHaveLength(12000);
    expect((await store.readFile(item.fileId, { offset: 12000, limit: 2 })).nextOffset).toBe(12002);
    expect(await store.getAttachment(item.fileId)).toMatchObject({ truncated: true });
    expect((await store.getAttachment(item.fileId)).text).toHaveLength(200000);
    for (const offset of [-1, Infinity, NaN]) await expect(store.readFile(item.fileId, { offset })).rejects.toThrow('paging');
    await expect(store.saveFile(new File(['x'], '../'.repeat(100) + '.txt'))).rejects.toThrow('name');
    await expect(store.renameFile(item.fileId, '')).rejects.toThrow('name');
    await expect(store.listFiles({ query: 'q'.repeat(201) })).rejects.toThrow('query');
    await expect(store.listFiles({ kind: 'exe' })).rejects.toThrow('kind');
    await expect(store.saveFile(new File([new Uint8Array(10 * 1024 * 1024 + 1)], 'huge.txt'))).rejects.toThrow('limit');
    const adapter = createMemoryFileAdapter();
    for (let i = 0; i < FILE_STORE_LIMITS.MAX_FILES; i++) await adapter.add({ fileId: `file_${i}`, size: 1 });
    await expect(adapter.add({ fileId: 'overflow', size: 1 })).rejects.toThrow('full');
});

it('reports extraction size overflow and rejects unsupported image types', async () => {
    const store = memoryStore({ parse: async () => ({ text: 'x'.repeat(FILE_STORE_LIMITS.MAX_TEXT_CHARS + 1) }) });
    const file = await store.saveFile(new File(['x'], 'long.txt'));
    expect(file.parseStatus).toBe('error');
    await expect(store.readFile(file.fileId)).rejects.toThrow('character limit');
    await expect(store.saveFile(new File(['<svg>'], 'attack.svg', { type: 'image/svg+xml' }))).rejects.toThrow('Only PNG');
});

it('searches names only, clamps pagination and restricts to allowed IDs before paging', async () => {
    const store = memoryStore();
    const a = await store.saveFile(new File(['secret'], 'alpha.txt'));
    await store.saveFile(new File(['alpha'], 'beta.txt'));
    expect((await store.searchFiles('ALPHA')).total).toBe(1);
    const page = await store.listFiles({ allowedIds: [a.fileId], limit: 999 });
    expect(page).toMatchObject({ total: 1, limit: 50, hasMore: false });
    expect(page.items[0].fileId).toBe(a.fileId);
    expect(JSON.stringify(page)).not.toContain('secret');
    expect((await store.listFiles({ offset: 99 })).items).toEqual([]);
});

it('produces safe bounded image attachments and corrects missing MIME from extensions', async () => {
    const store = memoryStore();
    const file = await store.saveFile(new File([new Uint8Array([137, 80, 78, 71])], 'pic.png'));
    const result = await store.readFile(file.fileId);
    expect(result.attachment.dataUrl).toBe('data:image/png;base64,iVBORw==');
    expect(JSON.stringify(await store.getFile(file.fileId))).not.toContain('base64');
    const unsafe = memoryStore({ parse: async () => ({ dataUrl: 'https://evil.test/image' }) });
    const invalid = await unsafe.saveFile(new File(['x'], 'unsafe.png'));
    await expect(unsafe.readFile(invalid.fileId)).rejects.toThrow('unsafe');
});

it('explicitly fails unavailable persistence and propagates all storage errors', async () => {
    const unavailable = createFileStore({ adapter: createIndexedDbFileAdapter({ indexedDB: null }) });
    await expect(unavailable.saveFile(new File(['x'], 'x.txt'))).rejects.toThrow('unavailable');
    await expect(unavailable.listFiles()).rejects.toThrow('unavailable');
    const adapter = createMemoryFileAdapter();
    const store = createFileStore({ adapter });
    const file = await store.saveFile(new File(['x'], 'x.txt'));
    adapter.update = async () => { throw new Error('disk failure'); };
    adapter.add = async () => { throw new Error('quota exceeded'); };
    await expect(store.deleteFile(file.fileId)).rejects.toThrow('disk failure');
    await expect(store.renameFile(file.fileId, 'new.txt')).rejects.toThrow('disk failure');
    await expect(store.saveFile(new File(['x'], 'x.txt'))).rejects.toThrow('quota');
});

it('does not resolve an IndexedDB request until transaction completion and rejects late aborts', async () => {
    let tx;
    let request;
    const db = { close: jest.fn(), transaction: () => {
        tx = { objectStore: () => ({ get: () => { request = {}; return request; } }), abort: () => tx.onabort() };
        return tx;
    } };
    const factory = { open: () => {
        const req = { result: db };
        queueMicrotask(() => req.onsuccess());
        return req;
    } };
    const adapter = createIndexedDbFileAdapter({ indexedDB: factory });
    let resolved = false;
    const pending = adapter.get('file_1').then(() => { resolved = true; });
    await new Promise((resolve) => setImmediate(resolve));
    request.result = { fileId: 'file_1' };
    request.onsuccess();
    await Promise.resolve();
    expect(resolved).toBe(false);
    tx.oncomplete();
    await pending;
    expect(resolved).toBe(true);
    const aborted = adapter.get('file_1');
    await new Promise((resolve) => setImmediate(resolve));
    request.onsuccess();
    tx.error = new Error('late transaction failure');
    tx.onabort();
    await expect(aborted).rejects.toThrow('late transaction failure');
});
