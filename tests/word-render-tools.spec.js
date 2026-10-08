/** @jest-environment jsdom */
jest.mock('pdfjs-dist/legacy/build/pdf.min.mjs', () => ({ GlobalWorkerOptions: {}, getDocument: jest.fn() }));
const pdfjs = require('pdfjs-dist/legacy/build/pdf.min.mjs');
const { canReadWordRendering, exportWordPdf, createWordRenderTools, sendWithWordVisuals } = require('../src/taskpane/word-render-tools.js');

const bytes = Uint8Array.from(Array.from('%PDF-test-data', (char) => char.charCodeAt(0)));
const success = (value) => ({ status: 'succeeded', value });
function office() {
    const file = {
        size: bytes.length, sliceCount: 2,
        getSliceAsync: jest.fn((index, callback) => callback(success({ index, data: [...(index ? bytes.slice(5) : bytes.slice(0, 5))] }))),
        closeAsync: jest.fn((callback) => callback?.(success())),
    };
    global.Office = { AsyncResultStatus: { Succeeded: 'succeeded' }, FileType: { Pdf: 'pdf' },
        context: { platform: 'Mac', document: { getFileAsync: jest.fn((_type, _options, callback) => callback(success(file))) } } };
    return file;
}
function rendererWorld(options = {}) {
    const texts = ['Example introduction', 'Example discussion and results', 'Example conclusion'];
    const pdf = { numPages: texts.length, getPage: jest.fn(async (number) => ({
        number, getTextContent: jest.fn(async () => ({ items: [{ str: texts[number - 1] }] })),
    })) };
    const task = { promise: Promise.resolve(pdf), destroy: jest.fn(async () => {}) };
    const exportPdf = jest.fn(async () => bytes);
    const loadPdf = jest.fn(async () => task);
    const render = jest.fn(async (page) => ({ dataUrl: `data:image/jpeg;base64,page${page.number}`, width: 1200, height: 1600 }));
    const renderer = createWordRenderTools({ exportPdf, loadPdf, render, scopeText: 'Example discussion', ...options });
    return { pdf, task, exportPdf, loadPdf, render, renderer };
}
afterEach(() => { delete global.Office; jest.useRealTimers(); });

test.each(['Mac', 'PC', 'iOS', 'OfficeOnline', 'unknown'])('native export availability is honest on %s', (platform) => {
    office(); Office.context.platform = platform;
    expect(canReadWordRendering()).toBe(['Mac', 'PC', 'iOS'].includes(platform));
});

test('exports only native PDF bytes, reads slices in order and closes the handle', async () => {
    const file = office();
    expect(await exportWordPdf()).toEqual(bytes);
    expect(Office.context.document.getFileAsync).toHaveBeenCalledWith('pdf', { sliceSize: 65536 }, expect.any(Function));
    expect(file.getSliceAsync.mock.calls.map((call) => call[0])).toEqual([0, 1]);
    expect(file.closeAsync).toHaveBeenCalledTimes(1);
});

test.each(['unsupported', 'native error', 'oversize', 'wrong slice', 'missing slice', 'incomplete', 'not PDF', 'bad bytes'])('refuses %s without inventing visual evidence', async (kind) => {
    const file = office();
    if (kind === 'unsupported') Office.context.platform = 'OfficeOnline';
    if (kind === 'native error') Office.context.document.getFileAsync.mockImplementation((_t, _o, cb) => cb({ status: 'failed', error: { message: 'GeneralException' } }));
    if (kind === 'oversize') file.size = 21 * 1024 * 1024;
    if (kind === 'wrong slice') file.getSliceAsync.mockImplementation((_i, cb) => cb(success({ index: 99, data: [1] })));
    if (kind === 'missing slice') file.getSliceAsync.mockImplementation((_i, cb) => cb({ status: 'failed', error: { message: 'Slice failed' } }));
    if (kind === 'incomplete') file.sliceCount = 1;
    if (kind === 'not PDF') file.getSliceAsync.mockImplementation((index, cb) => cb(success({ index, data: [...(index ? bytes.slice(5) : bytes.slice(0, 5))].map(() => 1) })));
    if (kind === 'bad bytes') file.getSliceAsync.mockImplementation((index, cb) => cb(success({ index, data: [999] })));
    await expect(exportWordPdf()).rejects.toThrow();
    expect(file.closeAsync).toHaveBeenCalledTimes(['unsupported', 'native error'].includes(kind) ? 0 : 1);
});

test('closes a file returned after export cancellation', async () => {
    const file = office(); const controller = new AbortController(); let callback;
    Office.context.document.getFileAsync.mockImplementation((_t, _o, cb) => { callback = cb; });
    const exporting = exportWordPdf({ signal: controller.signal });
    const rejected = expect(exporting).rejects.toMatchObject({ name: 'AbortError' });
    controller.abort(); await rejected;
    callback(success(file));
    expect(file.closeAsync).toHaveBeenCalledTimes(1);
});

test('export timeout is bounded and a late handle is still closed', async () => {
    jest.useFakeTimers();
    const file = office(); let callback;
    Office.context.document.getFileAsync.mockImplementation((_t, _o, cb) => { callback = cb; });
    const exporting = exportWordPdf();
    const rejected = expect(exporting).rejects.toThrow(/timed out/);
    await jest.advanceTimersByTimeAsync(30000); await rejected;
    callback(success(file));
    expect(file.closeAsync).toHaveBeenCalledTimes(1);
});

test('handle cleanup failure preserves successfully captured bytes', async () => {
    const file = office(); const log = jest.fn();
    file.closeAsync.mockImplementation((callback) => callback({ status: 'failed', error: { message: 'Close failed' } }));
    expect(await exportWordPdf({ log })).toEqual(bytes);
    expect(log).toHaveBeenCalledWith('Word PDF handle cleanup failed: Close failed', 'warning');
});

test('lists searchable physical pages, attaches real rendered images and disposes the snapshot', async () => {
    const w = rendererWorld();
    const listed = await w.renderer.execute('list_rendered_pages', { search: 'discussion' });
    expect(listed).toMatchObject({ ok: true, result: { source: 'word_native_pdf', totalPages: 3,
        pages: [{ pageNumber: 2, matchesScopeText: true }], visualInputAvailable: false } });
    expect(listed.attachments).toBeUndefined();
    const read = await w.renderer.execute('read_rendered_pages', { pages: [2] });
    expect(read.attachments).toEqual([{ dataUrl: 'data:image/jpeg;base64,page2' }]);
    expect(read.result).toMatchObject({ layout: 'print', pages: [{ pageNumber: 2, text: 'Example discussion and results', width: 1200, height: 1600 }], visualInputAvailable: true });
    expect(w.renderer.status().inspectedPages).toEqual([2]);
    expect(w.exportPdf).toHaveBeenCalledTimes(1);
    await w.renderer.dispose();
    expect(w.task.destroy).toHaveBeenCalledTimes(1);
    expect((await w.renderer.execute('read_rendered_pages', { pages: [1] })).ok).toBe(false);
});

test('native PDF loading uses locally bundled fonts, CJK mappings and decoders', async () => {
    const w = rendererWorld();
    pdfjs.getDocument.mockReturnValue(w.task);
    const reader = createWordRenderTools({ exportPdf: w.exportPdf, render: w.render });
    expect((await reader.execute('read_rendered_pages', { pages: [1] })).ok).toBe(true);
    expect(pdfjs.getDocument).toHaveBeenCalledWith({ data: bytes,
        cMapUrl: new URL('pdfjs/cmaps/', document.baseURI).toString(), cMapPacked: true,
        standardFontDataUrl: new URL('pdfjs/standard_fonts/', document.baseURI).toString(),
        wasmUrl: new URL('pdfjs/wasm/', document.baseURI).toString(),
    });
    await reader.dispose();
});

test.each(['render', 'text', 'cancel'])('multi-page %s failure never records unreturned images', async (kind) => {
    const controller = new AbortController(); const w = rendererWorld({ signal: controller.signal });
    if (kind === 'render') w.render.mockResolvedValueOnce({ dataUrl: 'data:image/jpeg;base64,first' }).mockRejectedValueOnce(new Error('Render failed'));
    if (kind === 'text') w.pdf.getPage.mockImplementation(async (number) => ({ number,
        getTextContent: async () => { if (number === 2) throw new Error('Text read failed'); return { items: [] }; },
    }));
    if (kind === 'cancel') w.render.mockImplementation(async (page) => {
        if (page.number === 2) controller.abort(); return { dataUrl: 'data:image/jpeg;base64,page' };
    });
    if (kind === 'cancel') await expect(w.renderer.execute('read_rendered_pages', { pages: [1, 2] })).rejects.toMatchObject({ name: 'AbortError' });
    else {
        const result = await w.renderer.execute('read_rendered_pages', { pages: [1, 2] });
        expect(result).toMatchObject({ ok: false, visualInputAvailable: false });
        expect(result.attachments).toBeUndefined();
        expect(w.renderer.status().unavailable).toMatch(/failed/);
    }
    expect(w.renderer.status().inspectedPages).toEqual([]);
    await w.renderer.dispose();
});

test('page listing is bounded and refresh discards the old native snapshot', async () => {
    const w = rendererWorld();
    const listed = await w.renderer.execute('list_rendered_pages', { limit: 1 });
    expect(listed.result.nextPage).toBe(2);
    const next = await w.renderer.execute('list_rendered_pages', { startPage: 2, limit: 1 });
    expect(next.result.pages[0].pageNumber).toBe(2);
    await w.renderer.execute('list_rendered_pages', { refresh: true });
    expect(w.exportPdf).toHaveBeenCalledTimes(2);
    expect(w.task.destroy).toHaveBeenCalledTimes(1);
    await w.renderer.dispose();
});

test.each([['other', {}], ['read_rendered_pages', {}], ['read_rendered_pages', { pages: [1, 1] }],
    ['read_rendered_pages', { pages: [0] }], ['read_rendered_pages', { pages: [1, 2, 3, 4] }],
    ['list_rendered_pages', { search: 1 }], ['list_rendered_pages', { startPage: 0 }], ['list_rendered_pages', { limit: 51 }]])('invalid arguments do not export a PDF or bypass visual checks: %s', async (name, args) => {
    const w = rendererWorld();
    expect((await w.renderer.execute(name, args)).ok).toBe(false);
    expect(w.exportPdf).not.toHaveBeenCalled();
    expect(w.renderer.status().unavailable).toBeNull();
});

test('out-of-range pages cannot masquerade as rendering unavailability', async () => {
    const w = rendererWorld();
    expect((await w.renderer.execute('read_rendered_pages', { pages: [4] })).ok).toBe(false);
    expect(w.render).not.toHaveBeenCalled();
    expect(w.renderer.status().unavailable).toBeNull();
    await w.renderer.dispose();
});

test('failed or oversized native snapshots produce explicit errors with no attachments', async () => {
    const w = rendererWorld();
    w.pdf.numPages = 501;
    const result = await w.renderer.execute('list_rendered_pages');
    expect(result).toMatchObject({ ok: false, visualInputAvailable: false });
    expect(w.renderer.status().unavailable).toMatch(/500-page/);
    expect((await w.renderer.execute('read_rendered_pages', { pages: [1] })).ok).toBe(false);
    expect(w.render).not.toHaveBeenCalled();
    await w.renderer.dispose();
});

test('cancellation after a render cannot return or record unobserved images', async () => {
    const controller = new AbortController(); const w = rendererWorld({ signal: controller.signal });
    w.render.mockImplementation(async () => { controller.abort(); return { dataUrl: 'data:image/jpeg;base64,x' }; });
    await expect(w.renderer.execute('read_rendered_pages', { pages: [1] })).rejects.toMatchObject({ name: 'AbortError' });
    expect(w.renderer.status().inspectedPages).toEqual([]);
    await w.renderer.dispose();
});

test('vision rejection removes bytes and corrects visual availability before a text-only retry', async () => {
    const messages = [{ role: 'user', content: [
        { type: 'text', text: JSON.stringify({ ok: true, result: { source: 'word_native_pdf', visualInputAvailable: true } }) },
        { type: 'image_url', image_url: { url: 'data:image/jpeg;base64,private' } },
    ] }];
    const send = jest.fn().mockRejectedValueOnce(new Error('HTTP 400: vision unsupported')).mockResolvedValue('done');
    const onUnavailable = jest.fn();
    expect(await sendWithWordVisuals(send, messages, onUnavailable)).toBe('done');
    const retry = send.mock.calls[1][0];
    expect(JSON.stringify(retry)).not.toContain('base64');
    expect(JSON.parse(retry[0].content[1].text).result.visualInputAvailable).toBe(false);
    expect(onUnavailable).toHaveBeenCalledTimes(1);
    expect(messages[0].content).toHaveLength(2);
    await sendWithWordVisuals(send, messages, onUnavailable, { textOnly: true });
    expect(JSON.stringify(send.mock.calls[2][0])).not.toContain('image_url');
});

test.each(['HTTP 401: unauthorized', 'HTTP 403: forbidden', 'HTTP 504: upstream timeout', 'cancelled'])('does not hide unrelated model errors: %s', async (message) => {
    const send = jest.fn(async () => { throw new Error(message); });
    await expect(sendWithWordVisuals(send, [{ content: [{ type: 'image_url', image_url: { url: 'x' } }] }])).rejects.toThrow(message);
    expect(send).toHaveBeenCalledTimes(1);
});
