const { execFileSync } = require('node:child_process');
const path = require('node:path');
const { readPdfPageText } = require('../src/lib/pdf-text-reader.js');

function world(chunks) {
    const read = jest.fn();
    chunks.forEach((items) => read.mockResolvedValueOnce({ done: false, value: { items } }));
    read.mockResolvedValue({ done: true });
    const reader = { read, cancel: jest.fn(async () => {}), releaseLock: jest.fn() };
    const stream = { getReader: jest.fn(() => reader) };
    return { reader, stream, page: { streamTextContent: jest.fn(() => stream) } };
}

test('reads reader-only PDF streams and skips nontext items without async iteration', async () => {
    const w = world([[{ str: 'Hello', hasEOL: true }, { type: 'beginMarkedContent' }], [], [{ str: 'world' }]]);
    expect(w.stream[Symbol.asyncIterator]).toBeUndefined();
    expect(await readPdfPageText(w.page)).toEqual({ text: 'Hello world', truncated: false });
    expect(w.reader.releaseLock).toHaveBeenCalledTimes(1);
    expect(w.reader.cancel).not.toHaveBeenCalled();
    const attachment = world([[{ str: 'Hello', hasEOL: true }], [{ str: 'wor' }, { str: 'ld' }]]);
    expect(await readPdfPageText(attachment.page, { separator: '', preserveEOL: true }))
        .toEqual({ text: 'Hello\nworld', truncated: false });
});

test('bounds retained text, cancels early and explicitly identifies truncation', async () => {
    const w = world([[{ str: 'abcdef' }], [{ str: 'unread' }]]);
    expect(await readPdfPageText(w.page, { maxChars: 3 })).toEqual({ text: 'abc', truncated: true });
    expect(w.reader.read).toHaveBeenCalledTimes(1);
    expect(w.reader.cancel).toHaveBeenCalledTimes(1);
    expect(w.reader.releaseLock).toHaveBeenCalledTimes(1);
    const exact = world([[{ str: 'abc' }], [{ type: 'endMarkedContent' }]]);
    expect(await readPdfPageText(exact.page, { maxChars: 3 })).toEqual({ text: 'abc', truncated: false });
    const empty = world([[{ str: 'abc' }]]);
    expect(await readPdfPageText(empty.page, { maxChars: 0 })).toEqual({ text: '', truncated: true });
});

test('strict extraction limits fail explicitly and release the stream without replacing errors', async () => {
    const w = world([[{ str: 'abc' }]]);
    w.reader.cancel.mockRejectedValue(new Error('Cancel failed'));
    w.reader.releaseLock.mockImplementation(() => { throw new Error('Release failed'); });
    await expect(readPdfPageText(w.page, { maxChars: 2, onLimit: 'error' })).rejects.toThrow('2 character limit');
    expect(w.reader.cancel).toHaveBeenCalledTimes(1);
    expect(w.reader.releaseLock).toHaveBeenCalledTimes(1);
});

test.each(['read', 'invalid chunk'])('%s failures release and cancel the reader', async (kind) => {
    const w = world([]);
    if (kind === 'read') w.reader.read.mockRejectedValue(new Error('Read failed'));
    else w.reader.read.mockResolvedValue({ done: false, value: {} });
    await expect(readPdfPageText(w.page)).rejects.toThrow(kind === 'read' ? 'Read failed' : 'invalid items');
    expect(w.reader.cancel).toHaveBeenCalledTimes(1);
    expect(w.reader.releaseLock).toHaveBeenCalledTimes(1);
});

test('cleanup failures preserve complete extracted text', async () => {
    const w = world([[{ str: 'body' }]]);
    w.reader.releaseLock.mockImplementation(() => { throw new Error('Release failed'); });
    expect(await readPdfPageText(w.page)).toEqual({ text: 'body', truncated: false });
});

test('pre-aborted reads never acquire a reader', async () => {
    const w = world([]); const controller = new AbortController(); controller.abort();
    await expect(readPdfPageText(w.page, { signal: controller.signal })).rejects.toMatchObject({ name: 'AbortError' });
    expect(w.page.streamTextContent).not.toHaveBeenCalled();
});

test.each(['normal', 'sync cancel failure'])('aborting a pending read rejects promptly with %s cleanup', async (kind) => {
    const w = world([]); const controller = new AbortController();
    w.reader.read.mockImplementation(() => new Promise(() => {}));
    if (kind === 'sync cancel failure') w.reader.cancel.mockImplementation(() => { throw new Error('Cancel failed'); });
    const pending = readPdfPageText(w.page, { signal: controller.signal });
    const rejected = expect(pending).rejects.toMatchObject({ name: 'AbortError' });
    controller.abort();
    await rejected;
    expect(w.reader.cancel).toHaveBeenCalledTimes(1);
    expect(w.reader.releaseLock).toHaveBeenCalledTimes(1);
});

test.each([{ maxChars: -1 }, { maxChars: Infinity }, { maxChars: 0.5 }, { onLimit: 'silent' }])('rejects invalid limits before reading: %j', async (options) => {
    const w = world([]);
    await expect(readPdfPageText(w.page, options)).rejects.toThrow('Invalid PDF text limit');
    expect(w.page.streamTextContent).not.toHaveBeenCalled();
});

test('installed PDF.js reads a real PDF when ReadableStream async iteration is unavailable', () => {
    // Run ESM in a separate Node process: Jest mocks/transforms cannot conceal
    // a dependency runtime incompatibility, and the test never patches the app.
    const script = String.raw`
        import * as pdfjs from 'pdfjs-dist/legacy/build/pdf.mjs';
        import { fileURLToPath } from 'node:url';
        import { readPdfPageText } from './src/lib/pdf-text-reader.js';
        const contentStream = 'BT /F1 12 Tf 20 100 Td (Hello) Tj ET';
        const objects = [
            '<< /Type /Catalog /Pages 2 0 R >>',
            '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
            '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>',
            '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>',
            '<< /Length ' + contentStream.length + ' >>\nstream\n' + contentStream + '\nendstream',
        ];
        let content = '%PDF-1.4\n'; const offsets = [];
        objects.forEach((body, i) => {
            offsets.push(content.length);
            content += (i + 1) + ' 0 obj\n' + body + '\nendobj\n';
        });
        const start = content.length;
        content += 'xref\n0 6\n0000000000 65535 f \n'
            + offsets.map(o => String(o).padStart(10, '0') + ' 00000 n \n').join('')
            + 'trailer\n<< /Size 6 /Root 1 0 R >>\nstartxref\n' + start + '\n%%EOF\n';
        const task = pdfjs.getDocument({ data: new TextEncoder().encode(content),
            standardFontDataUrl: fileURLToPath(new URL('./node_modules/pdfjs-dist/standard_fonts/', import.meta.url)) });
        const descriptor = Object.getOwnPropertyDescriptor(ReadableStream.prototype, Symbol.asyncIterator);
        try {
            const pdf = await task.promise;
            const page = await pdf.getPage(1);
            delete ReadableStream.prototype[Symbol.asyncIterator];
            let nativeError = null;
            try { await page.getTextContent(); } catch (error) { nativeError = error.message; }
            const portable = await readPdfPageText(page);
            console.log(JSON.stringify({ pages: pdf.numPages, nativeError, portable,
                asyncIterator: typeof ReadableStream.prototype[Symbol.asyncIterator] }));
        } finally {
            if (descriptor) Object.defineProperty(ReadableStream.prototype, Symbol.asyncIterator, descriptor);
            await task.destroy();
        }
    `;
    const output = execFileSync(process.execPath, ['--input-type=module'], {
        input: script, cwd: path.resolve(__dirname, '..'), encoding: 'utf8', timeout: 15000,
    });
    const result = JSON.parse(output.trim());
    expect(result).toMatchObject({ pages: 1, asyncIterator: 'undefined', portable: { text: 'Hello', truncated: false } });
    expect(result.nativeError).toMatch(/not async iterable/);
});
