/**
 * Specs for lib/file-attachments.js: type detection, size/count validation,
 * parsing (text/image/docx/pdf — mammoth and pdf.js mocked), prompt context
 * assembly, and persistence metadata. Runs in node env: File/btoa are Node
 * 22 globals, so no jsdom is needed.
 */

jest.mock('mammoth/mammoth.browser.min.js', () => ({
    extractRawText: jest.fn(async () => ({ value: '  extracted docx text  ' })),
}));

jest.mock('pdfjs-dist/legacy/build/pdf.min.mjs', () => ({
    GlobalWorkerOptions: {},
    getDocument: jest.fn(),
}));

const mammoth = require('mammoth/mammoth.browser.min.js');
const pdfjs = require('pdfjs-dist/legacy/build/pdf.min.mjs');

const {
    ATTACHMENT_KIND,
    ATTACHMENT_LIMITS,
    detectAttachmentKind,
    validateAttachment,
    formatBytes,
    parseAttachment,
    buildAttachmentContext,
    splitAttachments,
    attachmentMeta,
} = require('../src/lib/file-attachments.js');

describe('detectAttachmentKind', () => {
    test('classifies by extension first', () => {
        expect(detectAttachmentKind('notes.txt')).toBe('text');
        expect(detectAttachmentKind('README.MD')).toBe('text');
        expect(detectAttachmentKind('data.csv')).toBe('text');
        expect(detectAttachmentKind('config.json')).toBe('text');
        expect(detectAttachmentKind('app.log')).toBe('text');
        expect(detectAttachmentKind('photo.PNG')).toBe('image');
        expect(detectAttachmentKind('pic.jpg')).toBe('image');
        expect(detectAttachmentKind('pic.jpeg')).toBe('image');
        expect(detectAttachmentKind('anim.gif')).toBe('image');
        expect(detectAttachmentKind('modern.webp')).toBe('image');
        expect(detectAttachmentKind('contract.docx')).toBe('docx');
        expect(detectAttachmentKind('paper.pdf')).toBe('pdf');
    });

    test('extension wins over a generic MIME type', () => {
        expect(detectAttachmentKind('notes.md', 'application/octet-stream')).toBe('text');
    });

    test('falls back to MIME when the extension is unknown', () => {
        expect(detectAttachmentKind('blob', 'application/pdf')).toBe('pdf');
        expect(detectAttachmentKind('blob', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document')).toBe('docx');
        expect(detectAttachmentKind('blob', 'image/png')).toBe('image');
        expect(detectAttachmentKind('blob', 'text/plain')).toBe('text');
        expect(detectAttachmentKind('blob', 'application/json')).toBe('text');
    });

    test('returns null for unsupported types and junk input', () => {
        expect(detectAttachmentKind('archive.zip', 'application/zip')).toBeNull();
        expect(detectAttachmentKind('script.exe')).toBeNull();
        expect(detectAttachmentKind('')).toBeNull();
        expect(detectAttachmentKind(null)).toBeNull();
    });
});

describe('validateAttachment', () => {
    test('accepts a well-formed text file', () => {
        expect(validateAttachment({ name: 'a.txt', size: 100, kind: 'text' }, [])).toEqual({ ok: true });
    });

    test('rejects unsupported types', () => {
        const v = validateAttachment({ name: 'a.zip', size: 100, kind: null }, []);
        expect(v.ok).toBe(false);
        expect(v.error).toContain('unsupported file type');
    });

    test.each(['text', 'pdf', 'docx'])('%s accepts 10 MiB and rejects one byte over', (kind) => {
        const file = { name: `big.${kind}`, size: 10 * 1024 * 1024, kind };
        expect(validateAttachment(file)).toEqual({ ok: true });
        expect(validateAttachment({ ...file, size: file.size + 1 })).toEqual({
            ok: false,
            error: expect.stringContaining('per-file limit'),
        });
        expect(validateAttachment(file, [{ name: 'other.txt', size: 1 }])).toEqual({
            ok: false,
            error: expect.stringContaining('total'),
        });
    });

    test('keeps the image cap at 4.5 MiB', () => {
        const file = { name: 'big.png', size: 4.5 * 1024 * 1024, kind: 'image' };
        expect(validateAttachment(file)).toEqual({ ok: true });
        expect(validateAttachment({ ...file, size: file.size + 1 })).toEqual({
            ok: false,
            error: expect.stringContaining('per-file limit'),
        });
        expect(validateAttachment({ ...file, size: ATTACHMENT_LIMITS.MAX_TEXT_FILE_BYTES }).ok).toBe(false);
    });

    test('rejects beyond the max file count', () => {
        const existing = Array.from({ length: ATTACHMENT_LIMITS.MAX_FILES }, (_, i) => ({ name: `f${i}.txt`, size: 1 }));
        const v = validateAttachment({ name: 'extra.txt', size: 1, kind: 'text' }, existing);
        expect(v.ok).toBe(false);
        expect(v.error).toContain(`${ATTACHMENT_LIMITS.MAX_FILES} attachments`);
    });

    test('accepts the exact aggregate cap and rejects one byte over', () => {
        const existing = [{ name: 'a.txt', size: ATTACHMENT_LIMITS.MAX_TOTAL_BYTES - 10 }];
        expect(validateAttachment({ name: 'b.txt', size: 10, kind: 'text' }, existing)).toEqual({ ok: true });
        const v = validateAttachment({ name: 'b.txt', size: 11, kind: 'text' }, existing);
        expect(v.ok).toBe(false);
        expect(v.error).toContain('total');
    });

    test('tolerates missing/odd input', () => {
        expect(validateAttachment(null, []).ok).toBe(false);
        expect(validateAttachment({ name: 'a.txt', size: 0, kind: 'text' }, null).ok).toBe(true);
    });
});

describe('formatBytes', () => {
    test('scales through B/KB/MB', () => {
        expect(formatBytes(512)).toBe('512 B');
        expect(formatBytes(2048)).toBe('2.0 KB');
        expect(formatBytes(3 * 1024 * 1024)).toBe('3.0 MB');
        expect(formatBytes(undefined)).toBe('0 B');
    });
});

describe('parseAttachment', () => {
    test('reads text files via file.text()', async () => {
        const file = new File(['hello attachment'], 'notes.txt', { type: 'text/plain' });
        const att = await parseAttachment(file);
        expect(att).toMatchObject({ name: 'notes.txt', kind: 'text', text: 'hello attachment' });
        expect(att.size).toBeGreaterThan(0);
    });

    test('reads images as base64 data URLs', async () => {
        const file = new File([new Uint8Array([137, 80, 78, 71])], 'tiny.png', { type: 'image/png' });
        const att = await parseAttachment(file);
        expect(att.kind).toBe('image');
        expect(att.dataUrl).toBe(`data:image/png;base64,${btoa(String.fromCharCode(137, 80, 78, 71))}`);
    });

    test('extracts docx text through mammoth (trimmed)', async () => {
        const file = new File(['zip-bytes'], 'contract.docx');
        const att = await parseAttachment(file);
        expect(att).toMatchObject({ name: 'contract.docx', kind: 'docx', text: 'extracted docx text' });
        expect(mammoth.extractRawText).toHaveBeenCalledTimes(1);
        const arg = mammoth.extractRawText.mock.calls[0][0];
        expect(arg.arrayBuffer).toBeInstanceOf(ArrayBuffer);
    });

    test('accepts and parses a 3.1 MiB docx', async () => {
        const file = new File([new Uint8Array(Math.ceil(3.1 * 1024 * 1024))], 'large.docx');
        expect(validateAttachment({ name: file.name, size: file.size, kind: detectAttachmentKind(file.name) }))
            .toEqual({ ok: true });
        await expect(parseAttachment(file)).resolves.toMatchObject({
            name: 'large.docx', kind: 'docx', size: file.size, text: 'extracted docx text',
        });
        expect(mammoth.extractRawText.mock.calls.at(-1)[0].arrayBuffer.byteLength).toBe(file.size);
    });

    describe('PDF reader lifecycle', () => {
        let readers;
        let pages;
        let doc;
        let loadingTask;

        beforeEach(() => {
            const pageChunks = [
                [
                    [{ type: 'beginMarkedContent', id: 'p1' }, { str: '  Hello', hasEOL: true }],
                    [],
                    [{ str: 'wor' }, { type: 'endMarkedContent' }, { str: 'ld' }],
                ],
                [[{ str: 'Second ' }], [{ str: 'page', hasEOL: true }, { str: '' }]],
            ];
            readers = pageChunks.map((chunks) => {
                const read = jest.fn();
                for (const items of chunks) read.mockResolvedValueOnce({ done: false, value: { items } });
                read.mockResolvedValue({ done: true });
                return { read, releaseLock: jest.fn() };
            });
            pages = readers.map((reader) => ({
                streamTextContent: jest.fn(() => ({ getReader: jest.fn(() => reader) })),
            }));
            doc = { numPages: pages.length, getPage: jest.fn(async (pageNum) => pages[pageNum - 1]) };
            loadingTask = { promise: Promise.resolve(doc), destroy: jest.fn().mockResolvedValue() };
            pdfjs.getDocument.mockReset().mockReturnValue(loadingTask);
        });

        test('uses reader-only streams, retaining chunk/EOL/page formatting and skipping nontext items', async () => {
            const att = await parseAttachment(new File(['pdf-bytes'], 'paper.pdf'));
            expect(att.kind).toBe('pdf');
            expect(att.text).toBe('Hello\nworld\n\nSecond page');
            expect(pdfjs.getDocument).toHaveBeenCalledTimes(1);
            expect(pdfjs.GlobalWorkerOptions.workerSrc).toContain('pdf.worker.min.mjs');
            expect(doc.getPage.mock.calls).toEqual([[1], [2]]);
            for (const [index, page] of pages.entries()) {
                const stream = page.streamTextContent.mock.results[0].value;
                expect(stream[Symbol.asyncIterator]).toBeUndefined();
                expect(stream.getReader).toHaveBeenCalledTimes(1);
                expect(readers[index].releaseLock).toHaveBeenCalledTimes(1);
            }
            expect(readers[0].read).toHaveBeenCalledTimes(4);
            expect(readers[1].read).toHaveBeenCalledTimes(3);
            expect(loadingTask.destroy).toHaveBeenCalledTimes(1);
        });

        test.each(['load', 'page', 'stream', 'reader', 'read'])('%s failure cleans up without masking the parse error', async (stage) => {
            const failure = new Error(`${stage} failed`);
            if (stage === 'load') loadingTask.promise = Promise.reject(failure);
            if (stage === 'page') doc.getPage.mockImplementation(async (pageNum) => {
                if (pageNum === 2) throw failure;
                return pages[0];
            });
            if (stage === 'stream') pages[0].streamTextContent.mockImplementation(() => { throw failure; });
            if (stage === 'reader') pages[0].streamTextContent.mockReturnValue({
                getReader: () => { throw failure; },
            });
            if (stage === 'read') readers[0].read.mockReset()
                .mockResolvedValueOnce({ done: false, value: { items: [{ str: 'partial' }] } })
                .mockRejectedValue(failure);
            for (const reader of readers) reader.releaseLock.mockImplementation(() => { throw new Error('release failed'); });
            loadingTask.destroy.mockRejectedValue(new Error('destroy failed'));

            await expect(parseAttachment(new File(['x'], 'broken.pdf')))
                .rejects.toThrow(`broken.pdf: ${stage} failed`);
            expect(readers[0].releaseLock).toHaveBeenCalledTimes(['page', 'read'].includes(stage) ? 1 : 0);
            expect(readers[1].releaseLock).not.toHaveBeenCalled();
            expect(loadingTask.destroy).toHaveBeenCalledTimes(1);
        });

        test.each(['sync', 'async'])('%s cleanup failures do not discard successfully extracted text', async (mode) => {
            readers[0].releaseLock.mockImplementation(() => { throw new Error('release failed'); });
            if (mode === 'sync') loadingTask.destroy.mockImplementation(() => { throw new Error('destroy failed'); });
            else loadingTask.destroy.mockRejectedValue(new Error('destroy failed'));
            await expect(parseAttachment(new File(['x'], 'paper.pdf')))
                .resolves.toMatchObject({ text: 'Hello\nworld\n\nSecond page' });
            for (const reader of readers) expect(reader.releaseLock).toHaveBeenCalledTimes(1);
            expect(loadingTask.destroy).toHaveBeenCalledTimes(1);
        });
    });

    test('wraps parser failures with the file name', async () => {
        mammoth.extractRawText.mockRejectedValueOnce(new Error('corrupt zip'));
        await expect(parseAttachment(new File(['x'], 'broken.docx')))
            .rejects.toThrow('broken.docx: corrupt zip');
    });

    test('rejects unsupported files without touching parsers', async () => {
        await expect(parseAttachment(new File(['x'], 'a.zip', { type: 'application/zip' })))
            .rejects.toThrow('unsupported file type');
    });

    test('read failures without a message degrade to a generic reason', async () => {
        const file = { name: 'bad.txt', type: 'text/plain', size: 1, text: async () => { throw new Error(); } };
        await expect(parseAttachment(file)).rejects.toThrow('bad.txt: could not be read');
    });
});

describe('buildAttachmentContext', () => {
    test('returns empty string for no attachments', () => {
        expect(buildAttachmentContext()).toBe('');
        expect(buildAttachmentContext([])).toBe('');
        expect(buildAttachmentContext([null])).toBe('');
    });

    test('labels text files with name separators', () => {
        const out = buildAttachmentContext([
            { name: 'a.txt', kind: 'text', text: 'alpha body' },
            { name: 'b.pdf', kind: 'pdf', text: 'beta body' },
        ]);
        expect(out).toContain('--- ATTACHED FILE: a.txt ---\nalpha body');
        expect(out).toContain('--- ATTACHED FILE: b.pdf ---\nbeta body');
        expect(out.startsWith('\n\n')).toBe(true);
    });

    test('lists images by name only (bytes travel as image parts)', () => {
        const out = buildAttachmentContext([{ name: 'pic.png', kind: 'image', dataUrl: 'data:image/png;base64,xx' }]);
        expect(out).toContain('ATTACHED IMAGE: pic.png');
        expect(out).not.toContain('data:image');
    });

    test('missing extracted text degrades to a placeholder', () => {
        const out = buildAttachmentContext([{ name: 'empty.pdf', kind: 'pdf' }]);
        expect(out).toContain('(no text extracted)');
    });

    test('caps total output and notes the omission', () => {
        const big = 'x'.repeat(ATTACHMENT_LIMITS.MAX_CONTEXT_CHARS);
        const out = buildAttachmentContext([
            { name: 'big.txt', kind: 'text', text: big },
            { name: 'next.txt', kind: 'text', text: 'second file' },
        ]);
        expect(out.length).toBeLessThanOrEqual(ATTACHMENT_LIMITS.MAX_CONTEXT_CHARS + 100);
        expect(out).not.toContain('second file');
        expect(out).toContain('omitted');
    });
});

describe('splitAttachments', () => {
    test('separates text sources from image payloads', () => {
        const { textAttachments, imageAttachments } = splitAttachments([
            { name: 'a.txt', kind: 'text', text: 'body' },
            { name: 'p.png', kind: 'image', dataUrl: 'data:image/png;base64,zz' },
            { name: 'broken.png', kind: 'image' },
        ]);
        expect(textAttachments).toHaveLength(1);
        expect(imageAttachments).toEqual([{ name: 'p.png', dataUrl: 'data:image/png;base64,zz' }]);
    });

    test('handles empty/invalid input', () => {
        expect(splitAttachments(null)).toEqual({ textAttachments: [], imageAttachments: [] });
    });
});

describe('attachmentMeta', () => {
    test('keeps name/kind/size only — no text or dataUrl', () => {
        const meta = attachmentMeta([
            { name: 'a.txt', kind: 'text', size: 5, text: 'secret body' },
            { name: 'p.png', kind: 'image', size: 9, dataUrl: 'data:image/png;base64,zz' },
        ]);
        expect(meta).toEqual([
            { name: 'a.txt', kind: 'text', size: 5 },
            { name: 'p.png', kind: 'image', size: 9 },
        ]);
        expect(JSON.stringify(meta)).not.toContain('secret body');
        expect(JSON.stringify(meta)).not.toContain('data:image');
    });

    test('normalizes odd entries', () => {
        expect(attachmentMeta(null)).toEqual([]);
        expect(attachmentMeta([null, { name: 'a.txt' }])).toEqual([{ name: 'a.txt', kind: 'text', size: 0 }]);
    });

    test('nameless entries are dropped, not renamed', () => {
        // parseAttachment always assigns a name, so a nameless entry is
        // corrupt input. Keeping it as "(unnamed)" put a phantom chip in the
        // message and in persisted history.
        expect(attachmentMeta([{ size: 5 }, { name: '', kind: 'text' }])).toEqual([]);
    });
});

describe('ATTACHMENT_KIND / limits sanity', () => {
    test('kinds are frozen constants', () => {
        expect(Object.isFrozen(ATTACHMENT_KIND)).toBe(true);
        expect(Object.isFrozen(ATTACHMENT_LIMITS)).toBe(true);
        expect(ATTACHMENT_LIMITS).toEqual({
            MAX_FILES: 5,
            MAX_TEXT_FILE_BYTES: 10 * 1024 * 1024,
            MAX_IMAGE_FILE_BYTES: 4.5 * 1024 * 1024,
            MAX_TOTAL_BYTES: 10 * 1024 * 1024,
            MAX_CONTEXT_CHARS: 200000,
        });
    });
});
