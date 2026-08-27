/**
 * Specs for selection content reading with images (word-actions.js):
 *   - imageDataUrl: base64 MIME sniffing
 *   - readSelectionContent: text + inline pictures, cap + totalImages,
 *     error swallowing
 *   - watchSelection: debounced callback receives the content object
 *
 * The Word/Office globals are replaced with minimal proxies.
 */

const {
    imageDataUrl, readSelectionContent, watchSelection,
} = require('../src/taskpane/word-actions.js');

/** Builds an inline-picture proxy. */
function pic({ width = 300, height = 200, altTextDescription = '', base64 = 'iVBORw0KGgo=' } = {}) {
    return {
        width,
        height,
        altTextDescription,
        load: jest.fn(),
        getBase64ImageSrc: jest.fn(() => ({ value: base64 })),
    };
}

/** Installs a Word.run mock whose selection carries text + pictures. */
function setSelectionWorld({ text = '', pictures = [] } = {}) {
    global.Word = {
        run: jest.fn(async (cb) => cb({
            document: {
                getSelection: () => ({
                    text,
                    load: jest.fn(),
                    inlinePictures: { items: pictures, load: jest.fn() },
                }),
            },
            sync: jest.fn().mockResolvedValue(undefined),
        })),
    };
}

afterEach(() => {
    delete global.Word;
    delete global.Office;
});

describe('imageDataUrl', () => {
    test('sniffs the MIME type from the base64 magic prefix', () => {
        expect(imageDataUrl('iVBORw0KGgo=')).toBe('data:image/png;base64,iVBORw0KGgo=');
        expect(imageDataUrl('/9j/4AAQ')).toBe('data:image/jpeg;base64,/9j/4AAQ');
        expect(imageDataUrl('R0lGODlh')).toBe('data:image/gif;base64,R0lGODlh');
        expect(imageDataUrl('UklGRvIA')).toBe('data:image/webp;base64,UklGRvIA');
        // Unknown prefix falls back to PNG.
        expect(imageDataUrl('AAAA')).toBe('data:image/png;base64,AAAA');
    });

    test('passes existing data URLs through', () => {
        expect(imageDataUrl('data:image/jpeg;base64,/9j/x')).toBe('data:image/jpeg;base64,/9j/x');
    });
});

describe('readSelectionContent', () => {
    test('reads text plus picture metadata and data URLs', async () => {
        setSelectionWorld({
            text: 'hello',
            pictures: [pic({ width: 320, height: 240, altTextDescription: 'a chart', base64: '/9j/4AAQ' })],
        });

        const content = await readSelectionContent();

        expect(content.text).toBe('hello');
        expect(content.totalImages).toBe(1);
        expect(content.images).toHaveLength(1);
        expect(content.images[0]).toMatchObject({
            width: 320, height: 240, altText: 'a chart', base64: '/9j/4AAQ',
        });
        expect(content.images[0].dataUrl).toBe('data:image/jpeg;base64,/9j/4AAQ');
    });

    test('caps fully-read images at 6 while reporting the true total', async () => {
        const pictures = Array.from({ length: 9 }, (_, i) => pic({ base64: `iVBOR${i}` }));
        setSelectionWorld({ text: '', pictures });

        const content = await readSelectionContent();

        expect(content.text).toBe('');
        expect(content.totalImages).toBe(9);
        expect(content.images).toHaveLength(6);
    });

    test('empty selections resolve empty, errors swallow to empty', async () => {
        setSelectionWorld({});
        expect(await readSelectionContent()).toEqual({ text: '', images: [], totalImages: 0 });

        global.Word = { run: jest.fn(async () => { throw new Error('boom'); }) };
        expect(await readSelectionContent()).toEqual({ text: '', images: [], totalImages: 0 });
    });
});

describe('watchSelection', () => {
    function installOffice() {
        let handler = null;
        global.Office = {
            context: {
                document: {
                    addHandlerAsync: jest.fn((_type, h, _cb) => { handler = h; }),
                    removeHandlerAsync: jest.fn(),
                },
            },
            EventType: { DocumentSelectionChanged: 'selectionChanged' },
        };
        return () => handler;
    }

    test('debounced events deliver the full content object (image-only selections included)', async () => {
        const getHandler = installOffice();
        setSelectionWorld({
            text: '',
            pictures: [pic({ width: 100, height: 50, altTextDescription: 'logo' })],
        });

        jest.useFakeTimers();
        try {
            const seen = [];
            watchSelection((content) => seen.push(content));

            // Initial state is emitted (debounced).
            await jest.advanceTimersByTimeAsync(250);
            expect(seen).toHaveLength(1);
            expect(seen[0].text).toBe('');
            expect(seen[0].images).toHaveLength(1);
            expect(seen[0].images[0]).toMatchObject({ width: 100, height: 50, altText: 'logo' });

            // A later selection change fires once after the debounce.
            getHandler()();
            getHandler()();
            await jest.advanceTimersByTimeAsync(250);
            expect(seen).toHaveLength(2);
        } finally {
            jest.useRealTimers();
        }
    });

    test('no Office global resolves to a no-op unsubscribe', () => {
        expect(typeof watchSelection(jest.fn())).toBe('function');
    });
});
