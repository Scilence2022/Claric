/**
 * Specs for selection content reading (word-actions.js):
 *   - imageDataUrl: base64 MIME sniffing
 *   - readSelectionContent: text + inline pictures (cap + totalImages),
 *     plus table shape (multi-cell region flag + corner coords)
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

/** Null table anchor proxy (selection not inside a table/cell). */
function nullAnchor() {
    return { isNullObject: true, load: jest.fn(), rowIndex: 0, cellIndex: 0 };
}

/**
 * Installs a Word.run mock whose selection carries text + pictures +
 * table anchors. Defaults to "not in a table" (every anchor null) so
 * image/text-only selections work out of the box. Pass `tableAnchors`
 * to simulate in-table selections.
 */
function setSelectionWorld({
    text = '', pictures = [],
    parentTable = false, parentTableCell = false,
    startCell, endCell,
} = {}) {
    // Cells passed explicitly are normalized so the production code's
    // `.load('isNullObject,rowIndex,cellIndex')` calls succeed — test literals
    // typically include the sync'd fields but not the load jest.fn.
    const startCellProxy = startCell
        ? { load: jest.fn(), ...startCell }
        : nullAnchor();
    const endCellProxy = endCell
        ? { load: jest.fn(), ...endCell }
        : nullAnchor();
    global.Word = {
        run: jest.fn(async (cb) => cb({
            document: {
                getSelection: () => ({
                    text,
                    load: jest.fn(),
                    inlinePictures: { items: pictures, load: jest.fn() },
                    parentTableOrNullObject: parentTable
                        ? { isNullObject: false, load: jest.fn() }
                        : nullAnchor(),
                    parentTableCellOrNullObject: parentTableCell
                        ? { isNullObject: false, rowIndex: 0, cellIndex: 0, load: jest.fn() }
                        : nullAnchor(),
                    getRange: (loc) => ({
                        parentTableCellOrNullObject:
                            loc === 'Start' ? startCellProxy : endCellProxy,
                    }),
                }),
            },
            sync: jest.fn().mockResolvedValue(undefined),
        })),
        // Word.RangeLocation keys are the same string literals word-actions
        // passes (`Start` / `End`).
        RangeLocation: { start: 'Start', end: 'End' },
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

const EMPTY_CONTENT = {
    text: '', images: [], totalImages: 0,
    hasMultiCellTableRegion: false, tableRegion: null,
};

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
        expect(content.hasMultiCellTableRegion).toBe(false);
        expect(content.tableRegion).toBeNull();
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
        expect(await readSelectionContent()).toEqual(EMPTY_CONTENT);

        global.Word = { run: jest.fn(async () => { throw new Error('boom'); }) };
        expect(await readSelectionContent()).toEqual(EMPTY_CONTENT);
    });

    test('flags multi-cell region + captures corner coords', async () => {
        // Selection inside a table, NOT fully inside a single anchor cell, with
        // both endpoint cells resolving to real coordinates.
        setSelectionWorld({
            text: 'cell text',
            parentTable: true,
            parentTableCell: false,
            startCell: { isNullObject: false, rowIndex: 0, cellIndex: 0 },
            endCell: { isNullObject: false, rowIndex: 2, cellIndex: 1 },
        });

        const content = await readSelectionContent();
        expect(content.hasMultiCellTableRegion).toBe(true);
        expect(content.tableRegion).toEqual({
            startRow: 1, endRow: 3, startCol: 1, endCol: 2,
        });
    });

    test('whole-table boundary selections flag multi-cell but coords are null', async () => {
        // Boundary: start endpoint returns null (selection starts outside any
        // cell). Corner coords are unresolvable from the cheap read — the full
        // region read inside prepareTableToolEdit clamps to the table edge.
        setSelectionWorld({
            text: '',
            parentTable: true,
            parentTableCell: false,
            startCell: nullAnchor(),  // null = selection endpoint on boundary
            endCell: { isNullObject: false, rowIndex: 2, cellIndex: 1 },
        });

        const content = await readSelectionContent();
        expect(content.hasMultiCellTableRegion).toBe(true);
        expect(content.tableRegion).toBeNull();
    });

    test('intra-cell selection is NOT flagged multi-cell', async () => {
        // Selection anchor cell IS non-null (fully inside one cell). The table
        // exists but the selection is wholly within the anchor — text-pipeline.
        setSelectionWorld({
            text: 'inside one cell',
            parentTable: true,
            parentTableCell: true,
        });

        const content = await readSelectionContent();
        expect(content.hasMultiCellTableRegion).toBe(false);
        expect(content.tableRegion).toBeNull();
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
