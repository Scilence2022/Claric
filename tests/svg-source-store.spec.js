/**
 * Specs for src/taskpane/svg-source-store.js — the custom-XML-part store
 * that keeps an illustration's SVG source inside the document. The Office
 * shared API is mocked; DOMParser comes from jsdom (node env lacks it).
 */

const { JSDOM } = require('jsdom');
const {
    SVG_SOURCE_TITLE_PREFIX, isSvgSourceStoreAvailable,
    saveSvgSource, loadSvgSource, deleteSvgSource,
    svgSourceIdFromPicture, attachSvgSource,
} = require('../src/taskpane/svg-source-store.js');

if (typeof globalThis.DOMParser === 'undefined') {
    const dom = new JSDOM('');
    globalThis.DOMParser = dom.window.DOMParser;
}

const SVG = '<svg width="1200" height="800" viewBox="0 0 1200 800"><text x="10" y="10">Dispatch</text></svg>';

/**
 * Builds a mock Office global whose customXmlParts behave like the shared
 * API: addAsync mints an id, getByIdAsync returns a part with
 * getXmlAsync/deleteAsync backed by an in-memory map.
 */
function mockOffice() {
    const parts = new Map();
    let seq = 0;
    global.Office = {
        AsyncResultStatus: { Succeeded: 'succeeded', Failed: 'failed' },
        context: {
            document: {
                customXmlParts: {
                    addAsync: jest.fn((xml, cb) => {
                        seq += 1;
                        const id = `part-${seq}`;
                        parts.set(id, xml);
                        cb({ status: 'succeeded', value: { id } });
                    }),
                    getByIdAsync: jest.fn((id, cb) => {
                        if (!parts.has(id)) { cb({ status: 'failed', error: { message: 'not found' } }); return; }
                        cb({
                            status: 'succeeded',
                            value: {
                                id,
                                getXmlAsync: jest.fn((cb2) => cb2({ status: 'succeeded', value: parts.get(id) })),
                                deleteAsync: jest.fn((cb2) => { parts.delete(id); cb2({ status: 'succeeded' }); }),
                            },
                        });
                    }),
                    _parts: parts,
                },
            },
        },
    };
    return global.Office.context.document.customXmlParts;
}

afterEach(() => {
    delete global.Office;
});

describe('isSvgSourceStoreAvailable', () => {
    test('false without the Office global', () => {
        expect(isSvgSourceStoreAvailable()).toBe(false);
    });

    test('false when document.customXmlParts is missing', () => {
        global.Office = { context: { document: {} } };
        expect(isSvgSourceStoreAvailable()).toBe(false);
    });

    test('true with customXmlParts present', () => {
        mockOffice();
        expect(isSvgSourceStoreAvailable()).toBe(true);
    });
});

describe('save/load round-trip', () => {
    test('saves the SVG and loads it back verbatim', async () => {
        mockOffice();
        const id = await saveSvgSource(SVG);
        expect(typeof id).toBe('string');
        expect(await loadSvgSource(id)).toBe(SVG);
    });

    test('survives labels containing XML-special characters', async () => {
        mockOffice();
        const tricky = '<svg width="10" height="10"><text>A &amp; B &lt;tag&gt; "quoted"</text></svg>';
        const id = await saveSvgSource(tricky);
        expect(await loadSvgSource(id)).toBe(tricky);
    });

    test('save rejects non-SVG input and unavailable hosts', async () => {
        expect(await saveSvgSource(SVG)).toBeNull(); // no Office
        mockOffice();
        expect(await saveSvgSource('not svg')).toBeNull();
        expect(await saveSvgSource('')).toBeNull();
    });

    test('load returns null for unknown ids and non-SVG parts', async () => {
        const store = mockOffice();
        expect(await loadSvgSource('missing')).toBeNull();
        store._parts.set('junk', '<unrelated xmlns="urn:x"><data>hello</data></unrelated>');
        expect(await loadSvgSource('junk')).toBeNull();
        store._parts.set('badxml', '<claricIllustration xmlns="urn:claric:illustration-source"><broken');
        expect(await loadSvgSource('badxml')).toBeNull();
    });

    test('load/save resolve null when the host call fails', async () => {
        const store = mockOffice();
        store.addAsync.mockImplementation((xml, cb) => cb({ status: 'failed', error: { message: 'denied' } }));
        expect(await saveSvgSource(SVG)).toBeNull();
    });
});

describe('deleteSvgSource', () => {
    test('removes the part and swallows failures', async () => {
        const store = mockOffice();
        const id = await saveSvgSource(SVG);
        await deleteSvgSource(id);
        expect(store._parts.has(id)).toBe(false);
        // Unknown id and unavailable host: no throw.
        await deleteSvgSource('missing');
        delete global.Office;
        await deleteSvgSource(id);
    });
});

describe('svgSourceIdFromPicture', () => {
    test('parses the marker prefix off the alt-text title', () => {
        expect(svgSourceIdFromPicture({ altTextTitle: SVG_SOURCE_TITLE_PREFIX + 'part-7' }))
            .toBe('part-7');
    });

    test('returns null for ordinary titles and missing pictures', () => {
        expect(svgSourceIdFromPicture({ altTextTitle: 'Architecture diagram' })).toBeNull();
        expect(svgSourceIdFromPicture({ altTextTitle: '' })).toBeNull();
        expect(svgSourceIdFromPicture({})).toBeNull();
        expect(svgSourceIdFromPicture(null)).toBeNull();
    });
});

describe('attachSvgSource', () => {
    test('stores the SVG and queues the title link on the picture', async () => {
        mockOffice();
        const picture = { altTextTitle: '' };
        const ok = await attachSvgSource(picture, SVG);
        expect(ok).toBe(true);
        expect(picture.altTextTitle.startsWith(SVG_SOURCE_TITLE_PREFIX)).toBe(true);
        const id = svgSourceIdFromPicture(picture);
        expect(await loadSvgSource(id)).toBe(SVG);
    });

    test('returns false without persisting when the store is unavailable', async () => {
        const picture = { altTextTitle: '' };
        expect(await attachSvgSource(picture, SVG)).toBe(false);
        expect(picture.altTextTitle).toBe('');
    });

    test('returns false when the save fails, leaving the title untouched', async () => {
        const store = mockOffice();
        store.addAsync.mockImplementation((xml, cb) => cb({ status: 'failed' }));
        const picture = { altTextTitle: '' };
        expect(await attachSvgSource(picture, SVG)).toBe(false);
        expect(picture.altTextTitle).toBe('');
    });
});
