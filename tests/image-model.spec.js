/**
 * Specs for src/lib/image-model.js — the image tool-loop draft model.
 * Covers stable snapshot indexes, consumption rules (delete/replace),
 * validation on every record* API, and the proposal-card item shapes.
 */

const { createImageModel, IMAGE_TOOL_SPECS, IMAGE_POSITIONS, normalizeImageLink } = require('../src/lib/image-model.js');

const SNAPSHOT = [
    { index: 1, width: 300, height: 200, altText: 'sunset' },
    { index: 2, width: 450, height: 300, altText: '' },
];

const SVG = '<svg width="1200" height="800" viewBox="0 0 1200 800"><rect width="1200" height="800"/></svg>';

describe('createImageModel', () => {
    test('list_images returns the snapshot and pending ops without svg payloads', () => {
        const model = createImageModel(SNAPSHOT);
        model.recordDelete(2);
        const { ok, result } = model.listImages();

        expect(ok).toBe(true);
        expect(result.count).toBe(2);
        expect(result.images[0]).toEqual({ index: 1, widthPt: 300, heightPt: 200, altText: 'sunset' });
        expect(result.pendingOps).toEqual([{ id: 1, type: 'delete', index: 2 }]);
    });

    test('recordInsert validates position, instruction, and svg', () => {
        const model = createImageModel(SNAPSHOT);
        expect(model.recordInsert({ position: 'middle', instruction: 'x', svg: SVG }).error).toMatch(/position/);
        expect(model.recordInsert({ position: 'end', instruction: '', svg: SVG }).error).toMatch(/design brief/);
        expect(model.recordInsert({ position: 'end', instruction: 'a sun', svg: 'not svg' }).error).toMatch(/no usable SVG/);
        expect(model.recordInsert({ position: 'cursor', instruction: 'a sun', svg: SVG }).ok).toBe(true);
        expect(model.ops[0].type).toBe('insert');
    });

    test('consumed indexes reject further ops (delete/replace consume)', () => {
        const model = createImageModel(SNAPSHOT);
        expect(model.recordDelete(2).ok).toBe(true);
        expect(model.recordDelete(2).error).toMatch(/live snapshot index/);
        expect(model.recordReplace({ index: 2, instruction: 'x', svg: SVG }).error).toMatch(/live snapshot index/);
        expect(model.recordAltText(2, { text: 't' }).error).toMatch(/live snapshot index/);
        // Other indexes keep working.
        expect(model.recordAltText(1, { text: 't' }).ok).toBe(true);
    });

    test('out-of-range and non-integer indexes reject', () => {
        const model = createImageModel(SNAPSHOT);
        expect(model.recordDelete(0).error).toMatch(/live snapshot index/);
        expect(model.recordDelete(3).error).toMatch(/live snapshot index/);
        expect(model.recordResize(1.5, { widthPt: 200 }).error).toMatch(/live snapshot index/);
        expect(model.recordAltText('1', 't').error).toMatch(/live snapshot index/);
    });

    test('recordResize accepts width, height, scale, lock-only — one size target at a time', () => {
        const model = createImageModel(SNAPSHOT);
        expect(model.recordResize(1, { widthPt: 4 }).error).toMatch(/between 8 and 450/);
        expect(model.recordResize(1, { widthPt: 999 }).error).toMatch(/between 8 and 450/);
        expect(model.recordResize(1, { widthPt: 320 }).ok).toBe(true);
        expect(model.ops[0]).toEqual({ type: 'resize', index: 1, widthPt: 320 });

        const m2 = createImageModel(SNAPSHOT);
        expect(m2.recordResize(2, { heightPt: 8 }).ok).toBe(true);
        expect(m2.recordResize(1, { heightPt: 700 }).error).toMatch(/between 8 and 600/);
        expect(m2.ops[0]).toEqual({ type: 'resize', index: 2, heightPt: 8 });

        const m3 = createImageModel(SNAPSHOT);
        expect(m3.recordResize(1, { scalePct: 150 }).ok).toBe(true);
        expect(m3.recordResize(1, { scalePct: 3 }).error).toMatch(/between 5 and 400/);
        expect(m3.recordResize(1, { scalePct: 500 }).error).toMatch(/between 5 and 400/);

        expect(model.recordResize(2, { widthPt: 100, heightPt: 100 }).error).toMatch(/exactly ONE/);

        const m4 = createImageModel(SNAPSHOT);
        expect(m4.recordResize(1, { lockAspectRatio: true }).ok).toBe(true);
        expect(m4.ops[0]).toEqual({ type: 'resize', index: 1, lockAspectRatio: true });

        const m5 = createImageModel(SNAPSHOT);
        expect(m5.recordResize(1, { widthPt: 200, lockAspectRatio: false }).ok).toBe(true);
        expect(m5.recordResize(1, { lockAspectRatio: 'yes' }).error).toMatch(/true or false/);
    });

    test('recordAlign validates the alignment value', () => {
        const model = createImageModel(SNAPSHOT);
        expect(model.recordAlign(1, 'justify').error).toMatch(/alignment.*left.*centered.*right/);
        expect(model.recordAlign(2, 'centered').ok).toBe(true);
        expect(model.ops[0]).toEqual({ type: 'align', index: 2, alignment: 'centered' });
        // Consumed indexes reject further ops.
        model.recordDelete(1);
        expect(model.recordAlign(1, 'right').error).toMatch(/live snapshot index/);
    });

    test('set_alt_text accepts description and/or title', () => {
        const model = createImageModel(SNAPSHOT);
        expect(model.recordAltText(1, { text: 'chart' }).ok).toBe(true);
        expect(model.ops[0]).toEqual({ type: 'altText', index: 1, text: 'chart' });

        const m2 = createImageModel(SNAPSHOT);
        expect(m2.recordAltText(1, { title: 'Revenue 2025' }).ok).toBe(true);
        expect(m2.ops[0]).toEqual({ type: 'altText', index: 1, title: 'Revenue 2025' });

        const m3 = createImageModel(SNAPSHOT);
        expect(m3.recordAltText(1, { text: 'desc', title: 'short' }).ok).toBe(true);
        expect(m3.ops[0]).toEqual({ type: 'altText', index: 1, text: 'desc', title: 'short' });

        expect(model.recordAltText(1, {}).error).toMatch(/at least one/);
        expect(model.recordAltText(1, '  ').error).toMatch(/at least one/);
        // Legacy 2-arg string call still accepted for description only.
        expect(model.recordAltText(1, 'plain text').ok).toBe(true);
    });

    test('recordLink validates scheme / internal location / clear', () => {
        const model = createImageModel(SNAPSHOT);
        expect(model.recordLink(1, 'https://example.com/report').ok).toBe(true);
        expect(model.recordLink(2, 'mailto:[email protected]').ok).toBe(true);
        expect(model.recordLink(1, '#section-2').ok).toBe(true);
        expect(model.recordLink(1, '').ok).toBe(true);
        expect(model.recordLink(1, null).ok).toBe(true);

        // Bad scheme, missing scheme on a bare domain.
        expect(model.recordLink(1, 'example.com').error).toMatch(/scheme/);
        expect(model.recordLink(1, 'javascript:alert(1)').error).toMatch(/scheme/);
        expect(model.recordLink(1, 'a'.repeat(2049)).error).toMatch(/too long/);
        expect(model.recordLink(1, 42).error).toMatch(/must be a string/);

        // Clear: url field is null (apply deletes the link).
        const clearModel = createImageModel(SNAPSHOT);
        clearModel.recordLink(1, '');
        expect(clearModel.ops.find((o) => o.type === 'link')).toEqual({ type: 'link', index: 1, url: null });
    });

    test('recordReplace consumes the index and stages the svg', () => {
        const model = createImageModel(SNAPSHOT);
        expect(model.recordReplace({ index: 1, instruction: 'dusk', svg: SVG }).ok).toBe(true);
        expect(model.ops[0]).toEqual({ type: 'replace', index: 1, instruction: 'dusk', svg: SVG });
        expect(model.recordDelete(1).error).toMatch(/live snapshot index/);
    });

    test('recordReplace carries beforeSrc through to the card item', () => {
        const model = createImageModel(SNAPSHOT);
        const beforeSrc = 'data:image/png;base64,iVBORw0KGgo=';
        expect(model.recordReplace({ index: 1, instruction: 'dusk', svg: SVG, beforeSrc }).ok).toBe(true);
        expect(model.ops[0].beforeSrc).toBe(beforeSrc);
        const item = model.describeOps()[0];
        expect(item.beforeSrc).toBe(beforeSrc);
        expect(item.svg).toBe(SVG);
        // Without beforeSrc the item simply omits the field.
        const plain = createImageModel(SNAPSHOT);
        plain.recordReplace({ index: 2, instruction: 'x', svg: SVG });
        expect('beforeSrc' in plain.describeOps()[0]).toBe(false);
    });

    test('list_images surfaces the stored-SVG-source capability flag', () => {
        const model = createImageModel([
            { index: 1, width: 300, height: 200, altText: 'sunset', hasSvgSource: true },
            { index: 2, width: 450, height: 300, altText: '' },
        ]);
        const { result } = model.listImages();
        expect(result.images[0].hasSvgSource).toBe(true);
        expect('hasSvgSource' in result.images[1]).toBe(false);
    });

    test('describeOps renders one card item per op', () => {
        const model = createImageModel(SNAPSHOT);
        model.recordDelete(2);
        model.recordInsert({ position: 'start', instruction: 'a hero image', svg: SVG });
        model.recordResize(1, { widthPt: 200 });
        model.recordAltText(1, { text: 'chart', title: 'Q3' });
        model.recordAlign(1, 'centered');
        // Index 2 was deleted above — link targets index 1 instead.
        model.recordLink(1, 'https://example.com');

        const items = model.describeOps();
        expect(items.map((i) => i.id)).toEqual([1, 2, 3, 4, 5, 6]);
        expect(items[0]).toMatchObject({ label: expect.stringContaining('Delete image 2'), before: 'existing picture', after: '' });
        expect(items[1]).toMatchObject({ label: 'Insert illustration at start', svg: SVG });
        expect(items[2].label).toContain('Resize image 1');
        expect(items[2]).toMatchObject({ before: '300x200pt', after: expect.stringContaining('200pt wide') });
        expect(items[3]).toMatchObject({
            before: expect.stringContaining('desc: "sunset"'),
            after: expect.stringContaining('desc: "chart"'),
        });
        expect(items[3].after).toContain('title: "Q3"');
        expect(items[4]).toMatchObject({ before: 'paragraph alignment', after: 'centered' });
        expect(items[5]).toMatchObject({ after: 'https://example.com' });

        // Clear-link card item.
        const m2 = createImageModel(SNAPSHOT);
        m2.recordLink(1, '');
        expect(m2.describeOps()[0]).toMatchObject({ label: expect.stringContaining('Clear link'), after: '(none)' });
    });

    test('empty snapshot and empty ops behave', () => {
        const model = createImageModel([]);
        expect(model.imageCount).toBe(0);
        expect(model.listImages().result.images).toEqual([]);
        expect(model.describeOps()).toEqual([]);
        expect(model.recordDelete(1).error).toMatch(/live snapshot index/);
    });
});

describe('IMAGE_TOOL_SPECS', () => {
    test('covers the full management surface incl. visual reading', () => {
        expect(IMAGE_TOOL_SPECS.map((t) => t.name)).toEqual([
            'list_images', 'read_image', 'design_illustration', 'replace_illustration',
            'edit_illustration_text',
            'delete_image', 'resize_image', 'align_image', 'set_alt_text', 'set_image_link',
        ]);
        expect(IMAGE_POSITIONS).toEqual(['start', 'end', 'cursor']);
        // read_image is host-executed (agent-actions attaches the picture to
        // the next observation) — it stages no draft op of its own.
        const read = IMAGE_TOOL_SPECS.find((t) => t.name === 'read_image');
        expect(read.description).toMatch(/image input/i);
        expect(read.description).toMatch(/documentContext/);
        expect(read.description).toMatch(/captionCandidates/);
        expect(read.description).toMatch(/legend/i);
        expect(read.description).toMatch(/untrusted data/i);
        expect(read.description).toMatch(/do not call an ordinary nearby paragraph a caption/i);
        // Style-y tools cover align/alt/link.
        const align = IMAGE_TOOL_SPECS.find((t) => t.name === 'align_image');
        expect(align.description).toMatch(/centered/);
        const link = IMAGE_TOOL_SPECS.find((t) => t.name === 'set_image_link');
        expect(link.description).toMatch(/scheme/);
        // The deterministic label editor points at its fallback.
        const editText = IMAGE_TOOL_SPECS.find((t) => t.name === 'edit_illustration_text');
        expect(editText.description).toMatch(/stored SVG source/);
        expect(editText.description).toMatch(/replace_illustration/);
    });
});

describe('normalizeImageLink', () => {
    test('accepts scheme-bearing URLs and #locations, rejects bare hosts', () => {
        expect(normalizeImageLink('https://example.com').url).toBe('https://example.com');
        expect(normalizeImageLink('mailto:[email protected]').url).toBe('mailto:[email protected]');
        expect(normalizeImageLink('#anchor').url).toBe('#anchor');
        expect(normalizeImageLink('').url).toBeNull();
        expect(normalizeImageLink(null).url).toBeNull();
        expect(normalizeImageLink('example.com').error).toMatch(/scheme/);
        expect(normalizeImageLink(42).error).toMatch(/string/);
    });
});
