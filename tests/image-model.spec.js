/**
 * Specs for src/lib/image-model.js — the image tool-loop draft model.
 * Covers stable snapshot indexes, consumption rules (delete/replace),
 * validation on every record* API, and the proposal-card item shapes.
 */

const { createImageModel, IMAGE_TOOL_SPECS, IMAGE_POSITIONS } = require('../src/lib/image-model.js');

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
        expect(model.recordAltText(2, 't').error).toMatch(/live snapshot index/);
        // Other indexes keep working.
        expect(model.recordAltText(1, 't').ok).toBe(true);
    });

    test('out-of-range and non-integer indexes reject', () => {
        const model = createImageModel(SNAPSHOT);
        expect(model.recordDelete(0).error).toMatch(/live snapshot index/);
        expect(model.recordDelete(3).error).toMatch(/live snapshot index/);
        expect(model.recordResize(1.5, 200).error).toMatch(/live snapshot index/);
        expect(model.recordAltText('1', 't').error).toMatch(/live snapshot index/);
    });

    test('recordResize enforces the width range', () => {
        const model = createImageModel(SNAPSHOT);
        expect(model.recordResize(1, 4).error).toMatch(/between 8 and 450/);
        expect(model.recordResize(1, 999).error).toMatch(/between 8 and 450/);
        expect(model.recordResize(1, 320).ok).toBe(true);
        expect(model.ops[0]).toEqual({ type: 'resize', index: 1, widthPt: 320 });
    });

    test('recordAltText trims and caps the text', () => {
        const model = createImageModel(SNAPSHOT);
        model.recordAltText(1, '  ' + 'x'.repeat(300) + '  ');
        expect(model.ops[0].text).toHaveLength(200);
        expect(model.ops[0].text).toBe('x'.repeat(200));
    });

    test('recordReplace consumes the index and stages the svg', () => {
        const model = createImageModel(SNAPSHOT);
        expect(model.recordReplace({ index: 1, instruction: 'dusk', svg: SVG }).ok).toBe(true);
        expect(model.ops[0]).toEqual({ type: 'replace', index: 1, instruction: 'dusk', svg: SVG });
        expect(model.recordDelete(1).error).toMatch(/live snapshot index/);
    });

    test('describeOps renders one card item per op', () => {
        const model = createImageModel(SNAPSHOT);
        model.recordDelete(2);
        model.recordInsert({ position: 'start', instruction: 'a hero image', svg: SVG });
        model.recordResize(1, 200);
        model.recordAltText(1, 'chart');

        const items = model.describeOps();
        expect(items.map((i) => i.id)).toEqual([1, 2, 3, 4]);
        expect(items[0]).toMatchObject({ label: expect.stringContaining('Delete image 2'), before: 'existing picture', after: '' });
        expect(items[1]).toMatchObject({ label: 'Insert illustration at start', svg: SVG });
        expect(items[2].label).toContain('Resize image 1');
        expect(items[2]).toMatchObject({ before: '300pt wide', after: '200pt wide' });
        expect(items[3]).toMatchObject({ before: 'sunset', after: 'chart' });
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
            'delete_image', 'resize_image', 'set_alt_text',
        ]);
        expect(IMAGE_POSITIONS).toEqual(['start', 'end', 'cursor']);
        // read_image is host-executed (agent-actions attaches the picture to
        // the next observation) — it stages no draft op of its own.
        const read = IMAGE_TOOL_SPECS.find((t) => t.name === 'read_image');
        expect(read.description).toMatch(/image input/i);
    });
});
