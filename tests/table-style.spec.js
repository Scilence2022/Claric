/**
 * Specs for src/lib/table-style.js — the pure style-op vocabulary used by the
 * table tool loop: color/border/alignment/font normalization, target-region
 * clipping, and card labels.
 */

const {
    TABLE_STYLE_LIMITS,
    BUILT_IN_TABLE_STYLES,
    BORDER_TYPES,
    normalizeColor,
    normalizeBorderSpec,
    expandBorderSet,
    normalizeAlignment,
    normalizeFontPayload,
    normalizeStyleTarget,
    isFullWidthRegion,
    describeStyleOp,
} = require('../src/lib/table-style.js');

const BOUNDS = { startRow: 2, endRow: 4, startCol: 1, endCol: 3 };

describe('normalizeColor', () => {
    test('accepts #RRGGBB (any case), common names, and auto', () => {
        expect(normalizeColor('#deebf7')).toBe('#DEEBF7');
        expect(normalizeColor(' #DEEBF7 ')).toBe('#DEEBF7');
        expect(normalizeColor('gray')).toBe('#808080');
        expect(normalizeColor('grey')).toBe('#808080');
        expect(normalizeColor('auto')).toBe('auto');
    });

    test('rejects unusable values', () => {
        expect(normalizeColor('#12')).toBeNull();
        expect(normalizeColor('notacolor')).toBeNull();
        expect(normalizeColor(42)).toBeNull();
        expect(normalizeColor(null)).toBeNull();
    });
});

describe('normalizeBorderSpec', () => {
    test('string shorthand becomes {type}', () => {
        expect(normalizeBorderSpec('none')).toEqual({ spec: { type: 'none' } });
        expect(normalizeBorderSpec('Single')).toEqual({ spec: { type: 'single' } });
        expect(normalizeBorderSpec('dot-dashed')).toEqual({ spec: { type: 'dotDashed' } });
    });

    test('object spec carries color and width; "none" drops them', () => {
        expect(normalizeBorderSpec({ type: 'single', color: 'red', width: 1.5 }))
            .toEqual({ spec: { type: 'single', color: '#FF0000', width: 1.5 } });
        expect(normalizeBorderSpec({ type: 'none', color: 'red', width: 3 }))
            .toEqual({ spec: { type: 'none' } });
    });

    test('rejects bad type, color, and width', () => {
        expect(normalizeBorderSpec('sparkly').error).toMatch(/unknown border type/);
        expect(normalizeBorderSpec({ type: 'single', color: 'zebra' }).error).toMatch(/invalid border color/);
        expect(normalizeBorderSpec({ type: 'single', width: 99 }).error).toMatch(/width must be/);
        expect(normalizeBorderSpec(7).error).toMatch(/must be a type string/);
    });

    test('BORDER_TYPES uses the Word enum names (Med, not Medium)', () => {
        expect(BORDER_TYPES).toContain('thinThickMed');
        expect(BORDER_TYPES).not.toContain('thinThickMedium');
    });
});

describe('expandBorderSet', () => {
    test('expands all/inside/outside aliases into concrete locations', () => {
        const { borders } = expandBorderSet({ all: 'none' });
        expect(Object.keys(borders).sort()).toEqual(['bottom', 'insideH', 'insideV', 'left', 'right', 'top']);
        expect(borders.top).toEqual({ type: 'none' });
    });

    test('concrete keys override aliases; insideHorizontal alias works', () => {
        const { borders } = expandBorderSet({
            outside: { type: 'single', width: 1.5 },
            top: 'double',
            insideHorizontal: { type: 'single', color: 'auto' },
        });
        expect(borders.top).toEqual({ type: 'double' });
        expect(borders.bottom).toEqual({ type: 'single', width: 1.5 });
        expect(borders.left).toEqual({ type: 'single', width: 1.5 });
        expect(borders.insideH).toEqual({ type: 'single', color: 'auto' });
        // "outside" never touches inside borders.
        expect(borders.insideV).toBeUndefined();
    });

    test('errors on non-object input, bad spec, and empty set', () => {
        expect(expandBorderSet('single').error).toMatch(/must be an object/);
        expect(expandBorderSet({ top: 'zebra' }).error).toMatch(/top: unknown border type/);
        expect(expandBorderSet({}).error).toMatch(/no border locations/);
    });
});

describe('normalizeAlignment', () => {
    test('canonicalizes with aliases', () => {
        expect(normalizeAlignment('center', ['left', 'centered', 'right'])).toBe('centered');
        expect(normalizeAlignment('Centered', ['left', 'centered', 'right'])).toBe('centered');
        expect(normalizeAlignment('middle', ['top', 'center', 'bottom'])).toBe('center');
        expect(normalizeAlignment('justify', ['left', 'centered', 'right', 'justified'])).toBe('justified');
        expect(normalizeAlignment('diagonal', ['left', 'centered'])).toBeNull();
        expect(normalizeAlignment(5, ['left'])).toBeNull();
    });
});

describe('normalizeFontPayload', () => {
    test('keeps allowlisted keys and normalizes values', () => {
        expect(normalizeFontPayload({ bold: true, size: '11.3', name: ' SimSun ', color: '#000' }))
            .toEqual({ error: expect.stringContaining('font color') });
        expect(normalizeFontPayload({ bold: true, size: 11.3, name: ' SimSun ', color: 'black', underline: 'double' }))
            .toEqual({ font: { bold: true, size: 11.5, name: 'SimSun', color: '#000000', underline: 'double' } });
        expect(normalizeFontPayload({ italic: false })).toEqual({ font: { italic: false } });
    });

    test('ignores unknown keys but rejects empty/invalid payloads', () => {
        expect(normalizeFontPayload({ nonsense: 1 }).error).toMatch(/no supported keys/);
        expect(normalizeFontPayload({ bold: 'yes' }).error).toMatch(/bold.*true or false/);
        expect(normalizeFontPayload({ size: 400 }).error).toMatch(/font.size/);
        expect(normalizeFontPayload('bold').error).toMatch(/must be an object/);
    });
});

describe('normalizeStyleTarget', () => {
    test('cell / row band / column band / block / whole table', () => {
        expect(normalizeStyleTarget({ row: 3, col: 2 }, BOUNDS).region)
            .toEqual({ startRow: 3, endRow: 3, startCol: 2, endCol: 2 });
        expect(normalizeStyleTarget({ row: 3 }, BOUNDS).region)
            .toEqual({ startRow: 3, endRow: 3, startCol: 1, endCol: 3 });
        expect(normalizeStyleTarget({ col: 2 }, BOUNDS).region)
            .toEqual({ startRow: 2, endRow: 4, startCol: 2, endCol: 2 });
        expect(normalizeStyleTarget({ row: 2, col: 1, rows: 2, cols: 3 }, BOUNDS).region)
            .toEqual({ startRow: 2, endRow: 3, startCol: 1, endCol: 3 });
        expect(normalizeStyleTarget(null, BOUNDS).region).toBeNull();
    });

    test('clips to bounds and validates spans', () => {
        expect(normalizeStyleTarget({ row: 5 }, BOUNDS).error).toMatch(/outside the covered region/);
        expect(normalizeStyleTarget({ row: 3, col: 0 }, BOUNDS).error).toMatch(/positive integer/);
        expect(normalizeStyleTarget({ rows: 2 }, BOUNDS).error).toMatch(/"rows" requires "row"/);
        expect(normalizeStyleTarget({ row: 4, rows: 2 }, BOUNDS).error).toMatch(/outside the covered region/);
    });
});

describe('isFullWidthRegion', () => {
    test('true only when the region spans every column', () => {
        expect(isFullWidthRegion({ startRow: 1, endRow: 2, startCol: 1, endCol: 3 }, 3)).toBe(true);
        expect(isFullWidthRegion({ startRow: 1, endRow: 2, startCol: 2, endCol: 3 }, 3)).toBe(false);
        expect(isFullWidthRegion(null, 3)).toBe(false);
    });
});

describe('BUILT_IN_TABLE_STYLES', () => {
    test('covers the Word table style families, uniquely', () => {
        expect(BUILT_IN_TABLE_STYLES).toContain('TableGrid');
        expect(BUILT_IN_TABLE_STYLES).toContain('TableGridLight');
        expect(BUILT_IN_TABLE_STYLES).toContain('PlainTable3');
        expect(BUILT_IN_TABLE_STYLES).toContain('GridTable4_Accent1');
        expect(BUILT_IN_TABLE_STYLES).toContain('GridTable5Dark_Accent6');
        expect(BUILT_IN_TABLE_STYLES).toContain('ListTable7Colorful');
        expect(new Set(BUILT_IN_TABLE_STYLES).size).toBe(BUILT_IN_TABLE_STYLES.length);
    });
});

describe('describeStyleOp', () => {
    test('renders a readable label per op type', () => {
        expect(describeStyleOp({ type: 'tableStyle', style: 'GridTable4_Accent1', bandedRows: true }))
            .toBe('Table look: style → GridTable4_Accent1, banded rows on');
        expect(describeStyleOp({ type: 'borders', borders: { top: { type: 'single', width: 1.5 }, insideH: { type: 'none' } } }))
            .toBe('Borders: top single 1.5pt, insideH none');
        expect(describeStyleOp({ type: 'cellFormat', region: { startRow: 1, endRow: 1, startCol: 1, endCol: 2 }, shadingColor: '#DEEBF7', horizontalAlignment: 'centered' }))
            .toBe('Format R1C1–R1C2: shading #DEEBF7, align centered');
        expect(describeStyleOp({ type: 'cellFormat', region: null, verticalAlignment: 'center' }))
            .toBe('Format whole table: valign center');
        expect(describeStyleOp({ type: 'font', region: { startRow: 1, endRow: 1, startCol: 1, endCol: 1 }, font: { bold: true, size: 11 } }))
            .toBe('Font R1C1: bold, 11pt');
        expect(describeStyleOp({ type: 'headerRow', rows: 1, font: { bold: true }, shadingColor: '#DEEBF7' }))
            .toBe('Header: 1 header row(s), bold, shading #DEEBF7');
        expect(describeStyleOp({ type: 'layout', alignment: 'centered', autoFitWindow: true }))
            .toBe('Layout: table centered, autofit to window');
        expect(describeStyleOp({ type: 'columnWidths', widthsPt: [120, 80] }))
            .toBe('Column widths: 120pt, 80pt');
        expect(describeStyleOp({ type: 'mystery' })).toBe('Table style change');
    });
});

describe('TABLE_STYLE_LIMITS', () => {
    test('is frozen', () => {
        expect(Object.isFrozen(TABLE_STYLE_LIMITS)).toBe(true);
    });
});
