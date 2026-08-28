/**
 * Table Style Ops
 *
 * Vocabulary, validation, and description for the table styling half of the
 * table tool loop (lib/table-model.js). Style ops ride along the coordinate
 * patch: cell/row-scoped ops are coordinate-bound (applied before row
 * structure changes, while original coordinates are still valid); table-level
 * ops (borders, table style, layout, header row) run after the structure has
 * settled. Everything here mirrors what WordApi 1.3 exposes on Table /
 * TableRow / TableCell — border type/color/width via Table.getBorder(),
 * shadingColor, per-scope fonts, alignments, headerRowCount, banding flags,
 * cell padding, column widths.
 *
 * Pure module — no Office.js, no network. Hermetic-testable.
 *
 * @module table-style
 */

/** Hard limits for style ops accepted from the model. Frozen. */
export const TABLE_STYLE_LIMITS = Object.freeze({
    /** Max style ops one loop may stage. */
    MAX_STYLE_OPS: 64,
    /** Border width bounds in points (Word clamps internally too). */
    MIN_BORDER_WIDTH_PT: 0.25,
    MAX_BORDER_WIDTH_PT: 12,
    /** Font size bounds in points. */
    MIN_FONT_SIZE_PT: 1,
    MAX_FONT_SIZE_PT: 96,
    /** Font family name length cap. */
    MAX_FONT_NAME_CHARS: 80,
    /** Per-column width bounds in points. */
    MIN_COLUMN_WIDTH_PT: 5,
    MAX_COLUMN_WIDTH_PT: 1584,
    /** Whole-table width bounds in points (1584pt ≈ 22in, Word's max). */
    MAX_TABLE_WIDTH_PT: 1584,
    /** Cell padding bounds in points. */
    MIN_CELL_PADDING_PT: 0,
    MAX_CELL_PADDING_PT: 100,
});

/**
 * Canonical Word border types (Word.BorderType values, camelCase). "none"
 * removes a border. Validation is case/separator-insensitive; apply maps the
 * canonical name back onto the Word enum.
 */
export const BORDER_TYPES = Object.freeze([
    'none', 'single', 'dotted', 'dashedSmall', 'dashed', 'dotDashed', 'dot2Dashed',
    'double', 'wave', 'doubleWave', 'dashDotStroked', 'triple', 'thinThickSmall',
    'thickThinSmall', 'thinThickThinSmall', 'thinThickMed', 'thickThinMed',
    'thinThickThinMed', 'thinThickLarge', 'thickThinLarge', 'thinThickThinLarge',
    'threeDEmboss', 'threeDEngrave',
].map(Object.freeze));

/** Accepted border-set keys (input side). Location aliases map onto these. */
export const BORDER_SET_KEYS = Object.freeze([
    'top', 'bottom', 'left', 'right', 'insideH', 'insideV',
]);

/** Location aliases expanded into concrete border-set keys. */
const BORDER_ALIASES = Object.freeze({
    all: ['top', 'bottom', 'left', 'right', 'insideH', 'insideV'],
    outside: ['top', 'bottom', 'left', 'right'],
    inside: ['insideH', 'insideV'],
    insideHorizontal: ['insideH'],
    insideVertical: ['insideV'],
});

/** Canonical horizontal alignment values (Word.Alignment subset). */
export const HORIZONTAL_ALIGNMENTS = Object.freeze(['left', 'centered', 'right', 'justified']);

/** Canonical vertical alignment values (Word.VerticalAlignment subset). */
export const VERTICAL_ALIGNMENTS = Object.freeze(['top', 'center', 'bottom']);

/** Font payload keys the model may set (Word.Font subset). */
export const FONT_KEYS = Object.freeze(['bold', 'italic', 'underline', 'size', 'name', 'color']);

/** Accepted underline values (Word.UnderlineType subset). */
export const UNDERLINE_TYPES = Object.freeze(['none', 'single', 'double', 'dotted', 'dashed', 'dotDashed', 'words']);

/** Common color names mapped to hex; "auto" keeps the theme/default color. */
const NAMED_COLORS = Object.freeze({
    auto: 'auto',
    black: '#000000', white: '#FFFFFF',
    red: '#FF0000', green: '#008000', blue: '#0000FF',
    yellow: '#FFFF00', cyan: '#00FFFF', magenta: '#FF00FF',
    gray: '#808080', grey: '#808080', darkgray: '#A9A9A9', lightgray: '#D3D3D3',
    lightblue: '#ADD8E6', lightgreen: '#90EE90', lightyellow: '#FFFFE0',
    darkred: '#8B0000', darkblue: '#00008B', darkgreen: '#006400',
    orange: '#FFA500', purple: '#800080', brown: '#A52A2A', pink: '#FFC0CB',
});

/**
 * Canonical built-in table style names (Word.BuiltInStyleName values).
 * Generated from the family × accent pattern; the apply side maps them onto
 * the Word enum (portable across locales, unlike Table.style strings).
 */
export const BUILT_IN_TABLE_STYLES = Object.freeze((() => {
    const names = ['TableGrid', 'TableGridLight'];
    for (let i = 1; i <= 5; i++) names.push(`PlainTable${i}`);
    const gridFamilies = [
        'GridTable1Light', 'GridTable2', 'GridTable3', 'GridTable4', 'GridTable5Dark',
        'GridTable6Colorful', 'GridTable7Colorful',
    ];
    const listFamilies = [
        'ListTable1Light', 'ListTable2', 'ListTable3', 'ListTable4', 'ListTable5Dark',
        'ListTable6Colorful', 'ListTable7Colorful',
    ];
    for (const family of [...gridFamilies, ...listFamilies]) {
        names.push(family);
        for (let a = 1; a <= 6; a++) names.push(`${family}_Accent${a}`);
    }
    return names.sort();
})());

/**
 * Normalizes a color value: "#RRGGBB", a common color name, or "auto".
 * Returns null when the value is unusable.
 *
 * @param {*} value
 * @returns {string|null}
 */
export function normalizeColor(value) {
    if (typeof value !== 'string') return null;
    const key = value.trim().toLowerCase().replace(/\s+/g, '');
    if (NAMED_COLORS[key]) return NAMED_COLORS[key];
    if (/^#[0-9a-f]{6}$/i.test(key)) return key.toUpperCase();
    return null;
}

/**
 * Normalizes one border spec. Input is either a bare type string ("none") or
 * an object {type, color?, width?}. Returns {error} or the canonical spec
 * {type, color?, width?} — color/width are dropped for type "none".
 *
 * @param {string|object} raw
 * @returns {{error: string, spec?: undefined}|{error?: undefined,
 *   spec: {type: string, color?: string, width?: number}}}
 */
export function normalizeBorderSpec(raw) {
    const obj = typeof raw === 'string' ? { type: raw } : (raw && typeof raw === 'object' ? raw : null);
    if (!obj) return { error: 'border spec must be a type string or {type, color, width} object' };
    if (typeof obj.type !== 'string') return { error: 'border "type" is required (e.g. "single", "none")' };

    const type = _canonicalize(obj.type, BORDER_TYPES);
    if (!type) return { error: `unknown border type "${obj.type}"` };

    if (type === 'none') return { spec: { type } };

    const spec = { type };
    if (obj.color !== undefined && obj.color !== null) {
        const color = normalizeColor(obj.color);
        if (!color) return { error: `invalid border color "${obj.color}" (use "#RRGGBB", "auto", or a common color name)` };
        spec.color = color;
    }
    if (obj.width !== undefined && obj.width !== null) {
        const width = Number(obj.width);
        if (!Number.isFinite(width) || width < TABLE_STYLE_LIMITS.MIN_BORDER_WIDTH_PT
            || width > TABLE_STYLE_LIMITS.MAX_BORDER_WIDTH_PT) {
            return { error: `border width must be ${TABLE_STYLE_LIMITS.MIN_BORDER_WIDTH_PT}–${TABLE_STYLE_LIMITS.MAX_BORDER_WIDTH_PT}pt` };
        }
        spec.width = Math.round(width * 100) / 100;
    }
    return { spec };
}

/**
 * Expands a border-set input into concrete per-location specs. Keys: top,
 * bottom, left, right, insideH (insideHorizontal), insideV (insideVertical),
 * plus the aliases all / outside / inside — aliases apply first, concrete
 * keys win on conflict. Each value is a border spec (string or object).
 *
 * @param {object} raw - {top: ..., inside: ..., all: ...}
 * @returns {{error: string, borders?: undefined}|{error?: undefined,
 *   borders: Object<string, {type: string, color?: string, width?: number}>}}
 */
export function expandBorderSet(raw) {
    if (!raw || typeof raw !== 'object' || Array.isArray(raw)) {
        return { error: '"borders" must be an object keyed by border location' };
    }
    /** @type {Object<string, {type: string, color?: string, width?: number}>} */
    const borders = {};
    const apply = (rawSpec, locations) => {
        for (const location of locations) {
            const result = normalizeBorderSpec(rawSpec);
            if (result.error) return { error: `${location}: ${result.error}` };
            borders[location] = result.spec;
        }
        return null;
    };
    for (const [alias, locations] of Object.entries(BORDER_ALIASES)) {
        if (raw[alias] !== undefined) {
            const err = apply(raw[alias], locations);
            if (err) return err;
        }
    }
    const keyMap = { top: 'top', bottom: 'bottom', left: 'left', right: 'right', insideH: 'insideH', insideHorizontal: 'insideH', insideV: 'insideV', insideVertical: 'insideV' };
    for (const [key, location] of Object.entries(keyMap)) {
        if (raw[key] !== undefined) {
            const err = apply(raw[key], [location]);
            if (err) return err;
        }
    }
    if (Object.keys(borders).length === 0) {
        return { error: 'no border locations given (top/bottom/left/right/insideH/insideV/inside/outside/all)' };
    }
    return { borders };
}

/**
 * Normalizes an alignment value against a canonical list, tolerating common
 * aliases (center/centre/middle resolve to whichever of centered/center the
 * list uses; justify → justified).
 *
 * @param {*} value
 * @param {readonly string[]} canonical
 * @returns {string|null}
 */
export function normalizeAlignment(value, canonical) {
    if (typeof value !== 'string') return null;
    const key = value.trim().toLowerCase().replace(/\s+/g, '');
    if (canonical.includes(key)) return key;
    if (key === 'center' || key === 'centre' || key === 'middle') {
        if (canonical.includes('centered')) return 'centered';
        if (canonical.includes('center')) return 'center';
        return null;
    }
    if (key === 'justify') return canonical.includes('justified') ? 'justified' : null;
    return null;
}

/**
 * Normalizes a font payload to the allowlisted keys. Returns {error} or
 * {font} with only the valid, changed keys kept.
 *
 * @param {object} raw - e.g. {bold: true, size: 12, name: "SimSun", color: "red"}
 * @returns {{error: string, font?: undefined}|{error?: undefined, font: Object<string, *>}}
 */
export function normalizeFontPayload(raw) {
    if (!raw || typeof raw !== 'object' || Array.isArray(raw)) {
        return { error: '"font" must be an object (bold, italic, underline, size, name, color)' };
    }
    const font = {};
    for (const [key, value] of Object.entries(raw)) {
        switch (key) {
            case 'bold':
            case 'italic':
                if (typeof value !== 'boolean') return { error: `font.${key} must be true or false` };
                font[key] = value;
                break;
            case 'underline': {
                const type = _canonicalize(value, UNDERLINE_TYPES);
                if (!type) return { error: `unknown underline "${value}"` };
                font.underline = type;
                break;
            }
            case 'size': {
                const size = Number(value);
                if (!Number.isFinite(size) || size < TABLE_STYLE_LIMITS.MIN_FONT_SIZE_PT
                    || size > TABLE_STYLE_LIMITS.MAX_FONT_SIZE_PT) {
                    return { error: `font.size must be ${TABLE_STYLE_LIMITS.MIN_FONT_SIZE_PT}–${TABLE_STYLE_LIMITS.MAX_FONT_SIZE_PT} (points)` };
                }
                font.size = Math.round(size * 2) / 2;
                break;
            }
            case 'name': {
                if (typeof value !== 'string' || !value.trim()
                    || value.length > TABLE_STYLE_LIMITS.MAX_FONT_NAME_CHARS) {
                    return { error: 'font.name must be a font family name' };
                }
                font.name = value.trim();
                break;
            }
            case 'color': {
                const color = normalizeColor(value);
                if (!color) return { error: `invalid font color "${value}"` };
                font.color = color;
                break;
            }
            default:
                break; // unknown keys are ignored, not fatal
        }
    }
    if (Object.keys(font).length === 0) {
        return { error: 'font payload has no supported keys (bold, italic, underline, size, name, color)' };
    }
    return { font };
}

/**
 * Normalizes a cell-target descriptor into a rectangular region clipped to
 * the covered bounds. {row, col} select a cell (spans default 1×1), {row}
 * alone a full row band, {col} alone a full column band, and neither the
 * whole covered region (null means whole table at the apply side).
 *
 * @param {object|null} raw - {row?, col?, rows?, cols?}
 * @param {{startRow: number, endRow: number, startCol: number, endCol: number}} bounds
 * @returns {{error: string, region?: undefined}|{error?: undefined,
 *   region: {startRow: number, endRow: number, startCol: number, endCol: number}|null}}
 */
export function normalizeStyleTarget(raw, bounds) {
    if (raw === undefined || raw === null) return { region: null };
    if (typeof raw !== 'object' || Array.isArray(raw)) return { error: 'target must be {row?, col?, rows?, cols?}' };

    const row = _optInt(raw.row);
    const col = _optInt(raw.col);
    const rows = _optInt(raw.rows);
    const cols = _optInt(raw.cols);
    for (const [name, value] of [['row', row], ['col', col], ['rows', rows], ['cols', cols]]) {
        if (value === 'invalid') return { error: `"${name}" must be a positive integer` };
    }

    let startRow;
    let endRow;
    let startCol;
    let endCol;
    if (row !== null) {
        startRow = row;
        endRow = row + (rows !== null ? rows : 1) - 1;
    } else {
        startRow = bounds.startRow;
        endRow = bounds.endRow;
    }
    if (col !== null) {
        startCol = col;
        endCol = col + (cols !== null ? cols : 1) - 1;
    } else {
        startCol = bounds.startCol;
        endCol = bounds.endCol;
    }
    if (rows !== null && row === null) return { error: '"rows" requires "row"' };
    if (cols !== null && col === null) return { error: '"cols" requires "col"' };

    if (startRow < bounds.startRow || endRow > bounds.endRow
        || startCol < bounds.startCol || endCol > bounds.endCol) {
        return {
            error: `target R${startRow}C${startCol}–R${endRow}C${endCol} is outside the covered region ` +
                `R${bounds.startRow}C${bounds.startCol}–R${bounds.endRow}C${bounds.endCol}`,
        };
    }
    if (startRow > endRow || startCol > endCol) return { error: 'target region is empty (check rows/cols spans)' };
    return { region: { startRow, endRow, startCol, endCol } };
}

/** True when a region spans every column of the table (row-level ops valid). */
export function isFullWidthRegion(region, colCount) {
    return !!region && region.startCol === 1 && region.endCol === colCount;
}

/**
 * Human-readable label for one style op (proposal-card item). Style items
 * carry no before/after text — the label is the whole description.
 *
 * @param {object} op - A normalized style op (type + payload)
 * @returns {string}
 */
export function describeStyleOp(op) {
    const target = op.region
        ? `R${op.region.startRow}C${op.region.startCol}${(op.region.endRow > op.region.startRow || op.region.endCol > op.region.startCol) ? `–R${op.region.endRow}C${op.region.endCol}` : ''}`
        : 'whole table';
    switch (op.type) {
        case 'tableStyle': {
            const parts = [];
            if (op.style) parts.push(`style → ${op.style}`);
            for (const [key, label] of [
                ['bandedRows', 'banded rows'], ['bandedColumns', 'banded columns'],
                ['firstColumn', 'first column emphasis'], ['lastColumn', 'last column emphasis'],
                ['totalRow', 'total row'],
            ]) {
                if (op[key] !== undefined) parts.push(`${label} ${op[key] ? 'on' : 'off'}`);
            }
            return `Table look: ${parts.join(', ')}`;
        }
        case 'borders':
            return `Borders: ${BORDER_SET_KEYS
                .filter((location) => op.borders[location])
                .map((location) => `${location} ${_describeBorder(op.borders[location])}`)
                .join(', ')}`;
        case 'cellFormat': {
            const parts = [];
            if (op.shadingColor) parts.push(`shading ${op.shadingColor}`);
            if (op.horizontalAlignment) parts.push(`align ${op.horizontalAlignment}`);
            if (op.verticalAlignment) parts.push(`valign ${op.verticalAlignment}`);
            return `Format ${target}: ${parts.join(', ')}`;
        }
        case 'font':
            return `Font ${target}: ${_describeFont(op.font)}`;
        case 'headerRow': {
            const parts = [`${op.rows} header row(s)`];
            if (op.font) parts.push(_describeFont(op.font));
            if (op.shadingColor) parts.push(`shading ${op.shadingColor}`);
            return `Header: ${parts.join(', ')}`;
        }
        case 'layout': {
            const parts = [];
            if (op.alignment) parts.push(`table ${op.alignment}`);
            if (op.widthPt) parts.push(`width ${op.widthPt}pt`);
            if (op.autoFitWindow) parts.push('autofit to window');
            if (op.distributeColumns) parts.push('distribute columns evenly');
            if (op.cellPaddingPt !== undefined) parts.push(`cell padding ${op.cellPaddingPt}pt`);
            return `Layout: ${parts.join(', ')}`;
        }
        case 'columnWidths':
            return `Column widths: ${op.widthsPt.map((w) => `${w}pt`).join(', ')}`;
        default:
            return 'Table style change';
    }
}

/** @private */
function _describeBorder(spec) {
    let out = spec.type;
    if (spec.type !== 'none') {
        if (spec.width !== undefined) out += ` ${spec.width}pt`;
        if (spec.color && spec.color !== 'auto') out += ` ${spec.color}`;
    }
    return out;
}

/** @private */
function _describeFont(font) {
    return Object.entries(font)
        .map(([key, value]) => {
            if (key === 'size') return `${value}pt`;
            if (typeof value === 'boolean') return key;
            if (key === 'underline') return `underline ${value}`;
            return `${key} ${value}`;
        })
        .join(', ');
}

/** @private */
function _canonicalize(value, list) {
    if (typeof value !== 'string') return null;
    const wanted = value.trim().replace(/[\s_-]+/g, '').toLowerCase();
    return list.find((entry) => entry.replace(/[\s_-]+/g, '').toLowerCase() === wanted) || null;
}

/** @private */
function _optInt(value) {
    if (value === undefined || value === null) return null;
    const n = typeof value === 'string' && value.trim() !== '' ? Number(value) : value;
    if (!Number.isInteger(n) || n < 1) return 'invalid';
    return n;
}
