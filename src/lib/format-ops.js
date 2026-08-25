/**
 * Format Ops Module
 *
 * Structured formatting changes driven by natural language. The text diff
 * strategies (token map / sentence diff / char diff) only rewrite text — they
 * preserve existing formatting but cannot change it. This module covers the
 * other half: "make this bold and red", "set this paragraph to Heading 2",
 * "center all headings".
 *
 * The LLM is asked to translate the instruction into a JSON array of
 * formatting operations; parseFormatOps validates that output against a
 * strict allowlist so nothing arbitrary reaches the Word API:
 *
 *   [
 *     {
 *       "match": "optional exact substring to format",
 *       "paragraphStyle": "optional built-in style target (e.g. \"heading1\")",
 *       "font": { "bold": true, "color": "#FF0000", ... },
 *       "paragraph": { "styleBuiltIn": "heading2", "alignment": "centered", ... }
 *     }
 *   ]
 *
 * Pure module — no DOM, no Word API. Safe to import under Jest/node.
 *
 * @module format-ops
 */

/** Font properties: boolean flags. */
const FONT_BOOL_KEYS = ['bold', 'italic', 'strikeThrough', 'doubleStrikeThrough', 'superscript', 'subscript', 'allCaps', 'smallCaps'];
/** Font properties: Word enum names (underline, highlightColor). */
const FONT_ENUM_KEYS = ['underline', 'highlightColor'];
/** Font properties: free strings (name) or #RRGGBB (color). */
const FONT_NAME_KEYS = ['name'];
/** Font properties: point sizes. */
const FONT_NUMBER_KEYS = ['size'];
/** Paragraph properties: enum/style strings. */
const PARA_STRING_KEYS = ['style', 'styleBuiltIn', 'alignment'];
/** Paragraph properties: point values. */
const PARA_NUMBER_KEYS = ['lineSpacing', 'spaceBefore', 'spaceAfter', 'leftIndent', 'rightIndent', 'firstLineIndent'];

const COLOR_RE = /^#[0-9a-fA-F]{6}$/;

/**
 * Builds the LLM prompt that turns a formatting instruction into JSON ops.
 *
 * @param {string} instruction - The user's formatting instruction
 * @param {string} scopeText - Text of the selection or document (match source)
 * @param {string} scope - 'selection' | 'document'
 * @returns {string}
 */
export function buildFormatPrompt(instruction, scopeText, scope) {
    const scopeName = scope === 'document' ? 'document' : 'selection';
    return (
        'You are a formatting assistant embedded in Microsoft Word. The user describes FORMATTING changes ' +
        '(font, size, color, highlight, underline, paragraph style, alignment, spacing, indentation). ' +
        'Translate the instruction into a JSON array of formatting operations.\n\n' +
        'OUTPUT CONTRACT (strict):\n' +
        '- Output ONLY a JSON array. No markdown, no code fences, no explanations, no commentary.\n' +
        '- Each array item is one operation:\n' +
        '  {\n' +
        '    "match": "optional exact substring of the text to format (use one op per distinct target)",\n' +
        '    "paragraphStyle": "optional built-in style of paragraphs to target (e.g. \\"heading1\\")",\n' +
        '    "font": { "bold": true, "italic": true, "underline": "single|double|none", "strikeThrough": true, ' +
        '"superscript": false, "subscript": false, "allCaps": false, "smallCaps": false, "color": "#RRGGBB", ' +
        '"highlightColor": "yellow|green|cyan|magenta|red|blue|darkBlue|darkGreen|darkRed|darkYellow|darkCyan|darkMagenta|black|white", ' +
        '"name": "font name", "size": 12 },\n' +
        '    "paragraph": { "styleBuiltIn": "normal|noSpacing|heading1|heading2|...|heading9|title|subtitle|quote|intenseQuote|listParagraph", ' +
        '"style": "custom style name", "alignment": "left|centered|right|justified", ' +
        '"lineSpacing": 14, "spaceBefore": 6, "spaceAfter": 6, "leftIndent": 18, "rightIndent": 18, "firstLineIndent": 24 }\n' +
        '  }\n' +
        `- Omit both "match" and "paragraphStyle" to target the entire ${scopeName}.\n` +
        '- Include ONLY the properties the user asked to change; spacing/indent values are points.\n' +
        '- Do NOT rewrite or change any text — formatting only. If the instruction asks for content ' +
        'changes instead of formatting, output exactly [].\n\n' +
        'USER INSTRUCTION:\n' + (instruction || '').trim() + '\n\n' +
        `--- ${scopeName.toUpperCase()} TEXT (for choosing "match" substrings) ---\n` + (scopeText || '')
    );
}

/**
 * Parses and validates the model's JSON formatting ops. Tolerates code
 * fences and surrounding prose; drops malformed entries and unknown
 * properties rather than failing the whole batch.
 *
 * @param {string} raw - Raw model output
 * @param {function} [log] - Logging callback
 * @returns {Array<object>} Sanitized ops (possibly empty)
 */
export function parseFormatOps(raw, log = () => {}) {
    if (!raw) return [];
    let text = String(raw).trim();

    const fence = text.match(/```(?:json)?\s*([\s\S]*?)```/);
    if (fence) text = fence[1].trim();

    const start = text.indexOf('[');
    const end = text.lastIndexOf(']');
    if (start === -1 || end <= start) {
        log('Format ops: no JSON array found in the model response', 'warning');
        return [];
    }

    let parsed;
    try {
        parsed = JSON.parse(text.slice(start, end + 1));
    } catch (e) {
        log(`Format ops: response is not valid JSON (${e.message})`, 'warning');
        return [];
    }
    if (!Array.isArray(parsed)) return [];

    const ops = [];
    for (const entry of parsed) {
        const op = _sanitizeOp(entry, log);
        if (op) ops.push(op);
    }
    return ops;
}

/**
 * Validates one op entry: target selectors plus font/paragraph payloads.
 * Returns null when nothing applicable remains.
 * @private
 */
function _sanitizeOp(entry, log) {
    if (!entry || typeof entry !== 'object' || Array.isArray(entry)) return null;

    const op = {};
    if (typeof entry.match === 'string' && entry.match.trim()) {
        op.match = entry.match;
    } else if (typeof entry.paragraphStyle === 'string' && entry.paragraphStyle.trim()) {
        op.paragraphStyle = entry.paragraphStyle.trim();
    }

    const font = _sanitizeFont(entry.font, log);
    const paragraph = _sanitizeParagraph(entry.paragraph, log);
    if (font) op.font = font;
    if (paragraph) op.paragraph = paragraph;

    if (!font && !paragraph) {
        log('Format ops: dropped an entry with no valid font/paragraph payload', 'warning');
        return null;
    }
    return op;
}

/** @private */
function _sanitizeFont(font, log) {
    if (!font || typeof font !== 'object' || Array.isArray(font)) return null;
    const out = {};
    for (const key of FONT_BOOL_KEYS) {
        if (font[key] !== undefined) out[key] = !!font[key];
    }
    for (const key of FONT_ENUM_KEYS) {
        if (typeof font[key] === 'string' && font[key].trim()) out[key] = font[key].trim();
    }
    for (const key of FONT_NAME_KEYS) {
        if (typeof font[key] === 'string' && font[key].trim()) out[key] = font[key].trim();
    }
    if (font.color !== undefined) {
        const color = String(font.color).trim();
        if (COLOR_RE.test(color)) {
            out.color = color.toUpperCase();
        } else {
            log(`Format ops: dropped invalid color "${color}" (expected #RRGGBB)`, 'warning');
        }
    }
    for (const key of FONT_NUMBER_KEYS) {
        if (font[key] !== undefined && Number.isFinite(Number(font[key]))) {
            const n = Number(font[key]);
            if (n > 0 && n <= 1638) out[key] = n;
        }
    }
    return Object.keys(out).length > 0 ? out : null;
}

/** @private */
function _sanitizeParagraph(paragraph, log) {
    if (!paragraph || typeof paragraph !== 'object' || Array.isArray(paragraph)) return null;
    const out = {};
    for (const key of PARA_STRING_KEYS) {
        if (typeof paragraph[key] === 'string' && paragraph[key].trim()) out[key] = paragraph[key].trim();
    }
    for (const key of PARA_NUMBER_KEYS) {
        if (paragraph[key] !== undefined && Number.isFinite(Number(paragraph[key]))) {
            const n = Number(paragraph[key]);
            if (Math.abs(n) <= 1638) out[key] = n;
        }
    }
    if (out.style && out.styleBuiltIn) {
        // styleBuiltIn is locale-independent; prefer it when both are given.
        delete out.style;
        log('Format ops: both style and styleBuiltIn given; keeping styleBuiltIn', 'info');
    }
    return Object.keys(out).length > 0 ? out : null;
}
