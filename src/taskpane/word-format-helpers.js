import { nativeHighlightColor } from '../lib/word-format-values.js';

/**
 * Case-insensitive Word enum lookup (e.g. 'heading1' -> Word.BuiltInStyleName.heading1,
 * 'GridTable4_Accent1' ->
 * Word.BuiltInStyleName.gridTable4_Accent1). Separators are stripped on BOTH
 * sides so snake/camel enum keys match their display spellings. Returns
 * undefined on miss.
 * @private
 */
export function enumValue(enumObj, name) {
    if (!enumObj || name === undefined || name === null) return undefined;
    const strip = (s) => String(s).toLowerCase().replace(/[\s_-]+/g, '');
    const wanted = strip(name);
    const key = Object.keys(enumObj).find((k) => strip(k) === wanted);
    return key ? enumObj[key] : undefined;
}

/**
 * Applies validated font ops to a Word.Font object. Unknown/invalid enum
 * values are skipped with a warning rather than failing the batch.
 * @private
 */
export function applyFontOps(font, ops, log) {
    for (const [key, value] of Object.entries(ops)) {
        try {
            if (key === 'underline') {
                const v = enumValue(Word.UnderlineType, value);
                if (v === undefined) log(`Format ops: unknown underline "${value}"`, 'warning');
                else font.underline = v;
            } else if (key === 'highlightColor') {
                font.highlightColor = nativeHighlightColor(value);
            } else {
                font[key] = value;
            }
        } catch (e) {
            log(`Format ops: font.${key} failed (${e.message})`, 'warning');
        }
    }
}
