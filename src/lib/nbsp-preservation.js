/**
 * Models routinely retype U+00A0 as an ordinary space ("P.\u00a0L." becomes
 * "P. L."). That is a transport artifact, not an edit: applied, it turns every
 * reference list into a wall of invisible tracked changes. Put the source's
 * non-breaking space back wherever the model kept the surrounding text.
 *
 * Pure JavaScript -- no Word API dependency.
 */
import DiffMatchPatch from './vendor/diff-match-patch.js';

const NBSP = '\u00a0';
const NBSP_REQUEST_RE = /non[-\s]?breaking[-\s]+spaces?|\bnbsp\b|\\u00a0|不间断空格|不换行空格|不可断行空格|硬空格/i;

/** True when the instruction itself is about non-breaking spaces. */
export function requestsNbspChange(instruction = '') {
    return NBSP_REQUEST_RE.test(String(instruction));
}

/**
 * Restore each source non-breaking space that the amendment kept as unchanged
 * text but typed as U+0020. Inserted text and real rewording are untouched.
 *
 * @param {string} original - Source text
 * @param {string} amended - Model output
 * @returns {string} Amended text with the source's non-breaking spaces restored
 */
export function preserveNonBreakingSpaces(original, amended) {
    const source = String(original ?? '');
    const text = String(amended ?? '');
    if (!source.includes(NBSP)) return text;
    // Compare both sides with U+00A0 read as a space so the space type never
    // steers the alignment; equal runs then map back to both texts by offset.
    let from = 0;
    let to = 0;
    let restored = '';
    for (const [op, part] of new DiffMatchPatch().diff_main(source.replaceAll(NBSP, ' '), text.replaceAll(NBSP, ' '))) {
        if (op === 0) {
            for (let i = 0; i < part.length; i++) restored += source[from + i] === NBSP ? NBSP : text[to + i];
        } else if (op === 1) restored += text.slice(to, to + part.length);
        if (op !== 1) from += part.length;
        if (op !== -1) to += part.length;
    }
    return restored;
}
