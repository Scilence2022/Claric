/** Unicode-safe edit planning shared by draft checks and native Word writes. */
import DiffMatchPatch from './vendor/diff-match-patch.js';

export function diffUnicodeScalars(before, after) {
    const values = [''];
    const ids = new Map();
    const encode = (text) => Array.from(text || '').map((scalar) => {
        if (!ids.has(scalar)) { ids.set(scalar, values.length); values.push(scalar); }
        if (values.length > 0xd7ff) throw new Error('Too many distinct characters for a safe text diff.');
        return String.fromCharCode(ids.get(scalar));
    }).join('');
    return new DiffMatchPatch().diff_main(encode(before), encode(after), false)
        .map(([op, encoded]) => [op, Array.from(encoded).map((id) => values[id.charCodeAt(0)]).join('')]);
}

export function textEditHunks(before, after) {
    const hunks = [];
    let offset = 0;
    let pending = null;
    for (const [op, text] of diffUnicodeScalars(before, after)) {
        if (op === 0) { pending = null; offset += text.length; continue; }
        if (!pending) { pending = { start: offset, end: offset, text: '' }; hunks.push(pending); }
        if (op === -1) { offset += text.length; pending.end = offset; }
        else pending.text += text;
    }
    return hunks;
}

/** Native fields are immutable, including empty fields crossed by an edit. */
export function touchesProtectedSpan(hunk, spans = []) {
    return spans.some(({ start, end }) => hunk.start === hunk.end
        ? hunk.start > start && hunk.start < end
        : start === end ? hunk.start < start && hunk.end > start
            : hunk.start < end && hunk.end > start);
}
