/** Current text and revision boundaries, without accepting/rejecting Word revisions. */
import { documentPartRoot } from './ooxml-text.js';
import { rangeStructureFingerprint } from './ooxml-fingerprint.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const REVISIONS = new Set(['ins', 'del', 'delText', 'moveFrom', 'moveTo', 'rPrChange', 'pPrChange',
    'sectPrChange', 'tblPrChange', 'trPrChange', 'tcPrChange', 'cellIns', 'cellDel', 'cellMerge']);
const PROTECTED = new Set(['drawing', 'object', 'pict', 'fldChar', 'fldSimple', 'sdt',
    'footnoteReference', 'endnoteReference', 'oMath', 'tbl', 'sectPrChange', 'pPrChange',
    'tblPrChange', 'trPrChange', 'tcPrChange', 'cellIns', 'cellDel', 'cellMerge', 'moveFrom', 'moveTo']);

export class RevisionSafetyError extends Error {
    constructor(message) {
        super(message);
        this.name = 'RevisionSafetyError';
    }
}

export const normalizeRevisionText = (text) => String(text || '').replace(/\r\n|\r/g, '\n').replace(/\n$/, '');

/** Visible islands never span an insertion boundary or an earlier deletion. */
export function revisionTextState(ooxml) {
    const root = documentPartRoot(ooxml);
    if (!root) throw new RevisionSafetyError('Word revision XML is unreadable. Generate a fresh proposal.');
    const body = root.getElementsByTagNameNS(W, 'body')[0] || root;
    const elements = [body, ...Array.from(body.getElementsByTagName('*'))];
    if (!elements.some((element) => element.namespaceURI === W)) {
        throw new RevisionSafetyError('Word revision XML has no document content. Generate a fresh proposal.');
    }
    let hasRevisions = false;
    let protectedStructure = false;
    let hasHiddenContent = false;
    let paragraphCount = 0;
    let text = '';
    const segments = [];
    let boundary = true;
    function append(value) {
        if (!value) return;
        const previous = segments[segments.length - 1];
        if (!boundary && previous) previous.text += value;
        else segments.push({ start: text.length, text: value });
        text += value;
        boundary = false;
    }
    for (const element of elements) {
        if (element.namespaceURI !== W) continue;
        const name = element.localName;
        if (REVISIONS.has(name)) hasRevisions = true;
        const parent = element.parentElement?.localName;
        if (PROTECTED.has(name) || (name === 'del' && ['rPr', 'pPr'].includes(parent))
            || (['ins', 'del'].includes(name) && ['trPr', 'tcPr'].includes(parent))) protectedStructure = true;
    }
    function walk(element) {
        if (element.namespaceURI !== W) return;
        const name = element.localName;
        if (name === 'del' || name === 'moveFrom') {
            hasHiddenContent = true;
            boundary = true;
            return;
        }
        if (name === 'delText') { hasHiddenContent = true; boundary = true; return; }
        if (['pPr', 'rPr', 'instrText', 'proofErr', 'drawing', 'object', 'pict'].includes(name)) return;
        if (name === 'p') {
            if (paragraphCount++) { text += '\n'; boundary = true; }
        }
        if (name === 'ins' || name === 'moveTo') boundary = true;
        if (name === 't') append(element.textContent || '');
        else if (name === 'tab') append('\t');
        else if (name === 'br' || name === 'cr') append('\n');
        else if (name === 'noBreakHyphen') append('‑');
        else for (const child of Array.from(element.children)) walk(child);
        if (name === 'ins' || name === 'moveTo') boundary = true;
    }
    // pkg:xmlData is a wrapper, not a Word element.
    if (body.namespaceURI === W) walk(body);
    else for (const child of Array.from(body.children)) walk(child);
    return { text, segments, hasRevisions, hasHiddenContent, protectedStructure, paragraphCount,
        fingerprint: rangeStructureFingerprint(ooxml) };
}

/** Queue alongside existing loads, then resolve after the caller's sync. */
export function queueRevisionRead(range) {
    return typeof range.getOoxml === 'function' ? range.getOoxml() : null;
}

export function resolveRevisionRead(range, read) {
    return read ? revisionTextState(read.value) : { text: range.text || '', hasRevisions: false,
        fingerprint: null, protectedStructure: false };
}
