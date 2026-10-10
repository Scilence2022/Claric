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

/** A Word write may have reached the host. Retrying an old baseline is unsafe. */
export class MutationSafetyError extends RevisionSafetyError {
    constructor(error) {
        super(error?.message || String(error));
        this.name = 'MutationSafetyError';
        this.mutationAttempted = true;
        this.cause = error;
    }
}

export const normalizeRevisionText = (text) => String(text || '').replace(/\r\n|\r/g, '\n').replace(/\n$/, '');

/** Visible islands never span an insertion boundary or an earlier deletion. */
export function revisionTextState(ooxml, { paragraph = false, currentText = null } = {}) {
    const root = documentPartRoot(ooxml);
    if (!root) throw new RevisionSafetyError('Word revision XML is unreadable. Generate a fresh proposal.');
    let body = root.getElementsByTagNameNS(W, 'body')[0] || root;
    // Word can append an empty export paragraph to a paragraph's OOXML.
    // Only a known single-paragraph scope permits dropping this boundary;
    // arbitrary ranges must retain their genuine blank paragraphs.
    if (paragraph && body.namespaceURI === W && body.localName === 'body') {
        const paragraphs = Array.from(body.children).filter((node) => node.namespaceURI === W && node.localName === 'p');
        if (paragraphs.length === 2 && (emptyExportParagraph(paragraphs[1])
            || (typeof currentText === 'string' && emptyExportParagraph(paragraphs[1], true)
                && revisionTextState(new XMLSerializer().serializeToString(paragraphs[0])).text
                    === currentText.replace(/\r\n|\r/g, '\n')))) body = paragraphs[0];
    }
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
        fingerprint: rangeStructureFingerprint(body.localName === 'p' && body !== root
            ? new XMLSerializer().serializeToString(body) : ooxml) };
}

function emptyExportParagraph(paragraph, allowEmptyRevisionMarkers = false) {
    return Array.from(paragraph.getElementsByTagName('*')).every((node) => {
        if (node.namespaceURI !== W) return false;
        if (PROTECTED.has(node.localName) || (REVISIONS.has(node.localName)
            && !(allowEmptyRevisionMarkers && ['ins', 'del'].includes(node.localName)))
            || ['br', 'cr', 'tab', 'noBreakHyphen', 'sym', 'sectPr'].includes(node.localName)) return false;
        return !['t', 'delText', 'instrText'].includes(node.localName) || !(node.textContent || '');
    });
}

/** Queue alongside existing loads, then resolve after the caller's sync. */
export function queueRevisionRead(range, options = {}) {
    if (typeof range.getOoxml !== 'function') return null;
    const xml = range.getOoxml();
    const currentText = options.paragraph
        ? queueCurrentTextRead(typeof range.getRange === 'function' ? range.getRange('Content') : range) : null;
    // Preserve the ClientResult value contract while pairing independent
    // evidence that a revised, empty export paragraph is not native content.
    return currentText ? { get value() { return xml.value; }, currentText } : xml;
}

export function resolveRevisionRead(range, read, options = {}) {
    return read ? revisionTextState(read.value, { ...options, currentText: read.currentText?.value }) : { text: range.text || '', hasRevisions: false,
        fingerprint: null, protectedStructure: false };
}

/** Load read-back code before queuing writes, so a missing asset is read-only. */
export async function loadMutationVerifier() {
    return (await import(/* webpackChunkName: "word-mutation-verification" */ './word-mutation-verification.js')).verifyMutationText;
}

/** All callers share the same verifier, including paragraph replacement fallbacks. */
export async function verifyMutationText(...args) {
    return (await loadMutationVerifier())(...args);
}

/** Confirm tracking before any strategy queues its first text mutation. */
export async function enableTrackedWrites(context) {
    try {
        context.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
        await context.sync();
    } catch (error) {
        throw new RevisionSafetyError(`Could not enable tracked changes (${error.message}); no text was written.`);
    }
}

/** WordApi 1.4's final text avoids OOXML export-only paragraph boundaries. */
export function queueCurrentTextRead(range) {
    if (typeof range.getReviewedText !== 'function') return null;
    if (typeof Office !== 'undefined' && Office.context?.requirements?.isSetSupported
        && !Office.context.requirements.isSetSupported('WordApi', '1.4')) return null;
    return range.getReviewedText(typeof Word !== 'undefined' ? Word.ChangeTrackingVersion?.current || 'Current' : 'Current');
}
