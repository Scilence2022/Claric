/** Current text and revision boundaries, without accepting/rejecting Word revisions. */
import { documentPartRoot, symbolText } from './ooxml-text.js';
import { rangeStructureFingerprint } from './ooxml-fingerprint.js';
import { MATH_NS, mathTextState } from './ooxml-math.js';

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
        this.code = error?.code;
        this.errorLocation = error?.debugInfo?.errorLocation || error?.errorLocation;
    }
}

export const normalizeWordText = (text) => String(text || '').replace(/\r\n|\r/g, '\n').replaceAll('\v', '\n').replaceAll('\f', '\n').replaceAll('\u001E', '‑');
export const normalizeRevisionText = (text) => normalizeWordText(text).replace(/\n$/, '');
export const hasNativeFieldCodes = (text) => ['\u0013', '\u0014', '\u0015'].some((code) => text.includes(code));

/** Reconcile Word's text stream with XML evidence, retaining a map from native
 * prefix lengths to visible offsets. Field instructions are matched, never
 * blindly stripped; a '(' is a symbol placeholder only at an actual w:sym.
 * Returns null on any unexplained difference, including ordinary whitespace. */
export function projectNativeText(state, value, { trimEnd = false } = {}) {
    const native = trimEnd ? normalizeRevisionText(value) : normalizeWordText(value);
    const includeCodes = hasNativeFieldCodes(native);
    const offsets = [0];
    const linkCodes = [];
    let position = 0;
    let text = '';
    const tokens = [];
    for (const token of state.nativeTokens || [{ text: state.text }]) {
        const previous = tokens[tokens.length - 1];
        if (token.code !== undefined && previous?.code !== undefined && !token.optional && !previous.optional) previous.code += token.code;
        else tokens.push({ ...token });
    }
    if (trimEnd) {
        let last = tokens.length - 1;
        while (last >= 0 && !tokens[last].text) last--;
        if (last >= 0) tokens[last] = { ...tokens[last], text: tokens[last].text.replace(/\n$/, '') };
    }
    for (const token of tokens) {
        const visible = normalizeWordText(token.text);
        const codes = token.code === undefined ? [] : [token.code, ...(token.codeAlternatives || [])];
        const choices = token.linkEnd ? [linkCodes.pop() ? normalizeWordText(token.code) : ''] : token.code !== undefined
            ? includeCodes ? codes.map(normalizeWordText) : ['']
            : [visible, ...(token.alternatives || [])];
        let matched = choices.find((choice) => native.startsWith(choice, position));
        if (matched === undefined && native[position] === '\u0013' && codes.some((code) => code.startsWith('\u0013')
            && code.endsWith('\u0014') && !hasNativeFieldCodes(code.slice(1, -1)))) {
            const end = native.indexOf('\u0014', position + 1);
            const instruction = native.slice(position + 1, end);
            // Boundary padding in a serialized instruction is not visible
            // text. Keep its internal spacing and quoted arguments exact.
            const unpad = (code) => code.replace(/^ +| +$/g, '');
            if (end >= 0 && !hasNativeFieldCodes(instruction)
                && codes.some((code) => unpad(instruction) === unpad(code.slice(1, -1)))) matched = native.slice(position, end + 1);
        }
        if (matched === undefined && token.optional && !token.linkEnd && native[position] !== '\u0013') matched = '';
        if (matched === undefined) return null;
        if (token.linkStart) linkCodes.push(Boolean(matched));
        for (let i = 1; i <= matched.length; i++) {
            offsets[position + i] = text.length + Math.min(i, visible.length);
        }
        position += matched.length;
        text += visible;
    }
    return position === native.length ? { text, native, offsets } : null;
}

/** Report the mismatch precisely without logging document excerpts. */
export function textMismatchDetail(wanted, actual) {
    let offset = 0;
    while (offset < Math.min(wanted.length, actual.length) && wanted[offset] === actual[offset]) offset++;
    const code = (text) => offset >= text.length ? 'end' : `U+${text.codePointAt(offset).toString(16).toUpperCase().padStart(4, '0')}`;
    return `expected ${wanted.length} chars, received ${actual.length}; first mismatch at ${offset}; expected ${code(wanted)}, received ${code(actual)}`;
}

export function wordErrorDetail(error) {
    const code = error?.code;
    const location = error?.debugInfo?.errorLocation || error?.errorLocation;
    return (error?.message || String(error)) + (code || location
        ? ` [${[code, location].filter(Boolean).map((value) => String(value).slice(0, 160)).join('; ')}]` : '');
}

/** Visible islands never span an insertion boundary or an earlier deletion. */
export function revisionTextState(ooxml, { paragraph = false, currentText = null, paragraphText = null } = {}) {
    const root = documentPartRoot(ooxml);
    if (!root) throw new RevisionSafetyError('Word revision XML is unreadable. Generate a fresh proposal.');
    let body = root.getElementsByTagNameNS(W, 'body')[0] || root;
    // Word can export a whole generated TOC when asked for one of its
    // paragraphs. Resolve only a uniquely proven native paragraph, rather
    // than sending the neighbouring entries to the model repeatedly.
    if (paragraph && (typeof currentText === 'string' || typeof paragraphText === 'string')) {
        const toc = Array.from(body.getElementsByTagNameNS(W, 'sdt')).find((node) =>
            Array.from(node.getElementsByTagNameNS(W, 'docPartGallery')).some((gallery) =>
                gallery.getAttributeNS(W, 'val') === 'Table of Contents'));
        const entries = Array.from((toc || body).getElementsByTagNameNS(W, 'p'));
        const tocEntries = entries.filter((entry) => Array.from(entry.getElementsByTagNameNS(W, 'pStyle'))
            .some((style) => /^TOC[1-9]$/i.test(style.getAttributeNS(W, 'val') || '')));
        if (toc || tocEntries.length > 1) {
            if (entries.length > 1) {
                const states = entries.map((entry) => revisionTextState(new XMLSerializer().serializeToString(entry)));
                let matches = entries.filter((_entry, i) => typeof currentText === 'string'
                    && projectNativeText(states[i], currentText));
                // Reviewed text can expand to the entire TOC too. Paragraph.text
                // still identifies a pristine entry, but is revision-blind, so
                // it must never resolve an entry containing tracked changes.
                if (matches.length !== 1 && typeof paragraphText === 'string') {
                    matches = entries.filter((_entry, i) => !states[i].hasRevisions && !states[i].hasHiddenContent
                        && projectNativeText(states[i], paragraphText));
                }
                if (matches.length !== 1) {
                    if (typeof paragraphText === 'string') {
                        // An unresolvable generated field is readable document
                        // data, not permission to mutate it. Keep the native
                        // paragraph once and protect its entire text. Do not
                        // let this field abort review of the surrounding prose.
                        const text = normalizeWordText(paragraphText);
                        return { text, segments: [], nativeTokens: [{ text }], paragraphCount: 1,
                            hasRevisions: states.some((state) => state.hasRevisions), hasHiddenContent: true,
                            protectedStructure: true, fieldOnlyProtection: false,
                            protectedSpans: [{ start: 0, end: text.length }],
                            fingerprint: JSON.stringify([rangeStructureFingerprint(ooxml), text]) };
                    }
                    throw new RevisionSafetyError('Cannot isolate one native paragraph from the exported table of contents. Read a smaller selection.');
                }
                body = matches[0];
            }
        }
    }
    // Word can append an empty export paragraph to a paragraph's OOXML.
    // Only a known single-paragraph scope permits dropping this boundary;
    // arbitrary ranges must retain their genuine blank paragraphs.
    if (paragraph && body.namespaceURI === W && body.localName === 'body') {
        const paragraphs = Array.from(body.children).filter((node) => node.namespaceURI === W && node.localName === 'p');
        if (paragraphs.length === 2 && (emptyExportParagraph(paragraphs[1])
            || (typeof currentText === 'string' && emptyExportParagraph(paragraphs[1], true)
                && projectNativeText(revisionTextState(new XMLSerializer().serializeToString(paragraphs[0])),
                    currentText)))) body = paragraphs[0];
    }
    const elements = [body, ...Array.from(body.getElementsByTagName('*'))];
    if (!elements.some((element) => element.namespaceURI === W)) {
        throw new RevisionSafetyError('Word revision XML has no document content. Generate a fresh proposal.');
    }
    let hasRevisions = false;
    let protectedStructure = false;
    let nonFieldProtection = false;
    let hasHiddenContent = false;
    let paragraphCount = 0;
    let text = '';
    const segments = [];
    const nativeTokens = [];
    const protectedSpans = [];
    const fields = [];
    const linkedTargets = [];
    // Word's reviewed text may serialize a w:hyperlink as a HYPERLINK field.
    // Resolve only this document's relationship part, never a guessed URL or
    // a same-ID relationship belonging to an image/header/other part.
    const packageNS = 'http://schemas.microsoft.com/office/2006/xmlPackage';
    const relationshipNS = 'http://schemas.openxmlformats.org/package/2006/relationships';
    const links = new Map();
    const relationshipPart = Array.from(root.ownerDocument.getElementsByTagNameNS(packageNS, 'part'))
        .find((part) => part.getAttributeNS(packageNS, 'name') === '/word/_rels/document.xml.rels');
    for (const relation of Array.from(relationshipPart?.getElementsByTagNameNS(relationshipNS, 'Relationship') || [])) {
        if (relation.getAttribute('Type')?.endsWith('/hyperlink') && relation.getAttribute('TargetMode') === 'External') {
            links.set(relation.getAttribute('Id'), relation.getAttribute('Target'));
        }
    }
    let boundary = true;
    function append(value) {
        if (!value) return;
        nativeTokens.push({ text: value });
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
            || (['ins', 'del'].includes(name) && ['trPr', 'tcPr'].includes(parent))) {
            protectedStructure = true;
            if (!['fldChar', 'fldSimple'].includes(name)) nonFieldProtection = true;
        }
    }
    function walk(element) {
        if (element.namespaceURI === MATH_NS && ['oMath', 'oMathPara'].includes(element.localName)) {
            const math = mathTextState(element);
            const start = text.length;
            boundary = true;
            // An equation is an immutable island, never an editable text run.
            text += math.text;
            nativeTokens.push({ text: math.text, alternatives: math.nativeText === null ? [] : [math.nativeText] });
            protectedSpans.push({ start, end: text.length });
            protectedStructure = true;
            if (math.nativeText === null) nonFieldProtection = true;
            boundary = true;
            return;
        }
        if (element.namespaceURI !== W) return;
        const name = element.localName;
        if (name === 'del' || name === 'moveFrom') {
            hasHiddenContent = true;
            boundary = true;
            return;
        }
        if (name === 'delText') { hasHiddenContent = true; boundary = true; return; }
        if (name === 'hyperlink') {
            const id = element.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id');
            const anchor = element.getAttributeNS(W, 'anchor');
            const target = links.get(id);
            linkedTargets.push([id || '', anchor || '', target || '']);
            const instruction = target ? ` HYPERLINK "${target}"${anchor ? ` \\l "${anchor}"` : ''} `
                : anchor ? ` HYPERLINK \\l "${anchor}" ` : null;
            const start = text.length;
            boundary = true;
            if (instruction) nativeTokens.push({ text: '', code: '\u0013' + instruction + '\u0014',
                codeAlternatives: ['\u0013' + instruction.trimEnd() + ' \\h \u0014'], optional: true, linkStart: true });
            for (const child of Array.from(element.children)) walk(child);
            if (instruction) nativeTokens.push({ text: '', code: '\u0015', optional: true, linkEnd: true });
            protectedSpans.push({ start, end: text.length });
            boundary = true;
            return;
        }
        if (name === 'fldChar') {
            const type = element.getAttributeNS(W, 'fldCharType') || element.getAttribute('w:fldCharType');
            const code = { begin: '\u0013', separate: '\u0014', end: '\u0015' }[type];
            if (code) nativeTokens.push({ text: '', code });
            boundary = true;
            if (type === 'begin') fields.push(text.length);
            else if (type === 'separate' && !fields.length) fields.push(0);
            else if (type === 'end') protectedSpans.push({ start: fields.length ? fields.pop() : 0, end: text.length });
            return;
        }
        if (name === 'fldSimple') {
            const start = text.length;
            const instruction = element.getAttributeNS(W, 'instr') || element.getAttribute('w:instr') || '';
            nativeTokens.push({ text: '', code: '\u0013' + instruction + '\u0014' });
            boundary = true;
            for (const child of Array.from(element.children)) walk(child);
            nativeTokens.push({ text: '', code: '\u0015' });
            protectedSpans.push({ start, end: text.length });
            boundary = true;
            return;
        }
        if (name === 'instrText') {
            nativeTokens.push({ text: '', code: element.textContent || '' });
            return;
        }
        if (['pPr', 'rPr', 'proofErr', 'drawing', 'object', 'pict'].includes(name)) return;
        if (name === 'p') {
            if (paragraphCount++) { text += '\n'; nativeTokens.push({ text: '\n' }); boundary = true; }
        }
        if (name === 'ins' || name === 'moveTo') boundary = true;
        if (name === 't') append(element.textContent || '');
        else if (name === 'tab') append('\t');
        else if (name === 'br' || name === 'cr') append('\n');
        else if (name === 'noBreakHyphen') append('‑');
        else if (name === 'sym') {
            const start = text.length;
            boundary = true;
            append(symbolText(element));
            // Word's legacy chSymbol is U+0028. Other hosts return the glyph
            // or its font-specific private-use slot.
            const code = element.getAttributeNS(W, 'char') || element.getAttribute('w:char') || '';
            nativeTokens[nativeTokens.length - 1].alternatives = ['(',
                ...(/^[0-9a-f]{4}$/i.test(code) ? [String.fromCharCode(parseInt(code, 16))] : [])];
            protectedSpans.push({ start, end: text.length });
            boundary = true;
        }
        else for (const child of Array.from(element.children)) walk(child);
        if (name === 'ins' || name === 'moveTo') boundary = true;
    }
    // pkg:xmlData is a wrapper, not a Word element.
    if (body.namespaceURI === W) walk(body);
    else for (const child of Array.from(body.children)) walk(child);
    for (const start of fields) protectedSpans.push({ start, end: text.length });
    // TOC titles and page numbers are generated field results, including
    // entries whose paragraph export omits the outer multi-paragraph field.
    const scopeParagraph = body.localName === 'p' ? body
        : paragraphCount === 1 ? body.getElementsByTagNameNS(W, 'p')[0] : null;
    const tocStyle = scopeParagraph && Array.from(scopeParagraph.getElementsByTagNameNS(W, 'pStyle'))
        .some((style) => /^TOC[1-9]$/i.test(style.getAttributeNS(W, 'val') || ''));
    if (tocStyle) protectedSpans.push({ start: 0, end: text.length });
    const structure = rangeStructureFingerprint(body.localName === 'p' && body !== root
        ? new XMLSerializer().serializeToString(body) : ooxml);
    return { text, segments, nativeTokens, hasRevisions, hasHiddenContent, protectedStructure, paragraphCount,
        protectedSpans, fieldOnlyProtection: protectedStructure && !nonFieldProtection,
        fingerprint: structure && linkedTargets.length ? JSON.stringify([structure, linkedTargets]) : structure };
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
    // Keep XML exports on the caller's proxy. On Mac, exporting Content
    // ranges of blank/anchored paragraphs can fail at sync with a generic
    // host error. Native text still needs Content to exclude the paragraph
    // mark; expanded TOC XML is scoped using that independent text below.
    const xml = range.getOoxml();
    const content = options.paragraph && typeof range.getRange === 'function' ? range.getRange('Content') : range;
    const currentText = queueCurrentTextRead(content);
    // Preserve the ClientResult value contract while pairing independent
    // evidence that a revised, empty export paragraph is not native content.
    return currentText ? { get value() { return xml.value; }, currentText } : xml;
}

export function resolveRevisionRead(range, read, options = {}) {
    let paragraphText = null;
    // Parser/re-anchoring callers already load text. Search ranges may not;
    // never turn an optional scoping hint into a PropertyNotLoaded failure.
    if (options.paragraph) {
        try { if (typeof range.text === 'string') paragraphText = range.text; } catch { /* not loaded */ }
    }
    return read ? revisionTextState(read.value, { ...options, paragraphText, currentText: read.currentText?.value }) : { text: range.text || '', hasRevisions: false,
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

/**
 * True when this Word host can write tracked changes (WordApi 1.4:
 * Document.changeTrackingMode). The Word.ChangeTrackingMode enum ships with
 * office.js on every host, so on its own it is not evidence of support; the
 * requirement set is checked when Office exposes it.
 *
 * @returns {boolean}
 */
export function hostSupportsTrackedChanges() {
    if (typeof Word === 'undefined' || !Word.ChangeTrackingMode) return false;
    const requirements = typeof Office !== 'undefined' ? Office.context?.requirements : null;
    return typeof requirements?.isSetSupported === 'function' ? requirements.isSetSupported('WordApi', '1.4') : true;
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
