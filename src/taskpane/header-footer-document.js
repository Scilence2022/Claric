import JSZip from 'jszip';
import { rangeStructureFingerprint } from '../lib/ooxml-fingerprint.js';
import { readWordFile } from './word-file.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const RELS = 'http://schemas.openxmlformats.org/package/2006/relationships';
export const VARIANTS = { default: 'Primary', first: 'FirstPage', even: 'EvenPages' };
const KINDS = ['header', 'footer'];
const MAX_DOCX_BYTES = 64 * 1024 * 1024;
const children = (node, name) => Array.from(node?.children || []).filter((el) => el.namespaceURI === W && el.localName === name);
const child = (node, name) => children(node, name)[0];
const attr = (node, name) => node?.getAttributeNS(W, name);
const enabled = (node) => !!node && !['0', 'false', 'off'].includes(attr(node, 'val'));
const serialize = (node) => new XMLSerializer().serializeToString(node);

function parse(xml, ns, name) {
    const root = new DOMParser().parseFromString(xml, 'text/xml').documentElement;
    if (root.namespaceURI !== ns || root.localName !== name) throw new Error('Word returned an unsupported document structure.');
    return root;
}

async function part(zip, name, ns, root, optional = false) {
    const entry = zip.file(name);
    if (!entry && optional) return null;
    if (!entry) throw new Error('Word document is missing a part needed to verify headers and footers.');
    return parse(await entry.async('string'), ns, root);
}

function partPath(target) {
    if (!target || !/^(?:\/word\/)?[a-z0-9_./-]+\.xml$/i.test(target) || target.split('/').includes('..')) {
        throw new Error('Word returned an unsupported header/footer relationship.');
    }
    return target.startsWith('/') ? target.slice(1) : `word/${target}`;
}

const fieldName = (code) => code.trim().match(/^\w+/)?.[0].toUpperCase() || 'FIELD';
export const fieldToken = (code) => `{{${code}}}`;

/** Text that survives field recalculation, with editable direct-run spans. */
function inspectStory(root) {
    if (!root) return { fingerprint: '', paragraphs: [], fields: [], objects: [], spans: [], simple: true };
    const fields = [], objects = [], spans = [];
    const paragraphs = children(root, 'p');
    const stack = [];
    let text = '', span = '';
    const flush = () => { if (span) spans.push(span); span = ''; };
    const recordField = (code) => {
        const name = fieldName(code);
        fields.push(name);
        text += fieldToken(name);
    };
    function visit(node, editable) {
        const name = node.localName;
        if (['pPr', 'rPr', 'proofErr', 'lastRenderedPageBreak'].includes(name)) return;
        if (['del', 'moveFrom', 'bookmarkStart', 'bookmarkEnd'].includes(name)) { flush(); return; }
        if (name === 'fldSimple') {
            flush();
            if (!stack.length) recordField(attr(node, 'instr') || '');
            return;
        }
        if (name === 'fldChar') {
            flush();
            const type = attr(node, 'fldCharType');
            if (type === 'begin') stack.push('');
            if (type === 'end') {
                const code = stack.pop();
                if (code !== undefined && !stack.length) recordField(code);
            }
            return;
        }
        if (stack.length) {
            if (name === 'instrText') stack[stack.length - 1] += node.textContent;
            // Process nested field markers while skipping cached results.
            for (const item of node.children) visit(item, false);
            return;
        }
        if (['drawing', 'pict', 'object'].includes(name)) {
            flush();
            objects.push(name);
            text += '￼';
            return;
        }
        if (name === 't') {
            text += node.textContent;
            if (editable) span += node.textContent;
            return;
        }
        if (name === 'tab' || name === 'br' || name === 'cr') {
            flush();
            text += name === 'tab' ? '\t' : '\n';
            return;
        }
        const direct = editable && name === 'r' && node.parentElement.parentElement === root;
        if (!direct) flush();
        for (const item of node.children) visit(item, direct);
        if (!direct) flush();
    }
    const lines = Array.from(root.children).map((block) => {
        text = '';
        for (const node of (block.localName === 'p' ? block.children : [block])) visit(node, block.localName === 'p');
        flush();
        return text;
    });
    const simple = root.children.length === paragraphs.length
        && Array.from(root.getElementsByTagName('*')).every((node) => node.namespaceURI === W
            && !['sdt', 'customXml', 'ins', 'del', 'moveFrom', 'moveTo', 'drawing', 'pict', 'object', 'hyperlink',
                'bookmarkStart', 'bookmarkEnd', 'commentRangeStart', 'commentRangeEnd'].includes(node.localName));
    return { fingerprint: rangeStructureFingerprint(serialize(root)), paragraphs: lines, fields, spans, objects, simple };
}

// Read actual package relationships. Independent stories may have identical
// text, and an absent reference inherits its preceding section's story.
export async function readHeaderFooterDocument({ signal, log } = {}) {
    const bytes = await readWordFile(Office.FileType.Compressed, { label: 'Word DOCX', maxBytes: MAX_DOCX_BYTES, signal, log });
    const zip = await JSZip.loadAsync(bytes);
    const [document, settings, relationships] = await Promise.all([
        part(zip, 'word/document.xml', W, 'document'),
        part(zip, 'word/settings.xml', W, 'settings', true),
        part(zip, 'word/_rels/document.xml.rels', RELS, 'Relationships', true),
    ]);
    signal?.throwIfAborted();
    const body = child(document, 'body');
    if (!body) throw new Error('Word returned an unsupported document body.');
    const sectionProperties = Array.from(body.getElementsByTagNameNS(W, 'sectPr')).filter((node) =>
        node.parentNode === body || (node.parentNode.localName === 'pPr'
            && node.parentNode.parentNode.localName === 'p'));
    if (!sectionProperties.length || sectionProperties.length > 100 || sectionProperties.at(-1).parentNode !== body) {
        throw new Error('Word did not expose all section boundaries for header/footer editing.');
    }
    const layout = JSON.stringify(sectionProperties.map((node) => {
        const copy = node.cloneNode(true);
        for (const name of ['headerReference', 'footerReference']) for (const ref of children(copy, name)) ref.remove();
        return rangeStructureFingerprint(serialize(copy));
    }));
    const related = new Map(Array.from(relationships?.children || []).map((node) => [node.getAttribute('Id'), node]));
    const stories = new Map();
    const inherited = new Map();
    const even = enabled(child(settings, 'evenAndOddHeaders'));
    const protectedDocument = enabled(child(settings, 'documentProtection'))
        && ['1', 'true', 'on'].includes(attr(child(settings, 'documentProtection'), 'enforcement'));
    const slots = [];
    for (const [index, properties] of sectionProperties.entries()) {
        for (const kind of KINDS) for (const variant of Object.keys(VARIANTS)) {
            const slot = `${kind}:${variant}`;
            const refs = children(properties, `${kind}Reference`).filter((ref) => attr(ref, 'type') === variant);
            if (refs.length > 1) throw new Error('Word returned ambiguous header/footer relationships.');
            if (refs.length) {
                const relation = related.get(refs[0].getAttributeNS(R, 'id'));
                if (!relation || relation.getAttribute('Type') !== `${R}/${kind}` || relation.getAttribute('TargetMode') === 'External') {
                    throw new Error('Word could not verify a header/footer relationship.');
                }
                inherited.set(slot, partPath(relation.getAttribute('Target')));
            }
            const path = inherited.get(slot);
            const key = path || slot;
            const id = `s${index + 1}:${slot}`;
            if (!stories.has(key)) {
                const root = path ? await part(zip, path, W, kind === 'header' ? 'hdr' : 'ftr') : null;
                stories.set(key, { id, path, ...inspectStory(root) });
            }
            const story = stories.get(key);
            slots.push({ id, storyId: story.id, section: index + 1, kind, variant,
                active: variant === 'default' || (variant === 'first' ? enabled(child(properties, 'titlePg')) : even),
                story });
        }
    }
    for (const node of Array.from(body.getElementsByTagNameNS(W, 'sectPr'))) node.remove();
    return { slots, sectionCount: sectionProperties.length, layout, protectedDocument,
        bodyFingerprint: rangeStructureFingerprint(serialize(body)) };
}

export function topology(snapshot) {
    return JSON.stringify(snapshot.slots.map(({ id, storyId, active }) => [id, storyId, active]));
}
