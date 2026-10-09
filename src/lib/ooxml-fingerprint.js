/** Stable paragraph comparison for document-edit anchors. */
import { documentPartRoot } from './ooxml-text.js';

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const W14_NS = 'http://schemas.microsoft.com/office/word/2010/wordml';
const XML_NS = 'http://www.w3.org/XML/1998/namespace';
const XMLNS_NS = 'http://www.w3.org/2000/xmlns/';
const ON_OFF_PROPERTIES = new Set(['b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps', 'strike', 'dstrike',
    'outline', 'shadow', 'emboss', 'imprint', 'noProof', 'snapToGrid', 'vanish', 'webHidden',
    'rtl', 'cs', 'keepNext', 'keepLines', 'pageBreakBefore', 'widowControl', 'contextualSpacing']);

/**
 * Compares a paragraph's content and formatting across independent Word.js
 * reads. getOoxml() includes package parts and Word-generated proofing,
 * layout, bookmark and revision-session identifiers which can change without
 * an edit to the paragraph. Keep every other element and attribute so real
 * text, formatting, revisions and protected structure still invalidate an
 * edit anchor. Null means the range cannot be verified as one paragraph.
 *
 * @param {string} ooxml
 * @returns {string|null}
 */
export function paragraphStructureFingerprint(ooxml) {
    const root = documentPartRoot(ooxml);
    if (!root) return null;
    const paragraphs = root.namespaceURI === W_NS && root.localName === 'p'
        ? [root] : Array.from(root.getElementsByTagNameNS(W_NS, 'p'));
    if (paragraphs.length !== 1) return null;

    return JSON.stringify(stableNode(paragraphs[0]));
}

/**
 * Compare formatting CONTENT, independent of Word's range-export layout.
 * The export's trailing body sectPr is a document container, outside the
 * range's paragraphs; actual section breaks inside pPr remain significant.
 * Same-format text runs may split at bookmarks/proofing boundaries without
 * changing a character's formatting. Do not relax the prose-edit comparator.
 */
export function rangeStructureFingerprint(ooxml) {
    const root = documentPartRoot(ooxml);
    if (!root) return null;
    const body = root.namespaceURI === W_NS && root.localName === 'body'
        ? root : root.getElementsByTagNameNS(W_NS, 'body')[0];
    const content = body || root;
    if (content.namespaceURI !== W_NS) return null;
    const nodes = content.localName === 'body' ? Array.from(content.children)
        .filter((node) => node.namespaceURI !== W_NS || node.localName !== 'sectPr') : [content];
    return JSON.stringify(nodes.map((node) => formatNode(stableNode(node))).filter(Boolean));
}

function isNode(node, name) { return Array.isArray(node) && node[0] === W_NS && node[1] === name; }

/** Only plain text runs merge; fields, drawings, references and revisions keep boundaries. */
function plainRun(node) {
    if (!isNode(node, 'r')) return null;
    const properties = node[3].find((child) => isNode(child, 'rPr'));
    if (properties && JSON.stringify(properties).includes('PrChange')) return null;
    const content = node[3].filter((child) => !isNode(child, 'rPr'));
    if (!content.every((child) => isNode(child, 't') && !child[2].length
        && child[3].every((text) => text[0] === 'text'))) return null;
    return { key: JSON.stringify([node[2], properties || null]), properties,
        text: content.flatMap((child) => child[3]).map((text) => text[1]).join('') };
}

function formatNode(node) {
    if (!node || node[0] === 'text') return node;
    const [namespace, name, originalAttrs, children] = node;
    let attrs = originalAttrs;
    if (namespace === W_NS && ON_OFF_PROPERTIES.has(name)) {
        const value = attrs.find((attr) => attr[0] === W_NS && attr[1] === 'val')?.[2] ?? 'true';
        if (['true', '1', 'on', 'false', '0', 'off'].includes(value)) {
            attrs = [...attrs.filter((attr) => !(attr[0] === W_NS && attr[1] === 'val')),
                [W_NS, 'val', ['true', '1', 'on'].includes(value) ? 'true' : 'false']]
                .sort((a, b) => JSON.stringify(a).localeCompare(JSON.stringify(b)));
        }
    }
    let normalized = children.map(formatNode).filter(Boolean);
    if (namespace === W_NS && ['rPr', 'pPr'].includes(name)) {
        if (!attrs.length && !normalized.length) return null;
        // Distinct properties are unordered. Stable sort preserves duplicates,
        // so conflicting repeated properties cannot silently change precedence.
        normalized.sort((a, b) => JSON.stringify(a.slice(0, 2)).localeCompare(JSON.stringify(b.slice(0, 2))));
    }
    const combined = [];
    for (const child of normalized) {
        const run = plainRun(child);
        if (run && !run.text) continue; // empty export runs carry no characters
        const previous = combined[combined.length - 1];
        const priorRun = plainRun(previous);
        if (run && priorRun?.key === run.key) {
            previous[3] = [...(run.properties ? [run.properties] : []),
                [W_NS, 't', [], [['text', priorRun.text + run.text]]]];
        } else if (isNode(child, 't') && isNode(previous, 't') && !child[2].length && !previous[2].length
            && child[3].every((text) => text[0] === 'text') && previous[3].every((text) => text[0] === 'text')) {
            previous[3] = [['text', [...previous[3], ...child[3]].map((text) => text[1]).join('')]];
        } else combined.push(child);
    }
    if (namespace === W_NS && name === 't' && combined.every((text) => text[0] === 'text')) {
        return [namespace, name, attrs, [['text', combined.map((text) => text[1]).join('')]]];
    }
    return [namespace, name, attrs, combined];
}

/** Describe mismatch locations without recording document text or XML. */
export function rangeFingerprintDifference(before, after) {
    if (before === after) return '';
    if (!before || !after) return 'unreadable range structure';
    function difference(a, b, path) {
        if (JSON.stringify(a) === JSON.stringify(b)) return '';
        if (!Array.isArray(a) || !Array.isArray(b)) return path;
        if (a.length === 4 && b.length === 4 && typeof a[1] === 'string' && typeof b[1] === 'string') {
            const location = `${path}/${a[1]}`;
            if (a[0] !== b[0] || a[1] !== b[1]) return `${location}: element`;
            if (JSON.stringify(a[2]) !== JSON.stringify(b[2])) {
                const attributes = (values) => new Map(values.map(([ns, name, value]) => [JSON.stringify([ns, name]), value]));
                const oldAttrs = attributes(a[2]); const newAttrs = attributes(b[2]);
                const names = [...new Set([...oldAttrs.keys(), ...newAttrs.keys()])];
                const changed = names.filter((name) => oldAttrs.get(name) !== newAttrs.get(name)).slice(0, 3).map((name) => {
                    const [namespace, localName] = JSON.parse(name);
                    const prefix = namespace === W_NS ? 'w' : namespace === W14_NS ? 'w14'
                        : namespace === 'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing' ? 'wp14'
                            : namespace === 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing' ? 'wp' : 'other';
                    return `${prefix}:${localName} ${!oldAttrs.has(name) ? 'added' : !newAttrs.has(name) ? 'removed' : 'changed'}`;
                });
                return `${location}: attributes (${changed.join(', ')})`;
            }
            return difference(a[3], b[3], location);
        }
        if (a.length !== b.length) return `${path}: item count ${a.length} → ${b.length}`;
        for (let i = 0; i < a.length; i++) {
            const changed = difference(a[i], b[i], `${path}[${i + 1}]`);
            if (changed) return changed;
        }
        return path;
    }
    try { return difference(JSON.parse(before), JSON.parse(after), 'scope').slice(0, 200); }
    catch { return 'unreadable range structure'; }
}

function stableNode(node) {
    if (node.nodeType === 3 || node.nodeType === 4) {
        const value = node.nodeValue || '';
        // Formatting whitespace outside text-bearing elements is XML
        // serialization, not document content.
        return value.trim() || ['t', 'delText', 'instrText'].includes(node.parentNode?.localName)
            ? ['text', value] : null;
    }
    if (node.nodeType !== 1) return null;
    const element = /** @type {Element} */ (node);
    if (element.namespaceURI === W_NS && ['proofErr', 'bookmarkStart', 'bookmarkEnd',
        'lastRenderedPageBreak'].includes(element.localName)) return null;
    const attrs = Array.from(element.attributes).filter((attr) => {
        if (attr.namespaceURI === XMLNS_NS) return false;
        if (attr.namespaceURI === XML_NS && attr.localName === 'space') return false;
        if (attr.namespaceURI === W_NS && attr.localName.startsWith('rsid')) return false;
        if (attr.namespaceURI === W14_NS && ['paraId', 'textId'].includes(attr.localName)) return false;
        return true;
    }).map((attr) => [attr.namespaceURI || '', attr.localName, attr.value])
        .sort((a, b) => JSON.stringify(a).localeCompare(JSON.stringify(b)));
    return [element.namespaceURI || '', element.localName, attrs,
        Array.from(element.childNodes).map(stableNode).filter((child) => child !== null)];
}
