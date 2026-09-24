/** Stable paragraph comparison for document-edit anchors. */
import { documentPartRoot } from './ooxml-text.js';

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const W14_NS = 'http://schemas.microsoft.com/office/word/2010/wordml';
const XML_NS = 'http://www.w3.org/XML/1998/namespace';
const XMLNS_NS = 'http://www.w3.org/2000/xmlns/';

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
    return JSON.stringify(stableNode(paragraphs[0]));
}
