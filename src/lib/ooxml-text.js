/**
 * OOXML text extraction helpers (pure, namespace-agnostic).
 *
 * Word.js `paragraph.text` / `body.text` is revision-blind: when a paragraph
 * carries tracked changes, deleted text is inlined right next to its
 * replacement, interleaving old and new wording into unreadable word-salad
 * ("TheResults reassemblerare maps resultsmapped..."). These helpers walk the
 * OOXML directly and resolve revisions to their final ("accept-all") state:
 * w:t inside w:del / w:moveFrom containers is hidden, w:ins content is kept.
 *
 * Self-contained on purpose: test suites mock comment-extractor.js wholesale,
 * so importing from there would break under those mocks. DOMParser is
 * available in the add-in WebView and in Jest (jsdom).
 *
 * @module ooxml-text
 */

const PKG_NS = 'http://schemas.microsoft.com/office/2006/xmlPackage';
const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

/** Containers whose subtree is hidden in the final document view. */
const HIDDEN_CONTAINERS = new Set(['del', 'moveFrom']);
/** Field instructions (TOC etc.) — Word.js text excludes them; so do we. */
const SKIP_ELEMENTS = new Set(['instrText', 'proofErr']);

/**
 * Recursively collects visible text of one element in document order.
 *
 * @param {Element} node
 * @returns {string}
 */
function visibleText(node) {
    let text = '';
    for (const child of Array.from(node.childNodes || [])) {
        if (child.nodeType !== 1) continue;
        const element = /** @type {Element} */ (child);
        const name = element.localName;
        if (HIDDEN_CONTAINERS.has(name) || SKIP_ELEMENTS.has(name)) continue;
        if (name === 't') { text += element.textContent || ''; continue; }
        if (name === 'tab') { text += '\t'; continue; }
        if (name === 'br' || name === 'cr') { text += '\n'; continue; }
        if (name === 'noBreakHyphen') { text += '‑'; continue; }
        text += visibleText(element);
    }
    return text;
}

/**
 * Parses OOXML and returns the root element to extract from: the
 * /word/document.xml part when the pkg:package wrapper is present (sibling
 * parts such as comments.xml also carry w:t and would pollute the result),
 * otherwise the document element itself.
 *
 * @param {string} ooxml
 * @returns {Element|null} null when the input is missing or unparseable
 */
function documentPartRoot(ooxml) {
    if (!ooxml || typeof ooxml !== 'string') return null;
    let doc;
    try {
        doc = new DOMParser().parseFromString(ooxml, 'text/xml');
    } catch {
        return null;
    }
    if (doc.getElementsByTagName('parsererror').length > 0) return null;

    let parts = doc.getElementsByTagNameNS(PKG_NS, 'part');
    if (parts.length === 0) parts = doc.getElementsByTagName('pkg:part');
    for (const part of Array.from(parts)) {
        const name = part.getAttributeNS(PKG_NS, 'name')
            || part.getAttribute('pkg:name')
            || part.getAttribute('name')
            || '';
        if (name !== '/word/document.xml') continue;
        const xmlData = part.getElementsByTagNameNS(PKG_NS, 'xmlData')[0]
            || part.getElementsByTagName('pkg:xmlData')[0];
        if (xmlData) return xmlData;
    }
    return doc.documentElement || null;
}

/**
 * Finds the w:body element inside a parsed document part, falling back to
 * the root itself (a bare range fragment has no w:body wrapper).
 *
 * @param {Element} root
 * @returns {Element}
 */
function bodyOrRoot(root) {
    let bodies = root.getElementsByTagNameNS(W_NS, 'body');
    if (bodies.length === 0) bodies = root.getElementsByTagName('w:body');
    return bodies.length > 0 ? bodies[0] : root;
}

/**
 * Extracts the visible ("accept-all") text of an OOXML fragment — a range or
 * paragraph getOoxml() result, with or without the pkg:package wrapper.
 * Tracked deletions / move sources are hidden; insertions are kept.
 *
 * @param {string} ooxml
 * @returns {string|null} Final text (possibly ''), or null when the OOXML
 *   cannot be parsed — callers fall back to the Word.js text property then
 */
export function extractFinalTextFromOoxml(ooxml) {
    const root = documentPartRoot(ooxml);
    if (!root) return null;
    return visibleText(root);
}

/**
 * Extracts the visible text of every TOP-LEVEL body paragraph, in document
 * order. Direct w:p children of w:body only — table-cell, SDT, and textbox
 * paragraphs are nested and therefore excluded, matching what Word.js
 * body.paragraphs enumerates. Callers should map the result onto
 * paragraphs.items by index and only trust it when the counts agree.
 *
 * @param {string} ooxml - body.getOoxml() result
 * @returns {string[]|null} One final-text string per top-level paragraph, or
 *   null when the OOXML cannot be parsed
 */
export function extractTopLevelParagraphTexts(ooxml) {
    const root = documentPartRoot(ooxml);
    if (!root) return null;
    const body = bodyOrRoot(root);
    /** @type {Element[]} */
    const paragraphs = [];
    for (const child of Array.from(body.childNodes || [])) {
        if (child.nodeType !== 1) continue;
        const element = /** @type {Element} */ (child);
        if (element.localName === 'p') paragraphs.push(element);
    }
    return paragraphs.map(visibleText);
}
