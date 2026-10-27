import { documentPartRoot } from './ooxml-text.js';

const records = [];
/** Retain only document XML and its relationships, never package binaries.
 * Included only when the user explicitly copies the complete context. */
export function recordNativeTextMismatch(read, state) {
    const root = documentPartRoot(read.value);
    if (!root) return;
    const wordNamespace = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    const content = root.getElementsByTagNameNS(wordNamespace, 'body')[0]
        || (root.namespaceURI === wordNamespace ? root : null);
    if (!content) return;
    const xml = new XMLSerializer();
    const parts = Array.from(root.ownerDocument.getElementsByTagNameNS('http://schemas.microsoft.com/office/2006/xmlPackage', 'part'));
    const relationships = parts.find((p) => p.getAttributeNS('http://schemas.microsoft.com/office/2006/xmlPackage', 'name') === '/word/_rels/document.xml.rels');
    records.push({ capturedAt: new Date().toISOString(), expectedText: state.text,
        reviewedText: read.currentText?.value, documentXml: xml.serializeToString(content).slice(0, 128000),
        relationshipsXml: relationships ? xml.serializeToString(relationships).slice(0, 32000) : null });
    while (records.length > 8) records.shift();
}

export function getNativeTextDiagnostics() { return records.map((record) => ({ ...record })); }
