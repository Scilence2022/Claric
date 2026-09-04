/**
 * Specs for src/lib/ooxml-text.js — revision-aware ("accept-all") text
 * extraction from OOXML fragments. The canonical bug case: Word.js
 * paragraph.text renders a paragraph whose words were replaced under Track
 * Changes as interleaved old/new word-salad ("TheResults reassemblerare...").
 */

const { JSDOM } = require('jsdom');
const { extractFinalTextFromOoxml, extractTopLevelParagraphTexts } = require('../src/lib/ooxml-text.js');

// Provide DOMParser (node test environment lacks it; the add-in WebView has it)
if (typeof globalThis.DOMParser === 'undefined') {
    const dom = new JSDOM('');
    globalThis.DOMParser = dom.window.DOMParser;
}

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

/** Wraps body inner XML in the pkg:package envelope getOoxml() returns. */
function wrapPackage(inner, { extraParts = '' } = {}) {
    return '<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">'
        + '<pkg:part pkg:name="/word/document.xml" pkg:contentType="text/xml"><pkg:xmlData>'
        + `<w:document xmlns:w="${W}"><w:body>${inner}</w:body></w:document>`
        + '</pkg:xmlData></pkg:part>'
        + extraParts
        + '</pkg:package>';
}

describe('extractFinalTextFromOoxml', () => {
    test('hides tracked deletions, keeps insertions and plain runs', () => {
        const ooxml = wrapPackage(
            '<w:p>'
            + '<w:del><w:r><w:delText xml:space="preserve">The </w:delText></w:r></w:del>'
            + '<w:ins><w:r><w:t xml:space="preserve">Results </w:t></w:r></w:ins>'
            + '<w:del><w:r><w:delText xml:space="preserve">reassembler </w:delText></w:r></w:del>'
            + '<w:ins><w:r><w:t xml:space="preserve">are </w:t></w:r></w:ins>'
            + '<w:del><w:r><w:delText xml:space="preserve">maps results </w:delText></w:r></w:del>'
            + '<w:ins><w:r><w:t xml:space="preserve">mapped </w:t></w:r></w:ins>'
            + '<w:r><w:t xml:space="preserve">back onto the document.</w:t></w:r>'
            + '</w:p>'
        );
        expect(extractFinalTextFromOoxml(ooxml)).toBe('Results are mapped back onto the document.');
    });

    test('hides moveFrom, keeps moveTo', () => {
        const ooxml = wrapPackage(
            '<w:p>'
            + '<w:moveFrom><w:r><w:delText>old home </w:delText></w:r></w:moveFrom>'
            + '<w:r><w:t xml:space="preserve">middle </w:t></w:r>'
            + '<w:moveTo><w:r><w:t>new home</w:t></w:r></w:moveTo>'
            + '</w:p>'
        );
        expect(extractFinalTextFromOoxml(ooxml)).toBe('middle new home');
    });

    test('handles tabs, breaks, non-breaking hyphens; skips field instructions', () => {
        const ooxml = wrapPackage(
            '<w:p>'
            + '<w:r><w:t xml:space="preserve">a</w:t><w:tab/><w:t>b</w:t><w:br/><w:t>c</w:t></w:r>'
            + '<w:r><w:instrText> TOC \\o "1-3" </w:instrText></w:r>'
            + '<w:r><w:t>d</w:t><w:noBreakHyphen/><w:t>e</w:t></w:r>'
            + '</w:p>'
        );
        expect(extractFinalTextFromOoxml(ooxml)).toBe('a\tb\ncd‑e');
    });

    test('walks only the /word/document.xml part of a package', () => {
        const commentsPart =
            '<pkg:part pkg:name="/word/comments.xml" pkg:contentType="text/xml"><pkg:xmlData>'
            + `<w:comments xmlns:w="${W}"><w:comment><w:p><w:r><w:t>a reviewer note</w:t></w:r></w:p></w:comment></w:comments>`
            + '</pkg:xmlData></pkg:part>';
        const ooxml = wrapPackage('<w:p><w:r><w:t>visible body text</w:t></w:r></w:p>', { extraParts: commentsPart });
        expect(extractFinalTextFromOoxml(ooxml)).toBe('visible body text');
    });

    test('accepts a bare range fragment without the pkg wrapper', () => {
        const ooxml = `<w:p xmlns:w="${W}"><w:del><w:r><w:delText>gone</w:delText></w:r></w:del>`
            + '<w:r><w:t>kept</w:t></w:r></w:p>';
        expect(extractFinalTextFromOoxml(ooxml)).toBe('kept');
    });

    test('returns null for missing or unparseable input', () => {
        expect(extractFinalTextFromOoxml(null)).toBeNull();
        expect(extractFinalTextFromOoxml('')).toBeNull();
        expect(extractFinalTextFromOoxml(123)).toBeNull();
        expect(extractFinalTextFromOoxml('<w:p><w:r>')).toBeNull();
    });

    test('a picture-only paragraph resolves to empty text', () => {
        const ooxml = wrapPackage('<w:p><w:r><w:drawing/></w:r></w:p>');
        expect(extractFinalTextFromOoxml(ooxml)).toBe('');
    });
});

describe('extractTopLevelParagraphTexts', () => {
    test('returns one accept-all text per top-level paragraph, in order', () => {
        const ooxml = wrapPackage(
            '<w:p><w:r><w:t>first</w:t></w:r></w:p>'
            + '<w:p><w:del><w:r><w:delText>old </w:delText></w:r></w:del><w:ins><w:r><w:t>new</w:t></w:r></w:ins></w:p>'
            + '<w:tbl><w:tr><w:tc><w:p><w:r><w:t>table cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
            + '<w:p><w:r><w:t>third</w:t></w:r></w:p>'
        );
        // Table-cell paragraphs are nested — excluded, matching Word.js body.paragraphs.
        expect(extractTopLevelParagraphTexts(ooxml)).toEqual(['first', 'new', 'third']);
    });

    test('accepts a bare w:body document without the pkg wrapper', () => {
        const ooxml = `<w:document xmlns:w="${W}"><w:body>`
            + '<w:p><w:r><w:t>alpha</w:t></w:r></w:p>'
            + '<w:p/>'
            + '<w:p><w:r><w:t>omega</w:t></w:r></w:p>'
            + '</w:body></w:document>';
        expect(extractTopLevelParagraphTexts(ooxml)).toEqual(['alpha', '', 'omega']);
    });

    test('returns null for unparseable input', () => {
        expect(extractTopLevelParagraphTexts(null)).toBeNull();
        expect(extractTopLevelParagraphTexts('<w:body>')).toBeNull();
    });
});
