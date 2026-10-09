/** @jest-environment jsdom */
const { rangeStructureFingerprint: fingerprint, paragraphStructureFingerprint,
    rangeFingerprintDifference } = require('../src/lib/ooxml-fingerprint.js');
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const xml = (body) => `<w:document xmlns:w="${W}"><w:body>${body}</w:body></w:document>`;
const run = (text, properties = '') => `<w:r>${properties ? `<w:rPr>${properties}</w:rPr>` : ''}<w:t xml:space="preserve">${text}</w:t></w:r>`;
const paragraph = (content, properties = '') => `<w:p>${properties ? `<w:pPr>${properties}</w:pPr>` : ''}${content}</w:p>`;

test('bookmark/proofing run splits, equivalent booleans and property order do not change formatting', () => {
    const original = xml(paragraph(run('Example heading', '<w:b/><w:sz w:val="24"/>')));
    const exported = xml(paragraph(run('Example ', '<w:sz w:val="24"/><w:b w:val="1"/>')
        + '<w:bookmarkStart w:id="10" w:name="_claric_fmt_example"/><w:proofErr w:type="spellStart"/>'
        + run('heading', '<w:b w:val="true"/><w:sz w:val="24"/>')
        + '<w:bookmarkEnd w:id="10"/><w:r><w:rPr><w:i/></w:rPr></w:r>'));
    expect(fingerprint(exported)).toBe(fingerprint(original));
    // Prose replacement still uses the stricter comparator.
    expect(paragraphStructureFingerprint(exported)).not.toBe(paragraphStructureFingerprint(original));
});

test('range package/fragment containers and synthetic trailing section layout compare equally', () => {
    const p = paragraph(run('Example body'));
    const bare = `<w:p xmlns:w="${W}">${run('Example body')}</w:p>`;
    const packaged = '<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">'
        + `<pkg:part pkg:name="/word/document.xml"><pkg:xmlData>${xml(p + '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>')}</pkg:xmlData></pkg:part>`
        + '<pkg:part pkg:name="/word/settings.xml"><pkg:xmlData><settings/></pkg:xmlData></pkg:part></pkg:package>';
    expect(fingerprint(packaged)).toBe(fingerprint(bare));
});

test('empty property containers and adjacent text elements are serialization details', () => {
    expect(fingerprint(xml('<w:p><w:pPr/><w:r><w:rPr/><w:t>Example </w:t><w:t>body</w:t></w:r></w:p>')))
        .toBe(fingerprint(xml(paragraph(run('Example body')))));
});

test.each([
    ['bold', paragraph(run('Example body', '<w:b/>'))],
    ['italic', paragraph(run('Example body', '<w:i/>'))],
    ['font', paragraph(run('Example body', '<w:rFonts w:ascii="Arial"/>'))],
    ['size', paragraph(run('Example body', '<w:sz w:val="28"/>'))],
    ['spacing', paragraph(run('Example body'), '<w:spacing w:after="240"/>')],
    ['style', paragraph(run('Example body'), '<w:pStyle w:val="Heading1"/>')],
    ['section break', paragraph(run('Example body'), '<w:sectPr><w:type w:val="nextPage"/></w:sectPr>')],
    ['blank paragraph', paragraph(run('Example body')) + '<w:p/>'],
    ['field', paragraph(run('Example body') + '<w:r><w:fldChar w:fldCharType="begin"/></w:r>')],
    ['drawing', paragraph(run('Example body') + '<w:r><w:drawing/></w:r>')],
    ['tracked insertion', paragraph('<w:ins w:id="2" w:author="Editor">' + run('Example body') + '</w:ins>')],
    ['tracked deletion', paragraph(run('Example body') + '<w:del w:id="3"><w:r><w:delText>Old text</w:delText></w:r></w:del>')],
    ['whitespace', paragraph(run('Example  body'))],
])('actual %s changes still invalidate formatting before writes', (_name, changed) => {
    expect(fingerprint(xml(changed))).not.toBe(fingerprint(xml(paragraph(run('Example body')))));
});

test('moving bold between characters is detected even with identical text and mixed overall font', () => {
    expect(fingerprint(xml(paragraph(run('Example ', '<w:b/>') + run('body')))))
        .not.toBe(fingerprint(xml(paragraph(run('Example ') + run('body', '<w:b/>')))));
});

test('conflicting repeated property order and inherited vs explicitly disabled bold remain distinct', () => {
    expect(fingerprint(xml(paragraph(run('Example body', '<w:b w:val="0"/><w:b/>')))))
        .not.toBe(fingerprint(xml(paragraph(run('Example body', '<w:b/><w:b w:val="0"/>')))));
    expect(fingerprint(xml(paragraph(run('Example body')))))
        .not.toBe(fingerprint(xml(paragraph(run('Example body', '<w:b w:val="false"/>')))));
});

test('empty runs with fields or tracked formatting history are never discarded', () => {
    const original = paragraph(run('Example body'));
    for (const extra of ['<w:r><w:fldChar w:fldCharType="end"/></w:r>',
        '<w:r><w:rPr><w:rPrChange w:id="4"><w:rPr><w:b/></w:rPr></w:rPrChange></w:rPr></w:r>']) {
        expect(fingerprint(xml(original + extra))).not.toBe(fingerprint(xml(original)));
    }
});

test('unknown markup in text containers is preserved rather than merged into literal text', () => {
    const fragment = (attribute) => xml('<w:p><w:r><w:t>Example </w:t>'
        + `<w:t><custom xmlns="urn:custom" value="${attribute}"/></w:t></w:r></w:p>`);
    expect(fingerprint(fragment('before'))).not.toBe(fingerprint(fragment('after')));
});

test('diagnostics describe the changed structure without exposing document text', () => {
    const before = fingerprint(xml(paragraph(run('Private example text'))));
    const after = fingerprint(xml(paragraph(run('Changed private example'))));
    const detail = rangeFingerprintDifference(before, after);
    expect(detail).toContain('scope');
    expect(detail).not.toMatch(/Private|private|Changed|example/);
    expect(rangeFingerprintDifference(before, before)).toBe('');
    expect(rangeFingerprintDifference(before, null)).toBe('unreadable range structure');
});

test('drawing mismatch diagnostics name changed attributes without revealing their values', () => {
    const fragment = (attr) => xml('<w:p><w:r><w:drawing><wp:inline '
        + 'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
        + 'xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" '
        + `${attr}/></w:drawing></w:r></w:p>`);
    const detail = rangeFingerprintDifference(fingerprint(fragment('wp14:anchorId="private-old" wp:distT="private-remove"')),
        fingerprint(fragment('wp14:anchorId="private-new" wp:distB="private-add"')));
    expect(detail).toContain('wp14:anchorId changed');
    expect(detail).toContain('wp:distT removed');
    expect(detail).toContain('wp:distB added');
    expect(detail).not.toContain('private');
});
