/** @jest-environment jsdom */
const { requestsEmptyParagraphCleanup, isDeletableEmptyParagraphXml, inspectEmptyParagraphXml } = require('../src/lib/empty-paragraphs.js');
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const xml = (inner) => `<w:p xmlns:w="${W}">${inner}</w:p>`;

test.each([
    '调整选择部分的格式，包括多余的空行，不正确的粗体格式等',
    '清理选区空白行并恢复正文格式', 'delete all empty paragraphs',
    'Fix excess blank lines and incorrect bold', 'Remove whitespace paragraphs',
])('recognizes explicit cleanup: %s', (instruction) => {
    expect(requestsEmptyParagraphCleanup(instruction)).toBe(true);
});

test.each([undefined, null, 'Bold the heading', '段后间距设为6磅', '不要删除空行，只调整粗体',
    "Don't remove blank paragraphs", 'Do not delete empty lines', 'Never remove empty paragraphs',
    '调整格式，但保留空行', '修正粗体但不要动空行', '调整空行的字体大小', '处理空行的段后间距',
    'Fix bold formatting and keep blank lines', 'Fix blank paragraph spacing',
    'Remove extra bold but preserve empty paragraphs', 'Adjust formatting without removing empty lines',
    "Fix extra blank lines but don't change empty paragraphs"])('does not authorize deletion: %s', (instruction) => {
    expect(requestsEmptyParagraphCleanup(instruction)).toBe(false);
});

test('allows only verified whitespace and ordinary paragraph/font properties', () => {
    expect(isDeletableEmptyParagraphXml(xml(''))).toBe(true);
    expect(isDeletableEmptyParagraphXml(xml('<w:pPr><w:spacing w:after="120"/></w:pPr>'
        + '<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">  </w:t><w:tab/></w:r>'))).toBe(true);
    expect(isDeletableEmptyParagraphXml(`<w:document xmlns:w="${W}"><w:body><w:p/></w:body></w:document>`)).toBe(true);
});

test.each(['drawing', 'pict', 'object', 'fldChar', 'instrText', 'fldSimple', 'footnoteReference',
    'endnoteReference', 'commentReference', 'bookmarkStart', 'bookmarkEnd', 'commentRangeStart',
    'sdt', 'customXml', 'sectPr', 'numPr', 'sym', 'del', 'ins', 'moveFrom', 'moveTo',
    'pageBreakBefore'])('preserves blank-looking protected structure: %s', (tag) => {
    expect(isDeletableEmptyParagraphXml(xml(`<w:r><w:${tag}/></w:r>`))).toBe(false);
});

test('allows inserted whitespace and an inserted paragraph mark without accepting revisions', () => {
    expect(isDeletableEmptyParagraphXml(xml('<w:pPr><w:rPr><w:ins w:id="1" w:author="Example"/></w:rPr></w:pPr>'
        + '<w:ins w:id="2"><w:r><w:rPr><w:b/></w:rPr><w:t>  </w:t><w:tab/></w:r></w:ins>'))).toBe(true);
    expect(isDeletableEmptyParagraphXml(xml('<w:ins><w:r><w:t>Inserted prose</w:t></w:r></w:ins>'))).toBe(false);
    expect(isDeletableEmptyParagraphXml(xml('<w:ins><w:r><w:drawing/></w:r></w:ins>'))).toBe(false);
    expect(isDeletableEmptyParagraphXml(xml('<w:ins>Unexpected prose</w:ins>'))).toBe(false);
    expect(isDeletableEmptyParagraphXml(xml('<w:pPr><w:rPr><w:del w:id="1"/></w:rPr></w:pPr>'))).toBe(false);
    expect(inspectEmptyParagraphXml(xml('<w:pPr><w:rPr><w:del w:id="1"/></w:rPr></w:pPr>')).reason)
        .toBe('paragraph mark already tracked as deleted');
    expect(isDeletableEmptyParagraphXml(xml('<w:del><w:r><w:delText>Deleted prose</w:delText></w:r></w:del>'))).toBe(false);
});

test('distinguishes blank soft line breaks from pagination and wrapping commands', () => {
    expect(isDeletableEmptyParagraphXml(xml('<w:r><w:br/><w:br w:type="textWrapping"/><w:cr/></w:r>'))).toBe(true);
    for (const properties of ['w:type="page"', 'w:type="column"', 'w:clear="all"', 'w:clear="left"']) {
        expect(isDeletableEmptyParagraphXml(xml(`<w:r><w:br ${properties}/></w:r>`))).toBe(false);
    }
});

test('ordinary font effects and disabled pagination flags do not protect an empty paragraph', () => {
    for (const value of ['0', 'false', 'off']) {
        expect(isDeletableEmptyParagraphXml(xml(`<w:pPr><w:pageBreakBefore w:val="${value}"/></w:pPr>`
            + '<w:r><w:rPr><w:outline/><w:shadow/><w:emboss/><w:imprint/><w:spacing w:val="20"/></w:rPr><w:t> </w:t></w:r>'))).toBe(true);
    }
    expect(isDeletableEmptyParagraphXml(xml('<w:pPr><w:pageBreakBefore/></w:pPr>'))).toBe(false);
});

test('diagnostics expose structure without paragraph text, authors or bookmark names', () => {
    const evidence = inspectEmptyParagraphXml(xml('<w:bookmarkStart w:id="123" w:name="PrivateBookmark"/>'
        + '<w:del w:author="PrivateAuthor"><w:r><w:delText>Private prose</w:delText></w:r></w:del>'));
    expect(evidence).toEqual({ deletable: false, reason: 'protected or unsupported markup', element: 'w:bookmarkStart' });
    expect(JSON.stringify(evidence)).not.toMatch(/Private|123/);
    expect(inspectEmptyParagraphXml('<w:p>')).toEqual({ deletable: false, reason: 'invalid XML' });
    expect(inspectEmptyParagraphXml(`<w:body xmlns:w="${W}"><w:p/><w:p/></w:body>`).reason).toMatch(/one paragraph/);
});

test('ordinary formatting revision history does not make a whitespace paragraph undeletable', () => {
    expect(isDeletableEmptyParagraphXml(xml('<w:pPr><w:pPrChange w:id="1"><w:pPr><w:spacing w:after="120"/></w:pPr></w:pPrChange></w:pPr>'
        + '<w:r><w:rPr><w:rPrChange w:id="2"><w:rPr><w:b/></w:rPr></w:rPrChange></w:rPr><w:t> </w:t></w:r>'))).toBe(true);
    expect(isDeletableEmptyParagraphXml(xml('<w:pPr><w:pPrChange w:id="1"><w:pPr><w:sectPr/></w:pPr></w:pPrChange></w:pPr>'))).toBe(false);
});

test('rejects nonempty, multi-paragraph, malformed, unknown and nested protected XML', () => {
    expect(isDeletableEmptyParagraphXml(xml('<w:r><w:t>Body text</w:t></w:r>'))).toBe(false);
    expect(isDeletableEmptyParagraphXml(`<w:body xmlns:w="${W}"><w:p/><w:p/></w:body>`)).toBe(false);
    expect(isDeletableEmptyParagraphXml('<w:p>')).toBe(false);
    expect(isDeletableEmptyParagraphXml('')).toBe(false);
    expect(isDeletableEmptyParagraphXml('<root/>')).toBe(false);
    expect(isDeletableEmptyParagraphXml(xml('<x:drawing xmlns:x="other"/>'))).toBe(false);
    expect(isDeletableEmptyParagraphXml(`<w:sdt xmlns:w="${W}"><w:sdtContent><w:p/></w:sdtContent></w:sdt>`)).toBe(false);
    expect(isDeletableEmptyParagraphXml(`<w:tbl xmlns:w="${W}"><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>`)).toBe(false);
});
