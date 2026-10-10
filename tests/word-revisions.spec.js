/** @jest-environment jsdom */
import { revisionTextState, queueRevisionRead, resolveRevisionRead, normalizeRevisionText,
    verifyMutationText, MutationSafetyError } from '../src/lib/word-revisions.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const wrap = (xml) => `<w:document xmlns:w="${W}"><w:body>${xml}</w:body></w:document>`;

test('current text ignores deletions and move sources while islands retain physical revision boundaries', () => {
    const state = revisionTextState(wrap('<w:p><w:r><w:t>A</w:t></w:r><w:del><w:r><w:delText>old</w:delText></w:r></w:del>'
        + '<w:ins><w:r><w:t>B</w:t></w:r></w:ins><w:r><w:t>C</w:t></w:r></w:p>'));
    expect(state).toMatchObject({ text: 'ABC', hasRevisions: true, hasHiddenContent: true, protectedStructure: false,
        segments: [{ start: 0, text: 'A' }, { start: 1, text: 'B' }, { start: 2, text: 'C' }] });
    expect(revisionTextState(wrap('<w:p><w:moveFrom><w:r><w:t>old</w:t></w:r></w:moveFrom>'
        + '<w:moveTo><w:r><w:t>moved</w:t></w:r></w:moveTo></w:p>'))).toMatchObject({ text: 'moved', protectedStructure: true });
});

test('paragraph breaks and run formatting history cannot leak into the next round', () => {
    const state = revisionTextState(wrap('<w:p><w:r><w:rPr><w:rPrChange><w:rPr><w:t>historical</w:t></w:rPr></w:rPrChange></w:rPr>'
        + '<w:t>A</w:t><w:tab/><w:t>B</w:t><w:br/><w:cr/><w:noBreakHyphen/><w:instrText>FIELD</w:instrText></w:r></w:p>'
        + '<w:p><w:r><w:t>C</w:t></w:r></w:p>'));
    expect(state.text).toBe('A\tB\n\n‑\nC');
    expect(state.hasRevisions).toBe(true);
    expect(state.paragraphCount).toBe(2);
});

test('inserted paragraph marks allow another in-place edit; deleted marks remain protected', () => {
    expect(revisionTextState(wrap('<w:p><w:pPr><w:rPr><w:ins w:id="1"/></w:rPr></w:pPr>'
        + '<w:ins><w:r><w:t>New paragraph</w:t></w:r></w:ins></w:p>'))).toMatchObject({
        text: 'New paragraph', hasRevisions: true, protectedStructure: false,
    });
    expect(revisionTextState(wrap('<w:p><w:pPr><w:rPr><w:del/></w:rPr></w:pPr></w:p>')).protectedStructure).toBe(true);
    expect(revisionTextState(wrap('<w:p><w:r><w:delText>old</w:delText></w:r></w:p>'))).toMatchObject({
        text: '', hasRevisions: true, hasHiddenContent: true,
    });
});

test.each(['<w:pPr><w:rPr><w:del/></w:rPr></w:pPr>', '<w:pPr><w:pPrChange/></w:pPr>',
    '<w:drawing/>', '<w:fldSimple/>', '<w:sdt/>', '<w:oMath/>'])('structural changes and protected objects stay read-only: %s', (content) => {
    expect(revisionTextState(wrap(`<w:p>${content}<w:ins><w:r><w:t>Text</w:t></w:r></w:ins></w:p>`)).protectedStructure).toBe(true);
});

test('package side parts never enter the text or revision state', () => {
    const xml = '<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">'
        + `<pkg:part pkg:name="/word/document.xml"><pkg:xmlData>${wrap('<w:p><w:r><w:t>Body</w:t></w:r></w:p>')}</pkg:xmlData></pkg:part>`
        + `<pkg:part pkg:name="/word/comments.xml"><pkg:xmlData><w:comments xmlns:w="${W}"><w:ins><w:t>Comment</w:t></w:ins></w:comments></pkg:xmlData></pkg:part></pkg:package>`;
    expect(revisionTextState(xml)).toMatchObject({ text: 'Body', hasRevisions: false });
});

test('namespace aliases, empty text, missing APIs and malformed XML are handled explicitly', () => {
    expect(revisionTextState(`<x:p xmlns:x="${W}"><x:r><x:t>A</x:t><x:t>B</x:t></x:r></x:p>`).text).toBe('AB');
    expect(revisionTextState(wrap('<w:p/>')).text).toBe('');
    expect(() => revisionTextState('<w:p>')).toThrow(/unreadable/);
    expect(() => revisionTextState('<baseline/>')).toThrow(/no document content/);
    expect(queueRevisionRead({})).toBeNull();
    expect(resolveRevisionRead({ text: 'Plain' }, null).text).toBe('Plain');
    expect(resolveRevisionRead({}, null).text).toBe('');
    expect(normalizeRevisionText('A\r\nB\r')).toBe('A\nB');
});

test('known paragraph scopes omit export padding without deleting real line breaks or blank paragraphs', () => {
    const xml = wrap('<w:p><w:r><w:t>Text</w:t><w:br/></w:r></w:p><w:p/>');
    expect(revisionTextState(xml).text).toBe('Text\n\n');
    expect(resolveRevisionRead({}, { value: xml }, { paragraph: true }).text).toBe('Text\n');
    expect(revisionTextState(xml, { paragraph: true }).paragraphCount).toBe(1);
    const protectedEnd = wrap('<w:p><w:r><w:t>Text</w:t></w:r></w:p><w:p><w:pPr><w:sectPr/></w:pPr></w:p>');
    expect(revisionTextState(protectedEnd, { paragraph: true }).text).toBe('Text\n');
    const revisedEnd = wrap('<w:p><w:r><w:t>Text</w:t></w:r></w:p><w:p><w:ins/></w:p>');
    expect(revisionTextState(revisedEnd, { paragraph: true }).text).toBe('Text\n');
    expect(revisionTextState(wrap('<w:p><w:r><w:t>First</w:t></w:r></w:p><w:p/><w:p/>'), { paragraph: true }).text).toBe('First\n\n');
});

test('paragraph export padding does not multiply into model paragraph separators', () => {
    const texts = Array.from({ length: 36 }, (_, i) => `Generic paragraph ${i}.`);
    const reads = texts.map((text) => ({ value: wrap(`<w:p><w:r><w:t>${text}</w:t></w:r></w:p><w:p/>`) }));
    const actual = reads.map((read) => resolveRevisionRead({}, read, { paragraph: true }).text);
    expect(actual).toEqual(texts);
    expect(actual.join('\n')).toBe(texts.join('\n'));
    expect(revisionTextState(reads[0].value, { paragraph: true }).fingerprint)
        .toBe(revisionTextState(wrap(`<w:p><w:r><w:t>${texts[0]}</w:t></w:r></w:p>`)).fingerprint);
});

const revisedPadding = '<w:p><w:pPr><w:rPr><w:ins w:id="1"/></w:rPr></w:pPr></w:p>';
test('revised export padding requires independent current-text proof before it can be omitted', () => {
    const first = '<w:p><w:r><w:t>Text</w:t><w:br/></w:r></w:p>';
    const xml = wrap(first + revisedPadding);
    const range = { getOoxml: () => ({ value: xml }), getReviewedText: () => ({ value: 'Text\r' }) };
    const state = resolveRevisionRead(range, queueRevisionRead(range, { paragraph: true }), { paragraph: true });
    expect(state.text).toBe('Text\n');
    expect(state.paragraphCount).toBe(1);
    expect(state.fingerprint).toBe(revisionTextState(wrap(first), { paragraph: true }).fingerprint);
    expect(revisionTextState(xml, { paragraph: true }).text).toBe('Text\n\n');
    expect(revisionTextState(xml, { paragraph: true, currentText: 'Text\r\r' }).text).toBe('Text\n\n');
    expect(revisionTextState(xml, { paragraph: false, currentText: 'Text\r' }).text).toBe('Text\n\n');
});

test.each(['<w:del><w:r><w:delText>History</w:delText></w:r></w:del>', '<w:drawing/>', '<w:br/>', '<w:pPr><w:pPrChange/></w:pPr>'])('current-text proof never hides protected or meaningful trailing content: %s', (end) => {
    const xml = wrap('<w:p><w:r><w:t>Text</w:t></w:r></w:p>' + `<w:p>${end}</w:p>`);
    expect(revisionTextState(xml, { paragraph: true, currentText: 'Text' }).paragraphCount).toBe(2);
});

test('native current-text mismatches are rejected even when the XML export matches the proposal', async () => {
    const range = { getOoxml: jest.fn(() => ({ value: wrap('<w:p><w:r><w:t>Expected</w:t></w:r></w:p>') })),
        getReviewedText: jest.fn(() => ({ value: 'Unexpected' })) };
    const context = { sync: jest.fn(async () => {}) };
    await expect(verifyMutationText(context, range, 'Expected', { paragraph: true })).rejects.toBeInstanceOf(MutationSafetyError);
    await expect(verifyMutationText(context, range, 'Expected', { paragraph: true })).rejects.toThrow(/expected 8 chars, received 10; first mismatch at 0/);
    expect(range.getOoxml).not.toHaveBeenCalled();
});

test('an unreadable native current-text result cannot verify an empty proposal', async () => {
    const range = { getReviewedText: () => ({ value: undefined }) };
    await expect(verifyMutationText({ sync: async () => {} }, range, '', { paragraph: true })).rejects.toThrow(/unreadable/);
});

test('older hosts retain strict XML verification without calling unsupported current-text APIs', async () => {
    global.Office = { context: { requirements: { isSetSupported: () => false } } };
    try {
        const range = { getReviewedText: jest.fn(), getOoxml: () => ({ value: wrap('<w:p><w:r><w:t>Actual</w:t></w:r></w:p>') }) };
        await expect(verifyMutationText({ sync: async () => {} }, range, 'Expected', { paragraph: true })).rejects.toThrow(/OOXML/);
        expect(range.getReviewedText).not.toHaveBeenCalled();
    } finally { delete global.Office; }
});
