/** @jest-environment jsdom */
import { revisionTextState, queueRevisionRead, resolveRevisionRead, normalizeRevisionText,
    verifyMutationText, MutationSafetyError, hostSupportsTrackedChanges, projectNativeText } from '../src/lib/word-revisions.js';
import { parseDocument } from '../src/lib/document-parser.js';
import { checkChunkAmendment } from '../src/lib/amendment-guard.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const wrap = (xml) => `<w:document xmlns:w="${W}"><w:body>${xml}</w:body></w:document>`;

// Text-bearing TOC structure from report-subspecialty-training.docx. The
// original contains one generated TOC with 13 entries and no manual breaks.
const tefToc = require('fs').readFileSync(require('path').join(__dirname, 'fixtures/tef-generated-toc.xml'), 'utf8');
test('expanded TOC exports resolve each native paragraph once without inventing manual breaks', () => {
    const doc = new DOMParser().parseFromString(tefToc, 'text/xml');
    const entries = Array.from(doc.getElementsByTagNameNS(W, 'p'));
    const texts = entries.map((entry) => revisionTextState(new XMLSerializer().serializeToString(entry)).text);
    expect(texts).toHaveLength(13);
    const resolved = texts.map((text) => revisionTextState(wrap(tefToc), { paragraph: true, currentText: text }));
    expect(resolved.map((state) => state.text)).toEqual(texts);
    expect(resolved.every((state) => state.paragraphCount === 1 && !state.text.includes('\n'))).toBe(true);
    expect(resolved.every((state) => state.protectedSpans.some(({ start, end }) => start === 0 && end === state.text.length))).toBe(true);
    expect(resolved[1].text).toBe('Executive Summary\t3');
    expect(() => revisionTextState(wrap(tefToc), { paragraph: true, currentText: 'Unknown entry' })).toThrow(/isolate/);
    expect(revisionTextState(wrap(tefToc)).paragraphCount).toBe(13);
});

test('paragraph reads export and review the same Content range rather than the Paragraph proxy', () => {
    const content = { getOoxml: jest.fn(() => ({ value: wrap('<w:p><w:r><w:t>One entry</w:t></w:r></w:p>') })),
        getReviewedText: jest.fn(() => ({ value: 'One entry' })) };
    const paragraph = { getOoxml: jest.fn(() => ({ value: wrap(tefToc) })), getRange: jest.fn(() => content) };
    const read = queueRevisionRead(paragraph, { paragraph: true });
    expect(resolveRevisionRead(paragraph, read, { paragraph: true }).text).toBe('One entry');
    expect(paragraph.getOoxml).not.toHaveBeenCalled();
    expect(paragraph.getRange).toHaveBeenCalledWith('Content');
    expect(content.getReviewedText).toHaveBeenCalledTimes(1);
});

test('a mixed section with 25 paragraphs preserves one TOC and accepts surrounding prose edits', async () => {
    const entries = Array.from(new DOMParser().parseFromString(tefToc, 'text/xml').getElementsByTagNameNS(W, 'p'));
    const texts = ['TEF Thematic Report', ...entries.map((entry) => revisionTextState(new XMLSerializer().serializeToString(entry)).text),
        ...Array.from({ length: 11 }, (_, i) => `Report introduction paragraph ${i + 1}.`)];
    const items = texts.map((text, index) => {
        const xml = index > 0 && index <= 13 ? wrap(tefToc) : wrap(`<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`);
        return { text, style: 'Normal', styleBuiltIn: 'Normal', isListItem: false, load: jest.fn(), untrack: jest.fn(),
            parentTableOrNullObject: { isNullObject: true, load: jest.fn() }, getOoxml: () => ({ value: xml }),
            getRange: () => ({ getOoxml: () => ({ value: xml }), getReviewedText: () => ({ value: text }) }) };
    });
    global.Word = { run: async (callback) => callback({ document: { body: { paragraphs: { items, load: jest.fn() } } }, sync: async () => {} }) };
    try {
        const parsed = await parseDocument();
        expect(parsed.paragraphs.map((p) => p.text)).toEqual(texts);
        expect(parsed.paragraphs).toHaveLength(25);
        const output = texts.join('\n').replace('Report introduction paragraph 1.', 'Revised introduction paragraph 1.');
        expect(checkChunkAmendment({ paragraphs: parsed.paragraphs }, output).valid).toBe(true);
        const invalid = texts.join('\n').replace('Executive Summary\t3', 'Executive Summary\t4');
        expect(checkChunkAmendment({ paragraphs: parsed.paragraphs }, invalid).valid).toBe(false);
    } finally { delete global.Word; }
});

const refCode = ' REF _Ref533004983 \\h  \\* MERGEFORMAT ';
const refXml = '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
    + `<w:r><w:instrText>${refCode}</w:instrText><w:fldChar w:fldCharType="separate"/><w:t>Table 4</w:t>`
    + '<w:fldChar w:fldCharType="end"/></w:r>';

// Minimal text-bearing XML from the supplied NIST SOP 14 DOCX. Native streams
// are reconstructed using Word's controls and one boundary padding space in
// each REF instruction to exercise the logged +42 length difference. These
// are regression inputs, not captured getReviewedText responses.
test.each(require('./fixtures/nist-native-text.json'))('NIST SOP 14 chunk $chunk reconciles its reported native mismatch', (fixture) => {
    const state = revisionTextState(fixture.xml, { paragraph: true });
    const legacyText = state.text.replace('°C, for each run', 'C, for each run');
    expect(legacyText.length).toBe(fixture.legacyLength);
    expect(fixture.nativeText.length).toBe(fixture.nativeLength);
    let mismatch = 0;
    while (legacyText[mismatch] === fixture.nativeText[mismatch]) mismatch++;
    expect(mismatch).toBe(fixture.firstMismatch);
    expect(state.text).toBe(fixture.expectedText);
    const projection = projectNativeText(state, fixture.nativeText);
    expect(projection.text).toBe(fixture.expectedText);
    expect(projection.offsets[fixture.nativeText.length]).toBe(state.text.length);
});

test.each([false, true])('native REF text projects exact field codes to visible offsets, raw=%s', (raw) => {
    const state = revisionTextState(wrap('<w:p><w:r><w:t>See </w:t></w:r>' + refXml + '<w:r><w:t> after.</w:t></w:r></w:p>'));
    const field = raw ? '\u0013' + refCode + '\u0014Table 4\u0015' : 'Table 4';
    const projection = projectNativeText(state, 'See ' + field + ' after.');
    expect(projection.text).toBe('See Table 4 after.');
    expect(projection.offsets[4 + field.length + 1]).toBe('See Table 4 '.length);
    expect(state.protectedSpans).toEqual([{ start: 4, end: 11 }]);
    if (raw) {
        const padded = 'See ' + field.replace('\u0014', ' \u0014') + ' after.';
        expect(projectNativeText(state, padded).text).toBe(state.text);
        expect(projectNativeText(state, padded).offsets[padded.length]).toBe(state.text.length);
        expect(projectNativeText(state, 'See ' + field.replace('533004983', '533004984') + ' after.')).toBeNull();
        expect(projectNativeText(state, 'See ' + field.replace('Table 4', 'Table 5') + ' after.')).toBeNull();
        expect(projectNativeText(state, 'See ' + field.replace('\u0014', '') + ' after.')).toBeNull();
    }
});

test('native symbol placeholders require XML evidence and do not rewrite literal parentheses', () => {
    const state = revisionTextState(wrap('<w:p><w:r><w:t>(20 </w:t><w:sym w:font="Symbol" w:char="F0B0"/><w:t>C) </w:t>'
        + '<w:sym w:font="Symbol" w:char="F061"/><w:t> / </w:t><w:sym w:font="Symbol" w:char="F072"/></w:r></w:p>'));
    expect(state.text).toBe('(20 °C) α / ρ');
    expect(projectNativeText(state, '(20 (C) ( / (').text).toBe(state.text);
    expect(projectNativeText(state, '(20 °C) α / ρ').text).toBe(state.text);
    expect(projectNativeText(state, '(20 \uF0B0C) \uF061 / \uF072').text).toBe(state.text);
    expect(projectNativeText(state, '(20 ?C) ( / (')).toBeNull();
    expect(projectNativeText(state, '[20 (C) ( / (')).toBeNull();
    expect(state.protectedSpans).toEqual([{ start: 4, end: 5 }, { start: 8, end: 9 }, { start: 12, end: 13 }]);
});

test('nonbreaking hyphen controls are normalized without collapsing ordinary hyphens or spaces', () => {
    const state = revisionTextState(wrap('<w:p><w:r><w:t>For pour</w:t><w:noBreakHyphen/><w:t>type: 10 s.</w:t></w:r></w:p>'));
    expect(projectNativeText(state, 'For pour\u001Etype: 10 s.').text).toBe(state.text);
    expect(projectNativeText(state, 'For pour-type: 10 s.')).toBeNull();
    expect(projectNativeText(state, 'For pour\u001Etype: 10 s.')).toBeNull();
});

test('read-back reconciles raw fields using fresh XML and rejects changed visible results', async () => {
    const range = { getReviewedText: () => ({ value: 'See \u0013' + refCode + '\u0014Table 4\u0015 after.' }),
        getOoxml: () => ({ value: wrap('<w:p><w:r><w:t>See </w:t></w:r>' + refXml + '<w:r><w:t> after.</w:t></w:r></w:p>') }) };
    const context = { sync: async () => {} };
    await expect(verifyMutationText(context, range, 'See Table 4 after.', { paragraph: true })).resolves.toBeUndefined();
    await expect(verifyMutationText(context, range, 'See Table 5 after.', { paragraph: true })).rejects.toThrow(/differs/);
    range.getOoxml = () => ({ value: wrap('<w:p><w:r><w:t>See Table 4 after.</w:t></w:r></w:p>') });
    await expect(verifyMutationText(context, range, 'See Table 4 after.', { paragraph: true })).rejects.toThrow(/differs/);
});

test.each(['(', '\uF0B0', '°'])('symbol read-back validates fresh glyph XML as well as native text: %s', async (native) => {
    let slot = 'F0B0';
    const range = { getReviewedText: () => ({ value: '20 ' + native + 'C' }),
        getOoxml: () => ({ value: wrap(`<w:p><w:r><w:t>20 </w:t><w:sym w:font="Symbol" w:char="${slot}"/><w:t>C</w:t></w:r></w:p>`) }) };
    const context = { sync: async () => {} };
    await expect(verifyMutationText(context, range, '20 °C', { paragraph: true })).resolves.toBeUndefined();
    if (native !== '°') {
        // A different glyph has the same native '(' placeholder. Only the
        // post-write XML can distinguish it from the expected degree sign.
        slot = 'F061';
        await expect(verifyMutationText(context, range, '20 °C', { paragraph: true })).rejects.toThrow(/differs/);
    }
});

test('complex EndNote-style fields and nested links expose immutable visible spans without field codes', () => {
    const xml = wrap('<w:p><w:r><w:t>Before </w:t><w:fldChar w:fldCharType="begin"/></w:r>'
        + '<w:r><w:instrText>ADDIN EN.CITE private field code</w:instrText><w:fldChar w:fldCharType="separate"/></w:r>'
        + '<w:r><w:t>(Example 2025)</w:t><w:fldChar w:fldCharType="end"/><w:t> after.</w:t></w:r></w:p>');
    expect(revisionTextState(xml)).toMatchObject({ text: 'Before (Example 2025) after.',
        protectedStructure: true, fieldOnlyProtection: true, protectedSpans: [{ start: 7, end: 21 }] });
    const partial = wrap('<w:p><w:r><w:t>Continuing field result</w:t><w:fldChar w:fldCharType="end"/><w:t> prose</w:t></w:r></w:p>');
    expect(revisionTextState(partial).protectedSpans).toEqual([{ start: 0, end: 23 }]);
});

test('native manual and page break representations verify without ignoring meaningful character differences', async () => {
    const context = { sync: async () => {} };
    const range = { getReviewedText: () => ({ value: 'First\u000bSecond\u000cThird' }) };
    await expect(verifyMutationText(context, range, 'First\nSecond\nThird', { paragraph: true })).resolves.toBeUndefined();
    await expect(verifyMutationText(context, range, 'First Second\nThird', { paragraph: true })).rejects.toThrow(/U\+0020.*U\+000A/);
});

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

describe('hostSupportsTrackedChanges', () => {
    afterEach(() => {
        delete global.Word;
        delete global.Office;
    });

    test('requires the Word API and its tracking enum', () => {
        expect(hostSupportsTrackedChanges()).toBe(false);
        global.Word = {};
        expect(hostSupportsTrackedChanges()).toBe(false);
    });

    test('the enum alone is not host support when Office reports the requirement set', () => {
        global.Word = { ChangeTrackingMode: { trackAll: 'TrackAll', off: 'Off' } };
        const isSetSupported = jest.fn((name, version) => name === 'WordApi' && version === '1.3');
        global.Office = { context: { requirements: { isSetSupported } } };
        expect(hostSupportsTrackedChanges()).toBe(false);
        expect(isSetSupported).toHaveBeenCalledWith('WordApi', '1.4');
        isSetSupported.mockReturnValue(true);
        expect(hostSupportsTrackedChanges()).toBe(true);
    });

    test('without Office requirement data the enum decides', () => {
        global.Word = { ChangeTrackingMode: { trackAll: 'TrackAll', off: 'Off' } };
        expect(hostSupportsTrackedChanges()).toBe(true);
    });
});
