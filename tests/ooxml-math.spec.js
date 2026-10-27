/** @jest-environment jsdom */
import { MATH_NS, mathTextState } from '../src/lib/ooxml-math.js';
import { revisionTextState, projectNativeText, verifyMutationText } from '../src/lib/word-revisions.js';
import { extractFinalTextFromOoxml } from '../src/lib/ooxml-text.js';
import { checkChunkAmendment } from '../src/lib/amendment-guard.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const run = (text, properties = '') => `<m:r>${properties ? `<m:rPr>${properties}</m:rPr>` : ''}<m:t>${text}</m:t></m:r>`;
const equation = (xml) => `<m:oMath xmlns:m="${MATH_NS}">${xml}</m:oMath>`;
const power = `<m:sSup><m:e>${run('x')}</m:e><m:sup>${run('2')}</m:sup></m:sSup>`;
const paragraph = (math = power, before = 'An old ', after = ' result.') => `<w:p xmlns:w="${W}"><w:r><w:t>${before}</w:t></w:r>`
    + equation(math) + `<w:r><w:t>${after}</w:t></w:r></w:p>`;
const readMath = (xml) => mathTextState(new DOMParser().parseFromString(equation(xml), 'text/xml').documentElement);

test('native math argument CRs project to one protected inline equation rather than new paragraphs', () => {
    const state = revisionTextState(paragraph());
    expect(state.text).toBe('An old 𝑥^(2) result.');
    expect(state.paragraphCount).toBe(1);
    expect(state.protectedSpans).toEqual([{ start: 7, end: 13 }]);
    expect(state.segments).toEqual([{ start: 0, text: 'An old ' }, { start: 13, text: ' result.' }]);
    const native = 'An old \r𝑥\r2\r result.';
    const projection = projectNativeText(state, native);
    expect(projection.text).toBe(state.text);
    expect(projection.offsets[native.indexOf(' result.')]).toBe(13);
    expect(projection.offsets[native.length]).toBe(state.text.length);
    expect(extractFinalTextFromOoxml(paragraph())).toBe(state.text);
    expect(projectNativeText(state, native.replace('2', '3'))).toBeNull();
    expect(projectNativeText(state, native.replace(' result.', ' result!'))).toBeNull();
    expect(projectNativeText(state, native.replace('old ', 'old '))).toBeNull();
    expect(projectNativeText(state, native.replace('\r𝑥\r', '𝑥\r'))).toBeNull();
});

test('equation preservation is enforced before a chunk amendment reaches Word', () => {
    const state = revisionTextState(paragraph());
    const chunk = { paragraphs: [state] };
    expect(checkChunkAmendment(chunk, 'A revised 𝑥^(2) result.').valid).toBe(true);
    expect(checkChunkAmendment(chunk, 'A revised 𝑥^(3) result.').valid).toBe(false);
    expect(checkChunkAmendment(chunk, 'A revised result.').valid).toBe(false);
});

test('Mac reviewed structured math maps exact prose positions without normalizing surrounding punctuation', () => {
    const sub = `<m:sSub><m:e>${run('y')}</m:e><m:sub>${run('j')}</m:sub></m:sSub>`;
    const nary = '<m:nary><m:naryPr><m:chr m:val="∑"/></m:naryPr>'
        + `<m:sub>${run('j≥1')}</m:sub><m:sup>${run('​')}</m:sup><m:e>${power}</m:e></m:nary>`;
    const xml = paragraph(sub + run("-'") + nary, 'A plain-text - before ', ' after the sum.');
    const state = revisionTextState(xml);
    const native = 'A plain-text - before ,𝑦-𝑗.−′,𝑗≥1-​-,𝑥-2.. after the sum.';
    const projection = projectNativeText(state, native);
    expect(projection.text).toBe(state.text);
    expect(projection.offsets[native.indexOf(' after')]).toBe(state.text.indexOf(' after'));
    expect(projection.offsets[native.length]).toBe(state.text.length);
    for (const bad of [native.replace('plain-text', 'plain−text'), native.replace('𝑗.', '𝑙.'),
        native.replace(',𝑥-2.', ',𝑥-3.'), native.replace(' after', ' after'), native.replace('−′', '-\'')]) {
        expect(projectNativeText(state, bad)).toBeNull();
    }
});

test('reviewed math does not infer punctuation for uncaptured constructs', () => {
    const node = new DOMParser().parseFromString(equation(`<m:unknown><m:e>${run('x')}</m:e></m:unknown>`), 'text/xml').documentElement;
    expect(mathTextState(node, true).nativeText).toBeNull();
    expect(mathTextState(new DOMParser().parseFromString('<other/>', 'text/xml').documentElement, true).nativeText).toBeNull();
});

test.each([
    ['kΣh', '', '𝑘𝛴ℎ'], ['Alpha', '<m:nor/>', 'Alpha'],
    ['ABCxyz', '<m:sty m:val="p"/>', 'ABCxyz'], ['F', '<m:scr m:val="script"/><m:sty m:val="p"/>', 'ℱ'],
    ['R', '<m:scr m:val="double-struck"/>', 'ℝ'], ['A2', '<m:sty m:val="b"/>', '𝐀𝟐'],
    ['A', '<m:scr m:val="sans-serif"/><m:sty m:val="i"/>', '𝘈'],
    ['*', '<m:sty m:val="p"/>', '∗'],
])('math font properties retain native glyphs for %s', (text, properties, expected) => {
    expect(readMath(run(text, properties))).toEqual({ text: expected, nativeText: expected });
});

test('nested fractions, delimiters, limits and n-ary arguments have exact native boundaries', () => {
    const fraction = `<m:f><m:num>${power}</m:num><m:den>${run('y')}</m:den></m:f>`;
    expect(readMath(fraction)).toEqual({ text: '(𝑥^(2))/(𝑦)', nativeText: '\n\n𝑥\n2\n\n𝑦\n' });
    const binomial = `<m:f><m:fPr><m:type m:val="noBar"/></m:fPr><m:num>${run('n')}</m:num><m:den>${run('j')}</m:den></m:f>`;
    expect(readMath(binomial)).toEqual({ text: 'binom(𝑛,𝑗)', nativeText: '\n\n𝑛\n𝑗\n' });
    const delimiter = `<m:d><m:dPr><m:begChr m:val="["/><m:endChr m:val="]"/></m:dPr><m:e>${fraction}</m:e></m:d>`;
    expect(readMath(delimiter)).toEqual({ text: '[(𝑥^(2))/(𝑦)]', nativeText: '\n\n\n𝑥\n2\n\n𝑦\n\n' });
    const nary = `<m:nary><m:naryPr><m:chr m:val="∑"/></m:naryPr><m:sub>${run('j')}</m:sub><m:sup>${run('n')}</m:sup><m:e>${power}</m:e></m:nary>`;
    expect(readMath(nary)).toEqual({ text: '∑_(𝑗)^(𝑛)(𝑥^(2))', nativeText: '\n𝑗\n𝑛\n𝑥\n2\n\n' });
    const limit = `<m:limLow><m:e>${run('lim', '<m:nor/>')}</m:e><m:lim>${run('n')}</m:lim></m:limLow>`;
    expect(readMath(limit)).toEqual({ text: 'lim_(𝑛)', nativeText: '\nlim\n𝑛\n' });
    const both = `<m:sSubSup><m:e>${run('x')}</m:e><m:sub>${run('j')}</m:sub><m:sup>${run('2')}</m:sup></m:sSubSup>`;
    expect(readMath(both)).toEqual({ text: '𝑥_(𝑗)^(2)', nativeText: '\n𝑥\n𝑗\n2\n' });
});

test('unknown math structures stay protected without guessing native argument boundaries', () => {
    const state = revisionTextState(paragraph(`<m:unknown><m:e>${run('x')}</m:e></m:unknown>`));
    expect(state).toMatchObject({ text: 'An old 𝑥 result.', protectedStructure: true, fieldOnlyProtection: false });
    expect(projectNativeText(state, 'An old \r𝑥\r result.')).toBeNull();
});

test('read-back validates structured equations against fresh XML', async () => {
    let exponent = '2';
    const range = { getReviewedText: () => ({ value: `An old \r𝑥\r${exponent}\r result.` }),
        getOoxml: jest.fn(() => ({ value: paragraph(power.replace('<m:t>2', `<m:t>${exponent}`)) })) };
    const context = { sync: async () => {} };
    await expect(verifyMutationText(context, range, 'An old 𝑥^(2) result.', { paragraph: true })).resolves.toBeUndefined();
    expect(range.getOoxml).toHaveBeenCalled();
    exponent = '3';
    await expect(verifyMutationText(context, range, 'An old 𝑥^(2) result.', { paragraph: true })).rejects.toThrow(/read-back differs/);
});

const scriptL = run('l', '<m:scr m:val="script"/><m:sty m:val="p"/>');

test('Mac reviewed script l projects only inside the evidenced equation, with exact following prose offsets', () => {
    const xml = paragraph(scriptL, 'A literal 𝓁 and ', '-unit limit.');
    const state = revisionTextState(xml);
    expect(state.text).toBe('A literal 𝓁 and 𝓁-unit limit.');
    const native = 'A literal 𝓁 and ℓ-unit limit.';
    const projection = projectNativeText(state, native);
    expect(projection?.text).toBe(state.text);
    expect(projection?.offsets[native.indexOf('-unit')]).toBe(state.text.indexOf('-unit'));
    expect(projection?.offsets[native.length]).toBe(state.text.length);
    expect(projectNativeText(state, state.text)?.text).toBe(state.text);
    for (const bad of [native.replace('literal 𝓁', 'literal ℓ'), native.replace('ℓ-unit', 'l-unit'),
        native.replace('limit.', 'limits.'), native.replace(' and ', ' and ')]) {
        expect(projectNativeText(state, bad)).toBeNull();
    }
    const nested = revisionTextState(paragraph(`<m:sSup><m:e>${run('2')}</m:e><m:sup>${scriptL}</m:sup></m:sSup>`));
    expect(projectNativeText(nested, 'An old ,2-ℓ. result.')?.text).toBe('An old 2^(𝓁) result.');
});

test.each([
    run('l'), run('l', '<m:scr m:val="script"/><m:sty m:val="b"/>'),
    run('𝓁', '<m:nor/>'), run('𝓁', '<m:sty m:val="p"/>'),
])('script l aliases do not apply to other math styles or literal glyphs: %s', (math) => {
    expect(projectNativeText(revisionTextState(paragraph(math)), 'An old ℓ result.')).toBeNull();
});

test('script l read-back fetches fresh XML even when the native text contains no supplementary math glyph', async () => {
    let xml = paragraph(scriptL, 'A revised ', '-unit limit.');
    let native = 'A revised ℓ-unit limit.';
    const range = { getReviewedText: () => ({ value: native }), getOoxml: jest.fn(() => ({ value: xml })) };
    const context = { sync: async () => {} };
    const expected = 'A revised 𝓁-unit limit.';
    await expect(verifyMutationText(context, range, expected, { paragraph: true })).resolves.toBeUndefined();
    expect(range.getOoxml).toHaveBeenCalled();
    native = 'A revised ℓ-unit limits.';
    await expect(verifyMutationText(context, range, expected, { paragraph: true })).rejects.toThrow(/read-back differs/);
    native = 'A revised ℓ-unit limit.';
    xml = paragraph(run('l'), 'A revised ', '-unit limit.');
    await expect(verifyMutationText(context, range, expected, { paragraph: true })).rejects.toThrow(/read-back differs/);
    xml = `<w:p xmlns:w="${W}"><w:r><w:t>${expected}</w:t></w:r></w:p>`;
    await expect(verifyMutationText(context, range, expected, { paragraph: true })).rejects.toThrow(/read-back differs/);
});

const target = 'https://example.org/data';
const hyperlinkPackage = (relationship = target) => `<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">`
    + `<pkg:part pkg:name="/word/document.xml"><pkg:xmlData><w:document xmlns:w="${W}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p>`
    + '<w:r><w:t>See </w:t></w:r><w:hyperlink r:id="rId7"><w:r><w:t>dataset</w:t></w:r></w:hyperlink><w:r><w:t> here.</w:t></w:r>'
    + '</w:p></w:body></w:document></pkg:xmlData></pkg:part>'
    + '<pkg:part pkg:name="/word/_rels/document.xml.rels"><pkg:xmlData><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    + `<Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="${relationship}" TargetMode="External"/>`
    + '</Relationships></pkg:xmlData></pkg:part></pkg:package>';

test.each(['', ' \\h'])('native hyperlink field codes require a matching document relationship, switch=%s', (switches) => {
    const state = revisionTextState(hyperlinkPackage());
    const native = `See \u0013HYPERLINK "${target}"${switches}\u0014dataset\u0015 here.`;
    expect(projectNativeText(state, native).text).toBe('See dataset here.');
    expect(projectNativeText(state, 'See dataset here.').text).toBe(state.text);
    expect(projectNativeText(state, native).offsets[native.indexOf(' here.')]).toBe(11);
    expect(state.protectedSpans).toEqual([{ start: 4, end: 11 }]);
    expect(projectNativeText(state, native.replace('example.org', 'example.net'))).toBeNull();
    expect(projectNativeText(state, native.replace('dataset', 'changed'))).toBeNull();
    expect(projectNativeText(state, native.replace('\u0015', ''))).toBeNull();
    expect(revisionTextState(hyperlinkPackage('https://example.net/data')).fingerprint).not.toBe(state.fingerprint);
    const unrelated = revisionTextState(hyperlinkPackage().replace('/word/_rels/document.xml.rels', '/word/_rels/header1.xml.rels'));
    expect(projectNativeText(unrelated, native)).toBeNull();
});
