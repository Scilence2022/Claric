import { protectChunkText } from '../src/lib/protected-chunk-text.js';
import { checkChunkAmendment } from '../src/lib/amendment-guard.js';
import { textEditHunks, touchesProtectedSpan } from '../src/lib/text-edit-plan.js';

const paragraph = (text, values) => ({ text, protectedSpans: values.map((value) => ({ start: text.indexOf(value), end: text.indexOf(value) + value.length })) });

test('formulas and links survive prose rewriting without model-generated copies of their glyphs', () => {
    const original = paragraph('An old 𝑥_(𝑗) value. See the dataset.', ['𝑥_(𝑗)', 'dataset']);
    const protectedText = protectChunkText({ paragraphs: [original] });
    const masked = protectedText.chunk.paragraphs[0].text;
    expect(masked).not.toContain('𝑥_(𝑗)');
    const restored = protectedText.restore(masked.replace('An old', 'The revised').replace('See the', 'Refer to the'));
    expect(restored).toBe('The revised 𝑥_(𝑗) value. Refer to the dataset.');
    expect(checkChunkAmendment({ paragraphs: [original] }, restored).valid).toBe(true);
});

test('native placeholders cannot be removed, duplicated, reordered, renamed or moved to another paragraph', () => {
    const source = { paragraphs: [paragraph('Use 𝑥 and 𝑦.', ['𝑥', '𝑦']), { text: 'Next paragraph.' }] };
    const protectedText = protectChunkText(source);
    const [a, b] = protectedText.entries.map((e) => e.token);
    const masked = protectedText.chunk.paragraphs.map((p) => p.text).join('\n');
    for (const invalid of [masked.replace(a, ''), masked.replace(a, a + a), masked.replace(a, '𝒙'),
        masked.replace(a, b).replace(' and ' + b, ' and ' + a), masked.replace(a, '').replace('Next', a + ' Next'),
        masked + '\nAnother paragraph.', masked.replace(a, a + 'CLARIC_NATIVE_unknown')]) {
        expect(() => protectedText.restore(invalid)).toThrow();
    }
});

test('a repeated placeholder is reported for every paragraph at once, naming the tokens to fix', () => {
    const source = { paragraphs: [paragraph('Polish depends on 𝐿.', ['𝐿']),
        paragraph('Otherwise 𝛬 lies in 𝜌 modulo 𝑔, the gcd.', ['𝛬', '𝜌', '𝑔'])] };
    const protectedText = protectChunkText(source, { frameParagraphs: true });
    const [l, lambda, rho, g] = protectedText.entries.map((e) => e.token);
    const reply = protectedText.chunk.paragraphs.map((p) => p.text).join('\n')
        .replace(`on ${l}.`, `on ${l}, fixed by ${l}.`)
        .replace(`${g}, the gcd.`, `${g}, where ${g} is the gcd, and ${lambda} holds.`)
        .replace(`in ${rho}`, 'in a class');
    let error;
    try { protectedText.restore(reply); } catch (caught) { error = caught; }
    expect(error.issues).toEqual([
        `Paragraph 1: preserve each native placeholder exactly once and in order; ${l} appears more than once (a native object cannot be copied, so refer back to it in words).`,
        `Paragraph 2: preserve each native placeholder exactly once and in order; ${lambda}, ${g} appear more than once (a native object cannot be copied, so refer back to it in words); missing ${rho}.`,
    ]);
    expect(protectedText.instructions).toContain('A placeholder cannot be repeated');
});

test('identical equations get distinct IDs and native manual line breaks stay inside their paragraph', () => {
    const protectedText = protectChunkText({ paragraphs: [paragraph('First 𝑥.\nSecond 𝑦.', ['𝑥', '𝑦']), paragraph('Third 𝑥.', ['𝑥'])] });
    expect(new Set(protectedText.entries.map((e) => e.token)).size).toBe(3);
    const masked = protectedText.chunk.paragraphs.map((p) => p.text).join('\n');
    expect(protectedText.restore(masked)).toBe('First 𝑥.\nSecond 𝑦.\nThird 𝑥.');
});

test('source text resembling the placeholder protocol stays literal', () => {
    const source = paragraph('CLARIC_NATIVE sample 𝑥.', ['𝑥']);
    const protectedText = protectChunkText({ paragraphs: [source] });
    expect(protectedText.entries[0].token).toContain('CLARIC_NATIVEX');
    expect(protectedText.restore(protectedText.chunk.paragraphs[0].text)).toBe(source.text);
});

test('framed native paragraphs preserve headings, blank paragraphs and manual breaks', () => {
    const source = { paragraphs: [{ text: 'A short heading' }, { text: '' }, paragraph('Use 𝑥.\nKeep this line.', ['𝑥'])] };
    const protectedText = protectChunkText(source, { frameParagraphs: true });
    const masked = protectedText.chunk.paragraphs.map((p) => p.text).join('\n\n');
    expect(protectedText.restore(masked.replace('Use ', 'Consider '))).toBe('A short heading\n\nConsider 𝑥.\nKeep this line.');
    for (const bad of [masked.replace('A short heading', ''), masked.replace('⟦CLARIC_PARAGRAPH_1⟧', ''),
        masked.replace('⟦/CLARIC_PARAGRAPH_1⟧', ''), masked.replace('Use ', 'Use\n'),
        masked + '\nExtra text', 'Preamble\n' + masked,
        masked.replace('Use ', '⟦CLARIC_PARAGRAPH_99⟧Use ')]) {
        expect(() => protectedText.restore(bad)).toThrow();
    }
});

test('paragraph marker collision in source text selects a different protocol prefix', () => {
    const source = paragraph('⟦CLARIC_PARAGRAPH_1⟧ is literal beside 𝑥.', ['𝑥']);
    const protectedText = protectChunkText({ paragraphs: [source] }, { frameParagraphs: true });
    expect(protectedText.chunk.paragraphs[0].text).toContain('CLARIC_PARAGRAPHX_1');
    expect(protectedText.restore(protectedText.chunk.paragraphs[0].text)).toBe(source.text);
});

test('native objects anchor the diff even when surrounding prose repeats their characters', () => {
    const before = 'x old wording x final wording';
    const spans = [{ start: 14, end: 15 }];
    const after = 'New wording x with a final explanation';
    const hunks = textEditHunks(before, after, spans);
    expect(hunks.every((h) => !touchesProtectedSpan(h, spans))).toBe(true);
    let result = before;
    for (const hunk of [...hunks].reverse()) result = result.slice(0, hunk.start) + hunk.text + result.slice(hunk.end);
    expect(result).toBe(after);
});
