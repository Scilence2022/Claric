/** Keep native objects out of generative rewriting; restore only exact,
 * once-only placeholders in their original paragraph and original order. */
import { amendmentParagraphTexts } from './paragraph-alignment.js';

export function protectChunkText(chunk) {
    const source = chunk.paragraphs.map((p) => p.text).join('\n');
    let prefix = 'CLARIC_NATIVE';
    while (source.includes(prefix)) prefix += 'X';
    const entries = [];
    const paragraphs = chunk.paragraphs.map((p, paragraph) => {
        const merged = [];
        for (const span of [...(p.protectedSpans || [])].sort((a, b) => a.start - b.start)) {
            if (span.end <= span.start) continue;
            const prior = merged.at(-1);
            if (prior && span.start < prior.end) prior.end = Math.max(prior.end, span.end);
            else merged.push({ ...span });
        }
        let text = '', offset = 0;
        for (const span of merged) {
            const token = `⟦${prefix}_${paragraph + 1}_${entries.length + 1}⟧`;
            entries.push({ paragraph, token, text: p.text.slice(span.start, span.end) });
            text += p.text.slice(offset, span.start) + token;
            offset = span.end;
        }
        return { ...p, text: text + p.text.slice(offset), protectedSpans: [] };
    });
    return { chunk: { ...chunk, paragraphs }, entries,
        restore(amendment) {
            if (!entries.length) return amendment;
            const lines = amendmentParagraphTexts(paragraphs.map((p) => p.text), amendment);
            // A native object cannot migrate across paragraph boundaries.
            if (lines.length !== paragraphs.length) throw new Error('Preserve native paragraph boundaries and every native placeholder.');
            for (let i = 0; i < lines.length; i++) {
                const expected = entries.filter((entry) => entry.paragraph === i);
                const actual = lines[i].match(new RegExp(`⟦${prefix}_[^⟧]*⟧`, 'g')) || [];
                if (JSON.stringify(actual) !== JSON.stringify(expected.map((entry) => entry.token))) {
                    throw new Error(`Paragraph ${i + 1}: preserve each native placeholder exactly once and in order.`);
                }
                for (const entry of expected) lines[i] = lines[i].replace(entry.token, () => entry.text);
                if (lines[i].includes(prefix)) throw new Error('Return only the supplied native placeholders, without inventing or changing their IDs.');
            }
            return lines.join('\n');
        } };
}
