/** Read-back of written Word content, loaded before the first mutation. */
import { queueCurrentTextRead, queueRevisionRead, resolveRevisionRead, normalizeRevisionText, MutationSafetyError } from './word-revisions.js';

/** Verify final-view content after a strategy's writes without resetting it. */
export async function verifyMutationText(context, range, expected, options = {}, insertedRanges = []) {
    // Content ranges retain their boundary affinity while Word applies
    // revisions. Reacquire a paragraph's Content after the writes rather
    // than assuming its old range grew to include boundary insertions.
    const target = options.verificationParagraph
        ? options.verificationParagraph.getRange('Content')
        : insertedRanges.filter(Boolean).reduce((scope, inserted) => scope.expandTo(inserted), range);
    const current = queueCurrentTextRead(target);
    const read = current ? null : queueRevisionRead(target);
    if (!current && !read) return;
    await context.sync();
    if (current && typeof current.value !== 'string') {
        throw new MutationSafetyError(new Error('Native current-text read-back is unreadable. Inspect Word before retrying.'));
    }
    const comparable = (text) => options.paragraph
        ? String(text || '').replace(/\r\n|\r/g, '\n') : normalizeRevisionText(text);
    // Current reviewed text is Word's final view, excluding deleted runs.
    // Unlike an OOXML export, it has no synthetic trailing paragraph whose
    // copied revision properties can masquerade as an extra line break.
    const actual = comparable(current ? current.value : resolveRevisionRead(target, read, options).text);
    const wanted = comparable(expected);
    if (actual !== wanted) {
        let mismatch = 0;
        while (mismatch < Math.min(actual.length, wanted.length) && actual[mismatch] === wanted[mismatch]) mismatch++;
        const source = current ? 'current text' : 'OOXML';
        throw new MutationSafetyError(new Error(`Native read-back differs from the proposal (${source}: expected ${wanted.length} chars, received ${actual.length}; first mismatch at ${mismatch}). Inspect Word before retrying.`));
    }
}
