/** Read-back of written Word content, loaded before the first mutation. */
import { queueCurrentTextRead, queueRevisionRead, resolveRevisionRead, normalizeRevisionText,
    normalizeWordText, textMismatchDetail, MutationSafetyError, projectNativeText, hasNativeFieldCodes } from './word-revisions.js';
import { rangeFingerprintDifference } from './ooxml-fingerprint.js';
import { recordNativeTextMismatch } from './native-text-diagnostics.js';

/** Verify final-view content after a strategy's writes without resetting it. */
export async function verifyMutationText(context, range, expected, options = {}, insertedRanges = []) {
    // Content ranges retain their boundary affinity while Word applies
    // revisions. Reacquire a paragraph's Content after the writes rather
    // than assuming its old range grew to include boundary insertions.
    const target = options.verificationParagraph
        ? options.verificationParagraph.getRange('Content')
        : insertedRanges.filter(Boolean).reduce((scope, inserted) => scope.expandTo(inserted), range);
    if (options.verifiedInlineFingerprint) {
        const read = queueRevisionRead(target, { ...options, singleParagraphContent: true });
        if (!read) throw new MutationSafetyError(new Error('Native object read-back is unavailable. Inspect Word before retrying.'));
        await context.sync();
        const state = resolveRevisionRead(target, read, { ...options, singleParagraphContent: true });
        const wanted = normalizeWordText(expected);
        const mismatches = [];
        if (state.paragraphCount !== 1) mismatches.push(`paragraph count: expected 1, received ${state.paragraphCount}`);
        if (state.text !== wanted) mismatches.push(`prose: ${textMismatchDetail(wanted, state.text)}`);
        if (state.protectedStructure && !state.fieldOnlyProtection) mismatches.push('unsupported protected structure');
        if (state.inlineFingerprint !== options.verifiedInlineFingerprint) {
            mismatches.push(`native objects: ${rangeFingerprintDifference(options.verifiedInlineFingerprint, state.inlineFingerprint)}`);
        }
        if (mismatches.length) {
            const mismatch = mismatches.join('; ');
            // Previously only pre-write serializer mismatches were captured,
            // leaving the actual failed mutation without read-back evidence.
            // Full text/XML stays in the explicitly exported context report.
            try { recordNativeTextMismatch(read, state, { phase: 'post-write', expectedText: wanted, mismatch }); }
            catch { /* diagnostics must not mask an uncertain write */ }
            throw new MutationSafetyError(new Error(`Native object or prose read-back differs from the proposal (${mismatch}). Inspect Word before retrying.`));
        }
        return;
    }
    const current = queueCurrentTextRead(target);
    const read = current ? null : queueRevisionRead(target);
    if (!current && !read) return;
    await context.sync();
    if (current && typeof current.value !== 'string') {
        throw new MutationSafetyError(new Error('Native current-text read-back is unreadable. Inspect Word before retrying.'));
    }
    const comparable = (text) => options.paragraph
        ? normalizeWordText(text) : normalizeRevisionText(text);
    // Current reviewed text is Word's final view, excluding deleted runs.
    // Unlike an OOXML export, it has no synthetic trailing paragraph whose
    // copied revision properties can masquerade as an extra line break.
    let actual = comparable(current ? current.value : resolveRevisionRead(target, read, options).text);
    const wanted = comparable(expected);
    if (current && actual !== wanted && (hasNativeFieldCodes(current.value)
        || /[(\uF000-\uF0FF\r\n\uD835]/.test(current.value) || /\uD835/.test(wanted))) {
        // Fields, symbols and structured equations need fresh XML evidence. Never
        // use the pre-write projection to bless changed field results/glyphs.
        // A supplementary math glyph can serialize as a BMP letterlike symbol,
        // so inspect the expected text as well as Word's native representation.
        const xml = queueRevisionRead(target, options);
        if (xml) {
            await context.sync();
            const state = resolveRevisionRead(target, xml, options);
            const projection = projectNativeText(state, current.value, { trimEnd: !options.paragraph });
            if (projection) actual = comparable(projection.text);
        }
    }
    if (actual !== wanted) {
        const source = current ? 'current text' : 'OOXML';
        throw new MutationSafetyError(new Error(`Native read-back differs from the proposal (${source}: ${textMismatchDetail(wanted, actual)}). Inspect Word before retrying.`));
    }
}
