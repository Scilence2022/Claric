/** Map final-view differences to verified Word ranges, preserving earlier revisions. */
import { textEditHunks, touchesProtectedSpan } from '../text-edit-plan.js';
import { recordNativeTextMismatch } from '../native-text-diagnostics.js';
import { queueRevisionRead, resolveRevisionRead, revisionTextState, normalizeRevisionText,
    RevisionSafetyError, MutationSafetyError, verifyMutationText, loadMutationVerifier, hostSupportsTrackedChanges,
    normalizeWordText, textMismatchDetail, queueCurrentTextRead, wordErrorDetail, projectNativeText } from '../word-revisions.js';

const MAX_TARGETS = 256;
const MAX_CANDIDATES = 1024;
const READ_BATCH_SIZE = 24;

/** A native prefix ending inside a field result or instruction. Editable
 * targets never lie inside a field; Mac Word can still export such a prefix
 * as the whole field, giving it the offset of the prose after the field. */
function endsInsideNativeField(text) {
    let depth = 0;
    for (const char of text) {
        if (char === '\u0013') depth++;
        else if (char === '\u0015') depth = Math.max(0, depth - 1);
    }
    return depth > 0;
}

async function syncLocationRead(context, stage) {
    try { await context.sync(); }
    catch (error) {
        throw new RevisionSafetyError(`Could not read ${stage} (${wordErrorDetail(error)}). No text was written.`);
    }
}

function planEdits(state, after) {
    const hunks = textEditHunks(state.text, after, state.protectedSpans).map((hunk) => ({ ...hunk, targets: [], anchor: null, location: null }));
    for (const hunk of hunks) {
        if (/[\n\t]/.test(hunk.text + state.text.slice(hunk.start, hunk.end))) {
            throw new RevisionSafetyError('Cannot change paragraph or tab boundaries with unresolved revisions.');
        }
        if (touchesProtectedSpan(hunk, state.protectedSpans)) {
            throw new RevisionSafetyError('The proposed text changes a native field result or equation. Preserve formulas, citations, bibliography fields and links; edit surrounding prose only.');
        }
    }
    const targets = [];
    function target(start, text) {
        const item = { start, text, range: null };
        targets.push(item);
        if (targets.length > MAX_TARGETS) throw new RevisionSafetyError('Revision edit exceeds the safe range limit. Use a smaller selection.');
        return item;
    }
    for (const hunk of hunks) {
        if (hunk.start === hunk.end) {
            // Anchor to visible text, never the beginning of a deleted run.
            const prior = Array.from(state.text.slice(0, hunk.start)).pop();
            const next = Array.from(state.text.slice(hunk.start))[0];
            const protectedAt = (position) => state.protectedSpans.some(({ start, end }) => position >= start && position < end);
            if (prior && prior !== '\n' && prior !== '\t' && !protectedAt(hunk.start - prior.length)) {
                hunk.anchor = target(hunk.start - prior.length, prior);
                hunk.location = Word.InsertLocation.after;
            } else if (prior && protectedAt(hunk.start - prior.length) && next && next !== '\n' && next !== '\t'
                && !protectedAt(hunk.start)) {
                // Text inserted before a character inherits the preceding
                // character's run; Mac Word then places it inside the
                // equation or link that ends here. Keep the write in prose:
                // slide an insertion that begins with the next character past
                // it, otherwise replace that character together with the text.
                const anchor = target(hunk.start, next);
                if (Array.from(hunk.text)[0] === next) {
                    hunk.text = hunk.text.slice(next.length) + next;
                    hunk.start += next.length;
                    hunk.end = hunk.start;
                    hunk.anchor = anchor;
                    hunk.location = Word.InsertLocation.after;
                } else {
                    hunk.text += next;
                    hunk.end = hunk.start + next.length;
                    hunk.targets.push(anchor);
                }
            } else if (next && next !== '\n' && next !== '\t' && !protectedAt(hunk.start)) {
                hunk.anchor = target(hunk.start, next);
                hunk.location = Word.InsertLocation.before;
            } else if (!hunk.start || hunk.start === state.text.length) {
                hunk.location = hunk.start ? Word.InsertLocation.after : Word.InsertLocation.before;
            }
            else throw new RevisionSafetyError('No safe visible insertion anchor. Select one paragraph.');
        } else {
            for (const segment of state.segments) {
                const start = Math.max(hunk.start, segment.start);
                const end = Math.min(hunk.end, segment.start + segment.text.length);
                // Word searches are capped at 255 UTF-16 code units.
                let pos = start;
                while (pos < end) {
                    let stop = Math.min(pos + 200, end);
                    const code = state.text.charCodeAt(stop - 1);
                    if (stop < end && code >= 0xd800 && code <= 0xdbff) stop--;
                    hunk.targets.push(target(pos, state.text.slice(pos, stop)));
                    pos = stop;
                }
            }
            if (hunk.targets.reduce((n, t) => n + t.text.length, 0) !== hunk.end - hunk.start) {
                throw new RevisionSafetyError('The edit crosses an uneditable revision boundary.');
            }
        }
    }
    return { hunks, targets };
}

async function locate(context, scope, targets, options, allowSplit = true, nativeProjection = null, xmlBaseline = null) {
    const searches = new Map();
    for (const { text } of targets) {
        if (searches.has(text)) continue;
        const matches = scope.search(text, { matchCase: true, matchWholeWord: false, matchWildcards: false,
            ignorePunct: false, ignoreSpace: false, matchPrefix: false, matchSuffix: false });
        matches.load('items');
        searches.set(text, matches);
    }
    if (searches.size) await syncLocationRead(context, 'native search results');
    const candidates = [];
    for (const [text, matches] of searches) {
        for (const range of matches.items) candidates.push({ text, range });
    }
    if (candidates.length > MAX_CANDIDATES) throw new RevisionSafetyError('Too many ambiguous revision matches. Use a smaller selection.');
    const verified = new Map();
    const duplicates = new Map();
    const needed = new Set(targets.map((target) => JSON.stringify([target.start, target.text])));
    // Verify both content and final-view position. Nth raw search matches
    // are unsafe when old deletions contain the same word as visible text.
    for (let start = 0; start < candidates.length; start += READ_BATCH_SIZE) {
        const batch = candidates.slice(start, start + READ_BATCH_SIZE).map((candidate) => {
            const prefix = scope.getRange(Word.RangeLocation.start).expandTo(candidate.range.getRange(Word.RangeLocation.start));
            // OOXML exports can include a terminal empty paragraph that is
            // indistinguishable from a genuine paragraph-start fragment in
            // a multi-paragraph prefix. Current-view native text retains
            // the exact range boundary without that export padding.
            const currentPrefix = queueCurrentTextRead(prefix);
            const currentContent = queueCurrentTextRead(candidate.range);
            return { ...candidate, currentPrefix, currentContent, prefixRange: prefix,
                xml: currentContent ? null : queueRevisionRead(candidate.range, { paragraph: true }),
                prefix: currentPrefix ? null : prefix.getOoxml() };
        });
        await syncLocationRead(context, 'current-view search ranges');
        if (xmlBaseline !== null) {
            // A collapsed paragraph-start range has a proven empty native
            // prefix. Mac Word can throw GeneralException when exporting
            // that empty range; never request its OOXML.
            for (const candidate of batch) {
                if (candidate.currentPrefix && typeof candidate.currentPrefix.value !== 'string') {
                    throw new RevisionSafetyError('The current-view revision prefix is unreadable. Generate a fresh proposal.');
                }
                if (candidate.currentPrefix && normalizeWordText(candidate.currentPrefix.value)) {
                    candidate.prefix = candidate.prefixRange.getOoxml();
                }
            }
            if (batch.some((candidate) => candidate.currentPrefix && candidate.prefix)) {
                await syncLocationRead(context, 'nonempty XML position prefixes');
            }
        }
        const selected = [];
        for (const candidate of batch) {
            if (candidate.currentPrefix && typeof candidate.currentPrefix.value !== 'string') {
                throw new RevisionSafetyError('The current-view revision prefix is unreadable. Generate a fresh proposal.');
            }
            if (candidate.currentContent && typeof candidate.currentContent.value !== 'string') {
                throw new RevisionSafetyError('The current-view search target is unreadable. Generate a fresh proposal.');
            }
            if (xmlBaseline !== null && !candidate.currentContent) {
                throw new RevisionSafetyError('Native prose verification is unavailable. No text was written.');
            }
            if (candidate.currentContent && normalizeWordText(candidate.currentContent.value) !== candidate.text) continue;
            const prefix = candidate.currentPrefix && normalizeWordText(candidate.currentPrefix.value);
            if (prefix && endsInsideNativeField(prefix)) continue;
            // Word includes REF instructions and legacy symbol placeholders
            // in native prefixes. Their raw lengths are not visible offsets.
            // Match the prefix to the independently verified whole stream.
            if (nativeProjection && prefix !== null && !nativeProjection.native.startsWith(prefix)) continue;
            const prefixState = xmlBaseline !== null && candidate.currentPrefix && !prefix
                ? { text: '', paragraphCount: 1 }
                : candidate.currentPrefix && xmlBaseline === null ? null : revisionTextState(candidate.prefix.value,
                    xmlBaseline === null ? options : { ...options, singleParagraphContent: true });
            if (xmlBaseline !== null && prefixState.paragraphCount !== 1) continue;
            const prefixText = prefixState?.text;
            if (xmlBaseline !== null && !xmlBaseline.startsWith(prefixText)) continue;
            const offset = candidate.currentPrefix && xmlBaseline === null
                ? nativeProjection ? nativeProjection.offsets[prefix.length] : prefix.length
                : prefixText.length;
            const key = JSON.stringify([offset, candidate.text]);
            if (!needed.has(key)) continue;
            selected.push({ ...candidate, key, xml: candidate.xml || queueRevisionRead(candidate.range, { paragraph: true }) });
        }
        // Full OOXML exports are expensive in macOS Word. Read only the
        // candidates at requested native offsets, rather than every repeated
        // character match and every prefix in the paragraph.
        if (selected.some((candidate) => candidate.currentContent)) await syncLocationRead(context, 'verified prose target XML');
        for (const candidate of selected) {
            const content = resolveRevisionRead(candidate.range, candidate.xml, { paragraph: true });
            if (content.hasHiddenContent || (content.protectedStructure && !(options.nativeTargets && content.fieldOnlyProtection))
                || content.protectedSpans.some(({ start, end }) => end > start) || content.text !== candidate.text) continue;
            const key = candidate.key;
            if (verified.has(key)) {
                duplicates.set(key, [...(duplicates.get(key) || [{ ...candidate, range: verified.get(key) }]), candidate]);
                continue;
            }
            verified.set(key, candidate.range);
        }
    }
    if (duplicates.size) await resolveDuplicates(context, scope, duplicates, verified, options, nativeProjection, xmlBaseline);
    for (const target of targets) {
        target.range = verified.get(JSON.stringify([target.start, target.text]));
    }
    const missing = targets.filter((target) => !target.range);
    if (!missing.length) return;
    if (!allowSplit) throw new RevisionSafetyError('Unverified visible revision target. Generate a fresh proposal.');
    // Greedy Word searches can skip a self-overlapping occurrence of a long
    // piece. Locate its scalar endpoints, then verify the entire union.
    const endpoints = missing.map((target) => {
        const scalars = Array.from(target.text);
        const first = { start: target.start, text: scalars[0], range: null };
        const lastText = scalars[scalars.length - 1];
        const last = { start: target.start + target.text.length - lastText.length, text: lastText, range: null };
        return { target, first, last };
    });
    await locate(context, scope, endpoints.flatMap(({ first, last }) => [first, last]), options, false, nativeProjection, xmlBaseline);
    for (let start = 0; start < endpoints.length; start += READ_BATCH_SIZE) {
      const unions = endpoints.slice(start, start + READ_BATCH_SIZE).map(({ target, first, last }) => {
          target.range = first.range.expandTo(last.range);
          return { target, xml: queueRevisionRead(target.range, { paragraph: true }) };
      });
      await syncLocationRead(context, 'combined prose target XML');
      for (const { target, xml } of unions) {
        const content = resolveRevisionRead(target.range, xml, { paragraph: true });
        if ((xml?.currentText && normalizeWordText(xml.currentText.value) !== target.text)
            || content.hasHiddenContent || (content.protectedStructure && !(options.nativeTargets && content.fieldOnlyProtection))
            || content.protectedSpans.some(({ start, end }) => end > start) || content.text !== target.text) {
            throw new RevisionSafetyError('Located text crosses a revision boundary. Generate a fresh proposal.');
        }
      }
    }
}

/** Two native ranges proved the same start offset. Keep only a candidate
 * whose end boundary also maps to the expected final-view offset. */
async function resolveDuplicates(context, scope, duplicates, verified, options, nativeProjection, xmlBaseline) {
    const reads = [];
    for (const [key, candidates] of duplicates) {
        const [offset, text] = JSON.parse(key);
        for (const candidate of candidates) {
            const through = scope.getRange(Word.RangeLocation.start).expandTo(candidate.range.getRange(Word.RangeLocation.end));
            const current = xmlBaseline === null ? queueCurrentTextRead(through) : null;
            reads.push({ key, end: offset + text.length, candidate, current, xml: current ? null : through.getOoxml() });
        }
    }
    if (reads.length > MAX_CANDIDATES) throw new RevisionSafetyError('Too many ambiguous revision matches. Use a smaller selection.');
    await syncLocationRead(context, 'ambiguous target boundaries');
    const accepted = new Map();
    for (const { key, end, candidate, current, xml } of reads) {
        let valid;
        if (current) {
            if (typeof current.value !== 'string') continue;
            const native = normalizeWordText(current.value);
            valid = !endsInsideNativeField(native) && (nativeProjection
                ? nativeProjection.native.startsWith(native) && nativeProjection.offsets[native.length] === end
                : native.length === end);
        } else {
            const state = revisionTextState(xml.value, xmlBaseline === null ? options : { ...options, singleParagraphContent: true });
            valid = xmlBaseline === null ? state.text.length === end
                : state.paragraphCount === 1 && state.text === xmlBaseline.slice(0, end);
        }
        if (valid) accepted.set(key, accepted.has(key) ? null : candidate.range);
    }
    for (const key of duplicates.keys()) {
        if (!accepted.get(key)) throw new RevisionSafetyError('Ambiguous visible revision target. Generate a fresh proposal.');
        verified.set(key, accepted.get(key));
    }
}

/** Independently prove every editable character when a host's whole-text
 * serializer cannot represent native inline objects. Protected results are
 * excluded; no native prose discrepancy or ambiguous range is tolerated. */
function proseProofTargets(state) {
    const targets = [];
    for (const segment of state.segments) {
        let intervals = [{ start: segment.start, end: segment.start + segment.text.length }];
        for (const span of state.protectedSpans) intervals = intervals.flatMap(({ start, end }) =>
            span.end <= start || span.start >= end ? [{ start, end }]
                : [{ start, end: Math.max(start, span.start) }, { start: Math.min(end, span.end), end }].filter((s) => s.end > s.start));
        for (const interval of intervals) {
            let start = interval.start;
            while (start < interval.end) {
                let end = Math.min(start + 200, interval.end);
                if (end < interval.end && /[\uD800-\uDBFF]/.test(state.text[end - 1])) end--;
                targets.push({ start, text: state.text.slice(start, end), range: null });
                start = end;
                if (targets.length > MAX_TARGETS) throw new RevisionSafetyError('Too many prose ranges to verify safely. Use a smaller selection.');
            }
        }
    }
    return targets;
}

/** Returns null for pristine legacy ranges; nativeTargets verifies pristine ranges too.
 * All errors here must bypass destructive fallbacks.
 * @param {Word.RequestContext} context @param {Word.Range} range
 * @param {string} before @param {string} after @param {function} [log] @param {object} [options]
 */
export async function tryRevisionDiff(context, range, before, after, log = () => {}, options = {}) {
    let mutationAttempted = false;
    try {
        await loadMutationVerifier();
        // nativeTargets + paragraph is used only for one proven paragraph's
        // Content, so a copied empty export paragraph is not a real boundary.
        const readOptions = options.paragraph && options.nativeTargets ? { ...options, singleParagraphContent: true } : options;
        const read = queueRevisionRead(range, readOptions);
        if (!read) return null;
        await context.sync();
        const state = resolveRevisionRead(range, read, readOptions);
        // Paragraph reads preserve actual manual line breaks. Whole-range
        // protocols normalize the native terminal paragraph boundary.
        const comparable = (text) => options.paragraph
            ? normalizeWordText(text) : normalizeRevisionText(text);
        if (comparable(before) !== comparable(state.text)) {
            throw new RevisionSafetyError('Current text differs from the proposal baseline. Generate a fresh proposal.');
        }
        if (read.currentText && typeof read.currentText.value !== 'string') {
            throw new RevisionSafetyError('Native current text is unreadable before writing. Draft a new proposal.');
        }
        const nativeProjection = read.currentText && projectNativeText(state,
            read.currentText.value, { trimEnd: !options.paragraph });
        if (read.currentText && !nativeProjection) recordNativeTextMismatch(read, state);
        const xmlPositions = !!(read.currentText && !nativeProjection && options.paragraph && state.paragraphCount === 1
            && state.inlineFingerprint && state.protectedSpans.length && (!state.protectedStructure || state.fieldOnlyProtection));
        if (read.currentText && !nativeProjection && !xmlPositions) {
            throw new RevisionSafetyError(`Native text and XML disagree before writing (${textMismatchDetail(comparable(state.text), comparable(read.currentText.value))}). Draft a new proposal.`);
        }
        if (!state.hasRevisions && !options.nativeTargets && !state.protectedSpans.length) return null;
        if (!state.fingerprint || (state.protectedStructure && !state.fieldOnlyProtection)) {
            throw new RevisionSafetyError('Structural revisions or protected objects prevent this edit.');
        }
        state.text = comparable(state.text);
        const expected = comparable(after);
        const { hunks, targets } = planEdits(state, expected);
        if (!hunks.length) return { strategy: 'revision', insertions: 0, deletions: 0, replacements: 0 };
        const ownsTracking = options.trackChanges !== false;
        if (ownsTracking && !hostSupportsTrackedChanges()) throw new RevisionSafetyError('This Word host cannot enable tracked changes.');
        if (xmlPositions) {
            await locate(context, range, [...proseProofTargets(state), ...targets], options, true, null, state.text);
            log('Verified all editable prose against native ranges and OOXML positions; native objects remain protected.', 'info');
        } else await locate(context, range, targets, options, true, nativeProjection);
        if (ownsTracking) context.document.load('changeTrackingMode');
        const baseline = queueRevisionRead(range, readOptions);
        await context.sync();
        if (resolveRevisionRead(range, baseline, readOptions).fingerprint !== state.fingerprint) {
            throw new RevisionSafetyError('Revision state changed during location. Generate a fresh proposal.');
        }
        // A section can locate every revised paragraph before its first
        // write. The write pass repeats location and baseline validation so
        // this read-only check never becomes a stale permission to mutate.
        if (options.preflightOnly) return { strategy: 'revision', insertions: 0, deletions: 0, replacements: 0 };
        const previousMode = ownsTracking ? context.document.changeTrackingMode : null;
        let insertions = 0;
        let deletions = 0;
        let replacements = 0;
        const insertedRanges = [];
        try {
            if (ownsTracking) {
                try {
                    context.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
                    await context.sync();
                } catch (error) {
                    throw new RevisionSafetyError(`Could not enable tracked changes (${error.message}); no text was written.`);
                }
            }
            for (const hunk of [...hunks].reverse()) {
                mutationAttempted = true;
                if (hunk.start === hunk.end) {
                    const anchor = hunk.anchor?.range || range.getRange(hunk.location === Word.InsertLocation.before
                        ? Word.RangeLocation.start : Word.RangeLocation.end);
                    insertedRanges.push(anchor.insertText(hunk.text, hunk.location));
                    insertions++;
                } else {
                    for (let i = hunk.targets.length - 1; i >= 0; i--) {
                        const target = hunk.targets[i].range;
                        if (i === 0 && hunk.text) { insertedRanges.push(target.insertText(hunk.text, Word.InsertLocation.replace)); replacements++; }
                        else { target.delete(); deletions++; }
                    }
                }
            }
            await context.sync();
            await verifyMutationText(context, range, after, xmlPositions
                ? { ...options, verifiedInlineFingerprint: state.inlineFingerprint } : options, insertedRanges);
        } finally {
            if (ownsTracking) { context.document.changeTrackingMode = previousMode; await context.sync(); }
        }
        log(options.nativeTargets ? 'Verified edits against native text and positions.' : 'Verified edits against current revision text.', 'info');
        return { strategy: 'revision', insertions, deletions, replacements };
    } catch (error) {
        if (mutationAttempted) throw error instanceof MutationSafetyError ? error : new MutationSafetyError(error);
        throw error instanceof RevisionSafetyError ? error
            : new RevisionSafetyError(`Revision-safe application failed (${wordErrorDetail(error)}). Inspect Word before retrying.`);
    }
}
