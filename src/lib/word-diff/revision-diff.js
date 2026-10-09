/** Map final-view differences to verified Word ranges, preserving earlier revisions. */
import DiffMatchPatch from '../vendor/diff-match-patch.js';
import { queueRevisionRead, resolveRevisionRead, revisionTextState, normalizeRevisionText,
    RevisionSafetyError } from '../word-revisions.js';

const MAX_TARGETS = 256;
const MAX_CANDIDATES = 1024;

/** Diff Unicode scalars so an emoji replacement never searches for half a surrogate. */
function scalarDiff(before, after) {
    const values = [''];
    const ids = new Map();
    const encode = (text) => Array.from(text).map((scalar) => {
        if (!ids.has(scalar)) { ids.set(scalar, values.length); values.push(scalar); }
        if (values.length > 65535) throw new RevisionSafetyError('Too many distinct characters for a safe revision diff.');
        return String.fromCharCode(ids.get(scalar));
    }).join('');
    const diffs = new DiffMatchPatch().diff_main(encode(before), encode(after), false);
    return diffs.map(([op, encoded]) => [op, Array.from(encoded).map((id) => values[id.charCodeAt(0)]).join('')]);
}

function planEdits(state, after) {
    const hunks = [];
    let offset = 0;
    let pending = null;
    for (const [op, text] of scalarDiff(state.text, after)) {
        if (op === 0) { pending = null; offset += text.length; continue; }
        if (!pending) { pending = { start: offset, end: offset, text: '', targets: [], anchor: null, location: null }; hunks.push(pending); }
        if (text.includes('\n') || text.includes('\t')) {
            throw new RevisionSafetyError('A revision-bearing range cannot change paragraph or tab boundaries. Revise its text in place.');
        }
        if (op === -1) { offset += text.length; pending.end = offset; }
        else pending.text += text;
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
            if (prior && prior !== '\n' && prior !== '\t') {
                hunk.anchor = target(hunk.start - prior.length, prior);
                hunk.location = Word.InsertLocation.after;
            } else if (next && next !== '\n' && next !== '\t') {
                hunk.anchor = target(hunk.start, next);
                hunk.location = Word.InsertLocation.before;
            } else if (!state.text) hunk.location = Word.InsertLocation.before;
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

async function locate(context, scope, targets, allowSplit = true) {
    const searches = new Map();
    for (const { text } of targets) {
        if (searches.has(text)) continue;
        const matches = scope.search(text, { matchCase: true, matchWholeWord: false, matchWildcards: false });
        matches.load('items');
        searches.set(text, matches);
    }
    if (searches.size) await context.sync();
    const candidates = [];
    for (const [text, matches] of searches) {
        for (const range of matches.items) candidates.push({ text, range });
    }
    if (candidates.length > MAX_CANDIDATES) throw new RevisionSafetyError('Too many ambiguous revision matches. Use a smaller selection.');
    const verified = new Map();
    // Verify both content and final-view position. Nth raw search matches
    // are unsafe when old deletions contain the same word as visible text.
    for (let start = 0; start < candidates.length; start += 64) {
        const batch = candidates.slice(start, start + 64).map((candidate) => {
            const prefix = scope.getRange(Word.RangeLocation.start).expandTo(candidate.range.getRange(Word.RangeLocation.start));
            return { ...candidate, xml: candidate.range.getOoxml(), prefix: prefix.getOoxml() };
        });
        await context.sync();
        for (const candidate of batch) {
            const content = revisionTextState(candidate.xml.value);
            if (content.hasHiddenContent || content.protectedStructure || content.text !== candidate.text) continue;
            const offset = revisionTextState(candidate.prefix.value).text.length;
            const key = JSON.stringify([offset, candidate.text]);
            if (verified.has(key)) throw new RevisionSafetyError('Ambiguous visible revision target. Generate a fresh proposal.');
            verified.set(key, candidate.range);
        }
    }
    for (const target of targets) {
        target.range = verified.get(JSON.stringify([target.start, target.text]));
    }
    const missing = targets.filter((target) => !target.range);
    if (!missing.length) return;
    if (!allowSplit) throw new RevisionSafetyError('A visible revision target could not be verified. Generate a fresh proposal.');
    // Greedy Word searches can skip a self-overlapping occurrence of a long
    // piece. Locate its scalar endpoints, then verify the entire union.
    const endpoints = missing.map((target) => {
        const scalars = Array.from(target.text);
        const first = { start: target.start, text: scalars[0], range: null };
        const lastText = scalars[scalars.length - 1];
        const last = { start: target.start + target.text.length - lastText.length, text: lastText, range: null };
        return { target, first, last };
    });
    await locate(context, scope, endpoints.flatMap(({ first, last }) => [first, last]), false);
    const unions = endpoints.map(({ target, first, last }) => {
        target.range = first.range.expandTo(last.range);
        return { target, xml: target.range.getOoxml() };
    });
    await context.sync();
    for (const { target, xml } of unions) {
        const content = revisionTextState(xml.value);
        if (content.hasHiddenContent || content.protectedStructure || content.text !== target.text) {
            throw new RevisionSafetyError('The located revision span crosses hidden or changed text. Generate a fresh proposal.');
        }
    }
}

/** Returns null for pristine ranges; all errors here must bypass destructive fallbacks.
 * @param {Word.RequestContext} context @param {Word.Range} range
 * @param {string} before @param {string} after @param {function} [log] @param {object} [options]
 */
export async function tryRevisionDiff(context, range, before, after, log = () => {}, options = {}) {
    try {
        const read = queueRevisionRead(range);
        if (!read) return null;
        await context.sync();
        const state = resolveRevisionRead(range, read);
        if (!state.hasRevisions) return null;
        if (!state.fingerprint || state.protectedStructure) {
            throw new RevisionSafetyError('This range contains structural revisions or protected objects. Choose a plain-text paragraph.');
        }
        if (normalizeRevisionText(before) !== state.text) {
            throw new RevisionSafetyError('The current revision text differs from the proposal baseline. Generate a fresh proposal.');
        }
        const expected = normalizeRevisionText(after);
        const { hunks, targets } = planEdits(state, expected);
        if (!hunks.length) return { strategy: 'revision', insertions: 0, deletions: 0, replacements: 0 };
        const ownsTracking = options.trackChanges !== false;
        if (ownsTracking && !Word.ChangeTrackingMode) throw new RevisionSafetyError('This Word host cannot enable tracked changes.');
        await locate(context, range, targets);
        if (ownsTracking) context.document.load('changeTrackingMode');
        const baseline = range.getOoxml();
        await context.sync();
        if (revisionTextState(baseline.value).fingerprint !== state.fingerprint) {
            throw new RevisionSafetyError('Revision state changed while locating the edit. Generate a fresh proposal.');
        }
        const previousMode = context.document.changeTrackingMode;
        let insertions = 0;
        let deletions = 0;
        let replacements = 0;
        try {
            if (ownsTracking) context.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
            for (const hunk of [...hunks].reverse()) {
                if (hunk.start === hunk.end) {
                    const anchor = hunk.anchor?.range || range.getRange(Word.RangeLocation.end);
                    anchor.insertText(hunk.text, hunk.location);
                    insertions++;
                } else {
                    for (let i = hunk.targets.length - 1; i >= 0; i--) {
                        const target = hunk.targets[i].range;
                        if (i === 0 && hunk.text) { target.insertText(hunk.text, Word.InsertLocation.replace); replacements++; }
                        else { target.delete(); deletions++; }
                    }
                }
            }
            await context.sync();
            const result = range.getOoxml();
            await context.sync();
            if (revisionTextState(result.value).text !== expected) {
                throw new RevisionSafetyError('Revision read-back differs from the proposal. Inspect Word before retrying.');
            }
        } finally {
            if (ownsTracking) { context.document.changeTrackingMode = previousMode; await context.sync(); }
        }
        log('Applied and verified changes against current revision text; unrelated revisions were preserved.', 'info');
        return { strategy: 'revision', insertions, deletions, replacements };
    } catch (error) {
        throw error instanceof RevisionSafetyError ? error
            : new RevisionSafetyError(`Revision-safe application failed (${error.message}). Inspect Word before retrying.`);
    }
}
