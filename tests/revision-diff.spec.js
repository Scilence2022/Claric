/** @jest-environment jsdom */
import { tryRevisionDiff } from '../src/lib/word-diff/revision-diff.js';
import { applyTokenMapStrategy, applySentenceDiffStrategy, applyCharDiffStrategy } from '../src/lib/word-diff/index.js';
import { revisionTextState, RevisionSafetyError, MutationSafetyError } from '../src/lib/word-revisions.js';
import { parseDocument } from '../src/lib/document-parser.js';
import { readSelectionText, readSelectionContent, readSelectionSnippet, applySelectionAmendment } from '../src/taskpane/word-actions.js';
import { readDocumentEditSnapshot, anchorDocumentEdit, applyDocumentEdit } from '../src/taskpane/document-edit-actions.js';
import { applyChunkResults } from '../src/lib/reassembler.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const escape = (text) => text.replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;');

/** Word searches the physical text, including deleted runs. Writes retain
 * deleted atoms and adjust live range boundaries, unlike flat string mocks. */
function world(parts, { padding = false, nativeReviewed = false, nativeParagraphs = false,
    stickyContent = false, revisedPadding = false } = {}) {
    let sequence = 0;
    const atoms = parts.flatMap(({ text, type = 'plain', id = ++sequence }) =>
        Array.from(text).map((text) => ({ text, type, id })));
    const ranges = [];
    const writes = [];
    const bookmarks = new Map();
    const xml = (slice) => {
        const groups = [[]];
        for (const atom of slice) {
            if (nativeParagraphs && atom.text === '\n') groups.push([]);
            else groups.at(-1).push(atom);
        }
        const content = groups.map((group) => `<w:p xmlns:w="${W}">` + group.map((atom) => {
        const run = `<w:r><w:${atom.type === 'del' ? 'delText' : 't'}>${escape(atom.text)}</w:${atom.type === 'del' ? 'delText' : 't'}></w:r>`;
        return atom.type === 'plain' ? run : `<w:${atom.type} w:id="${atom.id}" w:author="Reviewer">${run}</w:${atom.type}>`;
        }).join('') + '</w:p>').join('');
        const end = revisedPadding && slice.some((atom) => atom.type !== 'plain')
            ? '<w:p><w:pPr><w:rPr><w:ins w:id="999"/></w:rPr></w:pPr></w:p>' : '<w:p/>';
        return padding ? `<w:document xmlns:w="${W}"><w:body>${content}${end}</w:body></w:document>` : content;
    };
    const table = { isNullObject: true, load: jest.fn() };
    const cell = { isNullObject: true, load: jest.fn() };
    const collection = (items) => ({ items, load: jest.fn(), getFirst: () => items[0], getLast: () => items.at(-1) });
    function makeRange(start, end, scope = false) {
        const span = { start, end, scope };
        ranges.push(span);
        const range = {
            isNullObject: false, load: jest.fn(), parentTableOrNullObject: table, parentTableCellOrNullObject: cell,
            inlinePictures: collection([]),
            get text() { return atoms.slice(span.start, span.end).map((a) => a.text).join(''); },
            getOoxml: jest.fn(() => ({ value: xml(atoms.slice(span.start, span.end)) })),
            ...(nativeReviewed ? { getReviewedText: jest.fn(() => ({ value:
                atoms.slice(span.start, span.end).filter((atom) => atom.type !== 'del').map((atom) => atom.text).join('').replace(/\n/g, '\r') })) } : {}),
            getRange: (location) => location === 'Start' ? makeRange(span.start, span.start)
                : location === 'End' ? makeRange(span.end, span.end) : range,
            expandTo: (other) => makeRange(Math.min(span.start, other._span.start), Math.max(span.end, other._span.end)),
            insertBookmark: (name) => bookmarks.set(name, range),
            search: jest.fn((text) => {
                const current = atoms.slice(span.start, span.end);
                const raw = current.map((a) => a.text).join('');
                const matches = [];
                let pos = 0;
                while ((pos = raw.indexOf(text, pos)) !== -1) {
                    const first = Array.from(raw.slice(0, pos)).length;
                    const last = first + Array.from(text).length;
                    matches.push(makeRange(span.start + first, span.start + last));
                    pos += text.length;
                }
                return collection(matches);
            }),
            getTextRanges: jest.fn(() => {
                const raw = atoms.slice(span.start, span.end).map((atom) => atom.text).join('');
                return collection(Array.from(raw.matchAll(/.+?(?:\. {1,2}|$)/gs), (match) =>
                    makeRange(span.start + match.index, span.start + match.index + match[0].length)));
            }),
            delete: jest.fn(() => {
                writes.push({ type: 'delete', start: span.start, end: span.end, mode: document.changeTrackingMode });
                for (const atom of atoms.slice(span.start, span.end)) {
                    if (atom.type === 'del') throw new Error('Attempted to edit an earlier deletion');
                    atom.type = 'del'; atom.id = ++sequence;
                }
            }),
            insertText: jest.fn((text, location) => {
                writes.push({ type: 'insert', text, location, start: span.start, end: span.end, mode: document.changeTrackingMode });
                if (location === 'Replace') range.delete();
                const at = location === 'Before' ? span.start : span.end;
                const added = Array.from(text).map((text) => ({ text, type: 'ins', id: ++sequence }));
                atoms.splice(at, 0, ...added);
                for (const live of ranges) {
                    if (live.start >= at && live !== span && !live.scope) live.start += added.length;
                    if (live.end > at || live === span || live.scope) live.end += added.length;
                }
                return makeRange(at, at + added.length);
            }),
            _span: span,
        };
        return range;
    }
    const scope = makeRange(0, atoms.length, !stickyContent);
    const paragraphScope = stickyContent ? makeRange(0, atoms.length, true) : scope;
    const paragraph = {
        get text() { return paragraphScope.text; }, style: 'Normal', styleBuiltIn: 'Normal', isListItem: false,
        parentTableOrNullObject: table, load: jest.fn(), untrack: jest.fn(), getOoxml: paragraphScope.getOoxml,
        getRange: () => stickyContent ? makeRange(paragraphScope._span.start, paragraphScope._span.end) : scope, delete: jest.fn(),
    };
    scope.paragraphs = collection([paragraph]);
    const document = { body: { paragraphs: collection([paragraph]) }, getSelection: () => scope,
        load: jest.fn(), changeTrackingMode: 'TrackMineOnly',
        getBookmarkRangeOrNullObject: (name) => bookmarks.get(name) || { isNullObject: true, load: jest.fn() },
        deleteBookmark: (name) => bookmarks.delete(name) };
    const context = { document, sync: jest.fn(async () => {}) };
    global.Word = { run: async (fn) => fn(context), RangeLocation: { start: 'Start', end: 'End', content: 'Content' },
        InsertLocation: { before: 'Before', after: 'After', replace: 'Replace' },
        ChangeTrackingMode: { trackAll: 'TrackAll', off: 'Off' } };
    const deps = { appState: { config: { trackChangesEnabled: true, lineDiffEnabled: false } }, log: jest.fn() };
    return { atoms, scope, paragraph, context, writes, deps, bookmarks,
        state: () => revisionTextState(paragraphScope.getOoxml().value, { paragraph: true }) };
}

afterEach(() => { delete global.Word; });

test.each([applyTokenMapStrategy, applyCharDiffStrategy])('pristine paragraph edits use native current text when tracked export padding appears: %p', async (strategy) => {
    const w = world([{ text: 'Kind Regards, ' }], { padding: true, revisedPadding: true, nativeReviewed: true });
    await strategy(w.context, w.scope, 'Kind Regards, ', 'Kind regards,', jest.fn(), { paragraph: true, verificationParagraph: w.paragraph });
    expect(w.atoms.filter((atom) => atom.type !== 'del').map((atom) => atom.text).join('')).toBe('Kind regards,');
    expect(w.scope.insertText).not.toHaveBeenCalled();
});

test.each([
    ['Clear wording.', 'More clear wording.'],
    ['Keep this', 'Keep this!'],
    ['Before', 'After'],
])('paragraph verification reacquires content after boundary edits: %s → %s', async (before, after) => {
    const w = world([{ text: before }], { stickyContent: true, nativeReviewed: true });
    await applyTokenMapStrategy(w.context, w.scope, before, after, jest.fn(), { paragraph: true, verificationParagraph: w.paragraph });
    expect(w.state().text).toBe(after);
});

test.each([applyTokenMapStrategy, applyCharDiffStrategy])('selection verification includes text inserted at either boundary: %p', async (strategy) => {
    const w = world([{ text: 'Keep this' }], { stickyContent: true, nativeReviewed: true });
    await strategy(w.context, w.scope, 'Keep this', 'Please keep this!', jest.fn());
    expect(w.state().text).toBe('Please keep this!');
});

test.each([
    ['Keep this. Keep that.', 'Start here. Keep this. Keep that. Finish here.'],
    ['Short text', 'Replacement text'],
])('sentence and block verification retain boundary insertion ranges: %s', async (before, after) => {
    const w = world([{ text: before }], { stickyContent: true, nativeReviewed: true });
    await applySentenceDiffStrategy(w.context, w.scope, before, after, jest.fn());
    expect(w.state().text).toBe(after);
});

test('successive polish rounds parse revised export padding without inventing a new paragraph', async () => {
    const w = world([{ text: 'Kind Regards, ' }], { padding: true, revisedPadding: true, nativeReviewed: true });
    for (const after of ['Kind regards,', 'Warm regards,']) {
        const document = await parseDocument();
        const before = document.paragraphs[0].text;
        await applyTokenMapStrategy(w.context, w.paragraph.getRange('Content'), before, after, jest.fn(),
            { paragraph: true, verificationParagraph: w.paragraph });
        const next = await parseDocument();
        expect(next.paragraphs).toHaveLength(1);
        expect(next.paragraphs[0].text).toBe(after);
    }
});

test('whole-document reassembly verifies a fresh paragraph after a boundary rewrite', async () => {
    const before = 'the wording is clear.';
    const after = 'Overall, the wording is clear.';
    const w = world([{ text: before }], { stickyContent: true, nativeReviewed: true, padding: true, revisedPadding: true });
    const parsed = await parseDocument();
    const chunk = { id: 'section', startIndex: 0, endIndex: 0, paragraphs: parsed.paragraphs };
    w.scope.insertBookmark('_section');
    const result = await applyChunkResults([{ chunkId: 'section', chunk, status: 'fulfilled', amendment: after }],
        new Map([['section', '_section']]), { trackChangesEnabled: true, log: jest.fn() });
    expect(result).toMatchObject({ amendmentsApplied: 1, appliedParagraphs: 1, errors: [], uncertainChunkIds: [] });
    expect((await parseDocument()).paragraphs[0].text).toBe(after);
});

test.each([applyTokenMapStrategy, applySentenceDiffStrategy, applyCharDiffStrategy])('every strategy edits current text without touching identical old deletions: %p', async (strategy) => {
    const w = world([{ text: 'cat ', type: 'del', id: 'old' }, { text: 'cat cat', type: 'ins', id: 'first-round' }]);
    await strategy(w.context, w.scope, 'cat cat', 'cat dog', jest.fn());
    expect(w.state().text).toBe('cat dog');
    expect(w.atoms.filter((a) => a.id === 'old').map((a) => a.text).join('')).toBe('cat ');
    expect(w.scope.insertText).not.toHaveBeenCalled();
    expect(w.context.document.changeTrackingMode).toBe('TrackMineOnly');
    expect(w.writes.every((entry) => entry.mode === 'TrackAll')).toBe(true);
});

test('three consecutive rounds use the latest proposal baseline without accepting previous changes', async () => {
    const w = world([{ text: 'The slow ', type: 'del', id: 'history' }, { text: 'The quick fox.', type: 'ins' }]);
    for (const next of ['The quick brown fox.', 'A quick brown fox.', 'A quick brown fox!']) {
        const before = w.state().text;
        await applyTokenMapStrategy(w.context, w.scope, before, next, jest.fn());
        expect(w.state().text).toBe(next);
    }
    expect(w.atoms.filter((a) => a.id === 'history').map((a) => a.text).join('')).toBe('The slow ');
    expect(w.scope.insertText).not.toHaveBeenCalled();
});

test('replaces a span split by an existing deletion using disjoint visible targets', async () => {
    const w = world([{ text: 'Hello ' }, { text: 'and ', type: 'del', id: 'history' }, { text: 'world!' }]);
    await tryRevisionDiff(w.context, w.scope, 'Hello world!', 'Hi!', jest.fn());
    expect(w.state().text).toBe('Hi!');
    expect(w.atoms.filter((a) => a.id === 'history').map((a) => a.text).join('')).toBe('and ');
});

test.each([
    ['原来', '我们继续修订这个文档。', '我们继续修订，这个文档。'],
    ['old', 'A 😀 fox', 'A 😃 fox'],
    ['gone', 'word', 'new word'],
    ['gone', 'word', 'word new'],
    ['gone', 'word', ''],
    ['only deleted text', '', 'New text'],
])('handles visible insertion/deletion with prior revisions: %s %s → %s', async (old, before, after) => {
    const w = world([{ text: old, type: 'del', id: 'history' }, { text: before, type: 'ins' }]);
    await tryRevisionDiff(w.context, w.scope, before, after, jest.fn());
    expect(w.state().text).toBe(after);
    expect(w.atoms.filter((a) => a.id === 'history').map((a) => a.text).join('')).toBe(old);
});

test('splits long visible deletion searches below the Word limit', async () => {
    const before = 'ab'.repeat(130);
    const w = world([{ text: 'old', type: 'del' }, { text: before }]);
    await tryRevisionDiff(w.context, w.scope, before, '', jest.fn());
    expect(w.state().text).toBe('');
    expect(w.scope.search.mock.calls.every(([text]) => text.length <= 200)).toBe(true);
});

test('caller-owned tracking is respected, and a no-op does not write', async () => {
    const w = world([{ text: 'old', type: 'del' }, { text: 'new' }]);
    w.context.document.changeTrackingMode = 'Off';
    await tryRevisionDiff(w.context, w.scope, 'new', 'newer', jest.fn(), { trackChanges: false });
    expect(w.context.document.changeTrackingMode).toBe('Off');
    expect(w.writes.every((entry) => entry.mode === 'Off')).toBe(true);
    const count = w.writes.length;
    await tryRevisionDiff(w.context, w.scope, 'newer', 'newer', jest.fn());
    expect(w.writes).toHaveLength(count);
});

test('caller-owned revision edits do not read an unloaded native tracking property', async () => {
    const w = world([{ text: 'Old ', type: 'del' }, { text: 'Current' }]);
    const mode = jest.fn(() => 'TrackAll');
    Object.defineProperty(w.context.document, 'changeTrackingMode', { get: mode });
    await tryRevisionDiff(w.context, w.scope, 'Current', 'Current text', jest.fn(), { trackChanges: false });
    // This mock's write recorder reads the mode once per actual write;
    // the strategy itself must not read a property its caller did not load.
    expect(mode).toHaveBeenCalledTimes(w.writes.length);
});

test('a missing or ambiguous search target fails before any write', async () => {
    const w = world([{ text: 'old', type: 'del' }, { text: 'new' }]);
    w.scope.search.mockImplementation(() => ({ items: [], load: jest.fn() }));
    await expect(applyTokenMapStrategy(w.context, w.scope, 'new', 'now', jest.fn())).rejects.toBeInstanceOf(RevisionSafetyError);
    expect(w.writes).toEqual([]);
    expect(w.scope.insertText).not.toHaveBeenCalled();
});

test('stale text, paragraph-boundary edits and tracking-unavailable hosts fail before writing', async () => {
    const w = world([{ text: 'old', type: 'del' }, { text: 'new' }]);
    await expect(tryRevisionDiff(w.context, w.scope, 'oldnew', 'now')).rejects.toThrow(/baseline/);
    await expect(tryRevisionDiff(w.context, w.scope, 'new', 'new\nparagraph')).rejects.toThrow(/boundaries/);
    delete global.Word.ChangeTrackingMode;
    await expect(tryRevisionDiff(w.context, w.scope, 'new', 'newer')).rejects.toThrow(/enable tracked/);
    expect(w.writes).toEqual([]);
});

test('revision changes during target location invalidate the baseline before writing', async () => {
    const w = world([{ text: 'old', type: 'del', id: 'history' }, { text: 'new' }]);
    w.context.sync.mockImplementation(async () => {
        if (w.context.sync.mock.calls.length === 2) w.atoms[0].id = 'accepted-then-reintroduced';
    });
    await expect(tryRevisionDiff(w.context, w.scope, 'new', 'now')).rejects.toThrow(/state changed/);
    expect(w.writes).toEqual([]);
});

test('read-back failure reports uncertainty and never resets the whole range', async () => {
    const w = world([{ text: 'old', type: 'del' }, { text: 'new' }]);
    w.context.sync.mockImplementation(async () => {
        if (w.writes.length) w.atoms.push({ text: 'corruption', type: 'plain', id: 'corrupt' });
        w.scope._span.end = w.atoms.length;
    });
    await expect(applySentenceDiffStrategy(w.context, w.scope, 'new', 'newer', jest.fn())).rejects.toThrow(/read-back/);
    expect(w.scope.insertText).not.toHaveBeenCalled();
    expect(w.context.document.changeTrackingMode).toBe('TrackMineOnly');
});

test('unreadable XML and host read errors never enter destructive fallbacks', async () => {
    const w = world([{ text: 'old', type: 'del' }, { text: 'new' }]);
    w.scope.getOoxml.mockReturnValue({ value: '<w:p>' });
    await expect(applyTokenMapStrategy(w.context, w.scope, 'new', 'now', jest.fn())).rejects.toThrow(/unreadable/);
    w.scope.getOoxml.mockImplementation(() => { throw new Error('Host failure'); });
    await expect(applyTokenMapStrategy(w.context, w.scope, 'new', 'now', jest.fn())).rejects.toThrow(/Host failure/);
    expect(w.writes).toEqual([]);
});

test('pristine ranges keep their normal strategy', async () => {
    const w = world([{ text: 'Text' }]);
    expect(await tryRevisionDiff(w.context, w.scope, 'Text', 'Updated')).toBeNull();
    expect(await tryRevisionDiff(w.context, { text: 'Text' }, 'Text', 'Updated')).toBeNull();
});

test('pristine ranges verify their baseline before allowing any fallback', async () => {
    const w = world([{ text: 'Actual text.' }]);
    await expect(applyTokenMapStrategy(w.context, w.scope, 'Stale text.', 'Edited text.', jest.fn())).rejects.toThrow(/baseline/);
    expect(w.writes).toEqual([]);
});

test('paragraph padding is normalized in revision content, target prefixes and read-back', async () => {
    const w = world([{ text: 'Earlier ', type: 'del' }, { text: 'Current text.', type: 'ins' }], { padding: true });
    await applyTokenMapStrategy(w.context, w.scope, 'Current text.', 'Current revised text.', jest.fn(), { paragraph: true });
    expect(revisionTextState(w.scope.getOoxml().value, { paragraph: true }).text).toBe('Current revised text.');
});

test.each(['First\nSecond', 'First\n\nSecond'])('native current-view prefixes preserve real paragraph boundaries while excluding export padding: %s', async (before) => {
    const w = world([{ text: before, type: 'ins' }], { padding: true, nativeReviewed: true, nativeParagraphs: true });
    const after = before.replace('Second', 'Better');
    await applyTokenMapStrategy(w.context, w.scope, before, after, jest.fn());
    expect(revisionTextState(w.scope.getOoxml().value).text).toBe(after + '\n');
});

test('unsupported current-view multi-paragraph prefixes fail before writing when XML export padding cannot resolve offsets', async () => {
    const w = world([{ text: 'First\nSecond', type: 'ins' }], { padding: true, nativeParagraphs: true });
    await expect(applyTokenMapStrategy(w.context, w.scope, 'First\nSecond', 'First\nBetter', jest.fn()))
        .rejects.toBeInstanceOf(RevisionSafetyError);
    expect(w.writes).toEqual([]);
});

test('unreadable native current-view prefixes cannot be treated as offset zero', async () => {
    const w = world([{ text: 'Original text.', type: 'ins' }]);
    const getRange = w.scope.getRange;
    w.scope.getRange = (location) => {
        const start = getRange(location);
        const expand = start.expandTo;
        start.expandTo = (other) => {
            const prefix = expand(other);
            prefix.getReviewedText = () => ({ value: undefined });
            return prefix;
        };
        return start;
    };
    await expect(tryRevisionDiff(w.context, w.scope, 'Original text.', 'Updated text.', jest.fn()))
        .rejects.toThrow(/prefix is unreadable/);
    expect(w.writes).toEqual([]);
});

test('read-only token mapping failure tries another locator without a tracked baseline reset', async () => {
    const w = world([{ text: 'Doctoral supervisor.' }]);
    w.context.document.changeTrackingMode = 'TrackAll';
    w.scope.search.mockImplementationOnce(() => ({ items: [], load: jest.fn() }));
    const result = await applyTokenMapStrategy(w.context, w.scope, 'Doctoral supervisor.', 'PhD supervisor.', jest.fn(), { trackChanges: false, paragraph: true });
    expect(result.strategy).toBe('char');
    expect(w.state().text).toBe('PhD supervisor.');
    expect(w.scope.insertText).not.toHaveBeenCalled();
    expect(w.writes.every((entry) => entry.mode === 'TrackAll')).toBe(true);
});

test.each([applyTokenMapStrategy, applyCharDiffStrategy])('a pristine strategy failure after writes never resets or applies another strategy: %p', async (strategy) => {
    const w = world([{ text: 'Original text.' }]);
    w.context.sync.mockImplementation(async () => {
        if (w.writes.length) throw new Error('Host interrupted the write');
    });
    await expect(strategy(w.context, w.scope, 'Original text.', 'Updated text.', jest.fn(), { trackChanges: false }))
        .rejects.toMatchObject({ mutationAttempted: true });
    expect(w.scope.insertText).not.toHaveBeenCalled();
});

test('revision read-back failure carries an explicit possible-mutation outcome', async () => {
    const w = world([{ text: 'Old ', type: 'del' }, { text: 'Current text.' }]);
    w.context.sync.mockImplementation(async () => { if (w.writes.length) throw new Error('Host write failed'); });
    await expect(tryRevisionDiff(w.context, w.scope, 'Current text.', 'Changed text.', jest.fn(), { trackChanges: false }))
        .rejects.toBeInstanceOf(MutationSafetyError);
});

test.each([applyTokenMapStrategy, applySentenceDiffStrategy, applyCharDiffStrategy])('tracking enable failures stop before text writes: %p', async (strategy) => {
    const w = world([{ text: 'Original text.' }]);
    w.context.sync.mockImplementation(async () => {
        if (w.context.document.changeTrackingMode === 'TrackAll') throw new Error('Tracking unavailable');
    });
    await expect(strategy(w.context, w.scope, 'Original text.', 'Updated text.', jest.fn())).rejects.toThrow(/enable tracked changes/);
    expect(w.writes).toEqual([]);
});

test.each([applyTokenMapStrategy, applySentenceDiffStrategy, applyCharDiffStrategy])('tracking-unavailable pristine hosts refuse requested tracked edits: %p', async (strategy) => {
    const w = world([{ text: 'Original text.' }]);
    delete global.Word.ChangeTrackingMode;
    await expect(strategy(w.context, w.scope, 'Original text.', 'Updated text.', jest.fn())).rejects.toThrow(/enable tracked changes/);
    expect(w.writes).toEqual([]);
});

test('structural revisions and protected objects refuse the revision path before writing', async () => {
    const w = world([{ text: 'new', type: 'ins' }]);
    const xml = w.scope.getOoxml().value.replace('<w:r>', '<w:pPr><w:rPr><w:del/></w:rPr></w:pPr><w:r>');
    w.scope.getOoxml.mockReturnValue({ value: xml });
    await expect(applyCharDiffStrategy(w.context, w.scope, 'new', 'newer')).rejects.toThrow(/structural revisions/i);
    expect(w.writes).toEqual([]);
});

test('ambiguous physical search matches and excessive matching are bounded before writing', async () => {
    const w = world([{ text: 'old', type: 'del' }, { text: 'new' }]);
    const search = w.scope.search.getMockImplementation();
    w.scope.search.mockImplementation((text) => {
        const matches = search(text);
        return { ...matches, items: [...matches.items, ...matches.items] };
    });
    await expect(tryRevisionDiff(w.context, w.scope, 'new', 'now')).rejects.toThrow(/Ambiguous/);
    w.scope.search.mockImplementation((text) => {
        const matches = search(text);
        return { ...matches, items: Array(1025).fill(matches.items[0]) };
    });
    await expect(tryRevisionDiff(w.context, w.scope, 'new', 'now')).rejects.toThrow(/Too many/);
    expect(w.writes).toEqual([]);
});

test('dense edits exceeding the target budget stop before host searches', async () => {
    const w = world(Array.from({ length: 257 }, () => ({ text: 'x', type: 'ins' })));
    await expect(tryRevisionDiff(w.context, w.scope, w.state().text, '')).rejects.toThrow(/safe range limit/);
    expect(w.writes).toEqual([]);
    expect(w.scope.search).not.toHaveBeenCalled();
});

test('the parser, selection prompt and live selection preview all read current revision text', async () => {
    const w = world([{ text: 'outdated ', type: 'del' }, { text: 'Current draft.', type: 'ins' }]);
    const parsed = await parseDocument();
    expect(parsed.paragraphs[0]).toMatchObject({ text: 'Current draft.', hasRevisions: true });
    expect(parsed.totalTokens).toBe(Math.ceil('Current draft.'.length / 4));
    expect(await readSelectionText(w.deps)).toMatchObject({ selectionText: 'Current draft.', plainSelectionText: 'Current draft.' });
    expect(await readSelectionSnippet()).toBe('Current draft.');
    expect(await readSelectionContent()).toMatchObject({ text: 'Current draft.' });
});

test('ordinary revision-bearing paragraphs remain editable through anchored document proposals', async () => {
    const w = world([{ text: 'old ', type: 'del', id: 'history' }, { text: 'Current draft.', type: 'ins' }]);
    const snapshot = await readDocumentEditSnapshot();
    expect(snapshot.blocks[0]).toMatchObject({ text: 'Current draft.', hasRevisions: true, readOnly: false });
    const patch = { snapshotId: snapshot.id, changes: [
        { id: 'r', kind: 'replace', blockId: 'p-1', before: 'Current draft.', after: 'Current revised draft.' },
    ] };
    const anchor = await anchorDocumentEdit(snapshot, patch);
    const result = await applyDocumentEdit(w.deps, { patch, anchor });
    expect(result).toMatchObject({ applied: true, verified: true, partial: false });
    expect(w.state().text).toBe('Current revised draft.');
    expect(w.atoms.filter((a) => a.id === 'history').map((a) => a.text).join('')).toBe('old ');
});

test('selection proposals use current text and refuse accepted/rejected revision state drift', async () => {
    const w = world([{ text: 'old ', type: 'del' }, { text: 'Current draft.', type: 'ins' }]);
    const baseline = await readSelectionText(w.deps);
    const proposal = { selectionText: baseline.plainSelectionText, revisionFingerprint: baseline.revisionFingerprint,
        amendedText: 'Current revised draft.' };
    // Accept the insertion without changing final text.
    for (const atom of w.atoms) if (atom.type === 'ins') atom.type = 'plain';
    expect(await applySelectionAmendment(w.deps, proposal)).toMatchObject({ skipped: true });
    expect(w.writes).toEqual([]);
    const fresh = await readSelectionText(w.deps);
    await applySelectionAmendment(w.deps, { ...proposal, revisionFingerprint: fresh.revisionFingerprint });
    expect(w.state().text).toBe('Current revised draft.');
});

test('whole-document reassembly aligns final text and verifies the staged revision fingerprint', async () => {
    const w = world([{ text: 'old ', type: 'del' }, { text: 'Current draft.', type: 'ins' }]);
    const parsed = await parseDocument();
    w.bookmarks.set('chunk', w.scope);
    const result = { chunkId: 'c', status: 'fulfilled', amendment: 'Current revised draft.',
        chunk: { paragraphs: parsed.paragraphs, endIndex: 0 } };
    const options = { trackChangesEnabled: true, lineDiffEnabled: false, log: jest.fn() };
    const applied = await applyChunkResults([result], new Map([['c', 'chunk']]), options);
    expect(applied).toMatchObject({ amendmentsApplied: 1, errors: [] });
    expect(w.state().text).toBe('Current revised draft.');
    const stale = await applyChunkResults([result], new Map([['c', 'chunk']]), options);
    expect(stale.amendmentsApplied).toBe(0);
    expect(stale.errors[0]).toMatch(/no longer matches/);
});

test('a substantial second-round rewrite stays inside the existing paragraph', async () => {
    const w = world([{ text: 'Earlier explanation.', type: 'del', id: 'history' },
        { text: 'The current draft describes one approach.', type: 'ins' }]);
    const parsed = await parseDocument();
    w.bookmarks.set('chunk', w.scope);
    const after = 'Our revised explanation considers a different interpretation.';
    const result = { chunkId: 'c', status: 'fulfilled', amendment: after,
        chunk: { paragraphs: parsed.paragraphs, endIndex: 0 } };
    const applied = await applyChunkResults([result], new Map([['c', 'chunk']]), { log: jest.fn() });
    expect(applied).toMatchObject({ amendmentsApplied: 1, errors: [] });
    expect(w.state().text).toBe(after);
    expect(w.paragraph.delete).not.toHaveBeenCalled();
    expect(w.atoms.filter((a) => a.id === 'history').map((a) => a.text).join('')).toBe('Earlier explanation.');
});
