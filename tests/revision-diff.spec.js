/** @jest-environment jsdom */
import { tryRevisionDiff } from '../src/lib/word-diff/revision-diff.js';
import { applyTokenMapStrategy, applySentenceDiffStrategy, applyCharDiffStrategy } from '../src/lib/word-diff/index.js';
import { revisionTextState, RevisionSafetyError, MutationSafetyError } from '../src/lib/word-revisions.js';
import { parseDocument } from '../src/lib/document-parser.js';
import { readSelectionText, readSelectionContent, readSelectionSnippet, applySelectionAmendment } from '../src/taskpane/word-actions.js';
import { readDocumentEditSnapshot, anchorDocumentEdit, applyDocumentEdit } from '../src/taskpane/document-edit-actions.js';
import { applyChunkResults, bookmarkChunkRanges } from '../src/lib/reassembler.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const escape = (text) => text.replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;');

/** Word searches the physical text, including deleted runs. Writes retain
 * deleted atoms and adjust live range boundaries, unlike flat string mocks. */
function world(parts, { padding = false, nativeReviewed = false, nativeParagraphs = false,
    stickyContent = false, revisedPadding = false, nativeControls = false, rejectEmptyOoxml = false,
    reviewedTransform = (text) => text } = {}) {
    let sequence = 0;
    const atoms = parts.flatMap(({ text, type = 'plain', id = ++sequence, field = false, symbol = null, math = null }) =>
        Array.from(text).map((text, index) => ({ text, type, id, field, symbol, math, mathStart: index === 0 })));
    const ranges = [];
    const writes = [];
    const bookmarks = new Map();
    let xmlReads = 0;
    const xml = (slice) => {
        const groups = [[]];
        for (const atom of slice) {
            if (nativeParagraphs && atom.text === '\n') groups.push([]);
            else groups.at(-1).push(atom);
        }
        const content = groups.map((group) => `<w:p xmlns:w="${W}">` + group.map((atom) => {
        if (atom.math) return atom.mathStart ? atom.math.xml : '';
        const run = atom.symbol ? `<w:r><w:sym w:font="Symbol" w:char="${atom.symbol}"/></w:r>`
            : atom.text === '‑' && nativeControls ? '<w:r><w:noBreakHyphen/></w:r>'
                : `<w:r><w:${atom.type === 'del' ? 'delText' : 't'}>${escape(atom.text)}</w:${atom.type === 'del' ? 'delText' : 't'}></w:r>`;
        const visible = atom.type === 'plain' ? run : `<w:${atom.type} w:id="${atom.id}" w:author="Reviewer">${run}</w:${atom.type}>`;
        return atom.field ? `<w:fldSimple w:instr="HYPERLINK example">${visible}</w:fldSimple>` : visible;
        }).join('') + '</w:p>').join('');
        const end = revisedPadding && slice.some((atom) => atom.type !== 'plain')
            ? '<w:p><w:pPr><w:rPr><w:ins w:id="999"/></w:rPr></w:pPr></w:p>' : '<w:p/>';
        return padding || nativeParagraphs ? `<w:document xmlns:w="${W}"><w:body>${content}${padding ? end : ''}</w:body></w:document>` : content;
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
            getOoxml: jest.fn(() => {
                if (rejectEmptyOoxml && span.start === span.end) {
                    throw Object.assign(new Error('GeneralException'), { code: 'GeneralException', debugInfo: { errorLocation: 'Range.getOoxml' } });
                }
                xmlReads++; return { value: xml(atoms.slice(span.start, span.end)) };
            }),
            ...(nativeReviewed ? { getReviewedText: jest.fn(() => ({ value: reviewedTransform(
                atoms.slice(span.start, span.end).filter((atom) => atom.type !== 'del').map((atom) => {
                    if (atom.math) return atom.mathStart ? atom.math.native : '';
                    const text = nativeControls && atom.symbol ? '(' : atom.text;
                    return nativeControls && atom.field ? '\u0013HYPERLINK example\u0014' + text + '\u0015' : text;
                }).join('').replace(/\n/g, '\r').replace(/‑/g, nativeControls ? '\u001E' : '‑')) })) } : {}),
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
                const sentences = nativeParagraphs ? /[^\n]+?(?:\. {1,2}|\n|$)|\n/g : /.+?(?:\. {1,2}|$)/gs;
                return collection(Array.from(raw.matchAll(sentences), (match) =>
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
        xmlReadCount: () => xmlReads,
        state: () => revisionTextState(paragraphScope.getOoxml().value, { paragraph: true, singleParagraphContent: true }) };
}

afterEach(() => { delete global.Word; });

const opaqueMath = { xml: '<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">'
    + '<m:sSub><m:e><m:r><m:t>x</m:t></m:r></m:e><m:sub><m:r><m:t>j</m:t></m:r></m:sub></m:sSub></m:oMath>',
native: ',𝑗𝑥' };

test('two tracked polish rounds verify all surrounding prose when whole-paragraph reviewed math cannot be projected', async () => {
    const w = world([{ text: 'Older text', type: 'del' }, { text: 'The old ' }, { text: '𝑥_(𝑗)', math: opaqueMath },
        { text: ' result. Keep the result.' }], { nativeReviewed: true, padding: true, revisedPadding: true, rejectEmptyOoxml: true });
    for (const after of ['The revised 𝑥_(𝑗) result. Keep the result.', 'The final 𝑥_(𝑗) result. Keep this result.']) {
        const before = w.state().text;
        await tryRevisionDiff(w.context, w.scope, before, after, jest.fn(), { paragraph: true, nativeTargets: true, preflightOnly: true });
        await tryRevisionDiff(w.context, w.scope, before, after, jest.fn(), { paragraph: true, nativeTargets: true });
        expect(w.state().text).toBe(after);
    }
    expect(w.atoms.filter((a) => a.math).every((a) => a.type === 'plain')).toBe(true);
    expect(w.scope.insertText).not.toHaveBeenCalled();
});

test('Mac reviewed equation punctuation locates prose without exporting collapsed position prefixes', async () => {
    const math = { ...opaqueMath, native: ',𝑥-𝑗.' };
    const w = world([{ text: 'Earlier text. ', type: 'del' }, { text: 'The old ' }, { text: '𝑥_(𝑗)', math },
        { text: ' result. Keep this result.' }],
    { nativeReviewed: true, padding: true, revisedPadding: true, rejectEmptyOoxml: true });
    const log = jest.fn();
    await tryRevisionDiff(w.context, w.scope, w.state().text, 'The revised 𝑥_(𝑗) result. Keep this result.', log,
        { paragraph: true, nativeTargets: true });
    expect(w.state().text).toBe('The revised 𝑥_(𝑗) result. Keep this result.');
    expect(log).not.toHaveBeenCalledWith(expect.stringContaining('OOXML positions'), expect.anything());
    expect(w.atoms.filter((a) => a.math).every((a) => a.type === 'plain')).toBe(true);
});

test('a Word preflight read failure identifies its stage and never attempts a write', async () => {
    const w = world([{ text: 'The old result.' }], { nativeReviewed: true });
    const error = Object.assign(new Error('GeneralException'), { code: 'GeneralException',
        debugInfo: { errorLocation: 'Range.getReviewedText' } });
    w.context.sync.mockResolvedValueOnce().mockResolvedValueOnce().mockRejectedValueOnce(error);
    await expect(tryRevisionDiff(w.context, w.scope, 'The old result.', 'The revised result.', jest.fn(),
        { paragraph: true, nativeTargets: true })).rejects.toThrow(/current-view search ranges.*Range.getReviewedText/);
    expect(w.writes).toHaveLength(0);
});

test('unverifiable unchanged prose blocks the native-object path before the first write', async () => {
    const w = world([{ text: 'The old ' }, { text: '𝑥_(𝑗)', math: opaqueMath }, { text: ' result. Keep this unchanged.' }], { nativeReviewed: true, reviewedTransform: (text) => text.replace('Keep', 'Fake') });
    await expect(tryRevisionDiff(w.context, w.scope, w.state().text, 'The revised 𝑥_(𝑗) result. Keep this unchanged.', jest.fn(),
        { paragraph: true, nativeTargets: true })).rejects.toThrow(/Unverified|crosses/);
    expect(w.writes).toHaveLength(0);
});

test('post-write changes to equation structure cause an uncertain outcome even if visible text stays the same', async () => {
    const math = { ...opaqueMath };
    const w = world([{ text: 'The old ' }, { text: '𝑥_(𝑗)', math }, { text: ' result.' }], { nativeReviewed: true });
    w.context.sync.mockImplementation(async () => {
        if (w.writes.length) math.xml = opaqueMath.xml.replace('<m:sSub>', '<m:sSub><m:sSubPr><m:ctrlPr/></m:sSubPr>');
    });
    await expect(tryRevisionDiff(w.context, w.scope, w.state().text, 'The new 𝑥_(𝑗) result.', jest.fn(),
        { paragraph: true, nativeTargets: true })).rejects.toMatchObject({ mutationAttempted: true });
});

test.each([[false, false], [false, true], [true, false]])('selection edits keep headings, blank paragraphs and repeated sentences in place, sentence mode=%s, missing searches=%s', async (lineDiffEnabled, missingSearches) => {
    const before = '3.1 Summary\n\nThe first result is clear. The next result is preliminary.\nThe first result is clear. Further tests are needed.';
    const after = '3.1 Summary\n\nThe first result is robust. The next result is preliminary.\nThe first result is clear. Further tests remain necessary.';
    const w = world([{ text: before }], { nativeReviewed: true, nativeParagraphs: true });
    w.deps.appState.config.lineDiffEnabled = lineDiffEnabled;
    // Native searches do not locate a literal LF across Word paragraphs.
    const search = w.scope.search.getMockImplementation();
    w.scope.search.mockImplementation((text) => missingSearches || text.includes('\n') ? { items: [], load: jest.fn() } : search(text));
    const baseline = await readSelectionText(w.deps);
    await applySelectionAmendment(w.deps, { selectionText: baseline.plainSelectionText,
        revisionFingerprint: baseline.revisionFingerprint, amendedText: after });
    expect(w.state().text).toBe(after);
    expect(w.scope.insertText).not.toHaveBeenCalled();
    if (missingSearches) expect(w.deps.log).toHaveBeenCalledWith(expect.stringContaining('trying sentence diff'), 'warning');
});

test('sentence editing a partial selection preserves text outside both boundaries', async () => {
    const prefix = 'Unselected beginning. ';
    const before = 'Selected ending.\nSelected start. Keep this sentence.';
    const after = 'Revised ending.\nRevised start. Keep this sentence.';
    const suffix = ' Unselected ending.';
    const w = world([{ text: prefix + before + suffix }], { nativeReviewed: true, nativeParagraphs: true });
    w.scope._span.start = prefix.length;
    w.scope._span.end -= suffix.length;
    w.deps.appState.config.lineDiffEnabled = true;
    await applySelectionAmendment(w.deps, { selectionText: before, amendedText: after });
    expect(w.atoms.filter((atom) => atom.type !== 'del').map((atom) => atom.text).join('')).toBe(prefix + after + suffix);
    expect(w.scope.insertText).not.toHaveBeenCalled();
});

test.each([['甲\u00a0乙 丙', '甲\u00a0乙，丙', ' '],
    ['An old‑form and old-form.', 'An old‑form and new-form.', 'old']])(
    'native location checks handle approximate whitespace matches and repeated text: %s', async (before, after, needle) => {
        const w = world([{ text: before }], { nativeReviewed: true });
        if (needle === ' ') {
            const search = w.scope.search.getMockImplementation();
            w.scope.search.mockImplementation((text) => {
                const matches = search(text);
                return text === ' ' ? { ...matches, items: [...search('\u00a0').items, ...matches.items] } : matches;
            });
        }
        await applyCharDiffStrategy(w.context, w.scope, before, after, jest.fn(),
            { paragraph: true, nativeTargets: true });
        expect(w.state().text).toBe(after);
    });

test.each([false, true])('native field results survive surrounding prose edits, existing revisions=%s', async (revised) => {
    const w = world([{ text: 'An old ', type: revised ? 'ins' : 'plain' }, { text: '(Example 2025)', field: true }, { text: ' result.' }],
        { nativeReviewed: true });
    const before = w.state().text;
    await applyTokenMapStrategy(w.context, w.scope, before, 'A revised (Example 2025) result.', jest.fn(),
        { paragraph: true, nativeTargets: true });
    expect(w.state().text).toBe('A revised (Example 2025) result.');
    expect(w.atoms.filter((atom) => atom.field).every((atom) => atom.type === 'plain')).toBe(true);
});

test('a pristine native field mutation fails before any write, including before surrounding edits', async () => {
    const w = world([{ text: 'An old ' }, { text: '(Example 2025)', field: true }, { text: ' result.' }], { nativeReviewed: true });
    await expect(applyCharDiffStrategy(w.context, w.scope, w.state().text, 'A new (Example 2026) result.', jest.fn(),
        { paragraph: true, nativeTargets: true })).rejects.toThrow(/native field result/);
    expect(w.writes).toHaveLength(0);
});

test.each([false, true])('raw native field codes preserve surrounding edit offsets and existing revisions: %s', async (revised) => {
    const w = world([{ text: 'result before. ', type: revised ? 'ins' : 'plain' },
        { text: 'Table 4', field: true }, { text: ' result after. result after.' }],
    { nativeReviewed: true, nativeControls: true });
    const before = w.state().text;
    const after = 'result before. Table 4 result after. revised result after.';
    await tryRevisionDiff(w.context, w.scope, before, after, jest.fn(),
        { paragraph: true, nativeTargets: true, preflightOnly: true });
    expect(w.writes).toHaveLength(0);
    await applyTokenMapStrategy(w.context, w.scope, before, after, jest.fn(),
        { paragraph: true, nativeTargets: true });
    expect(w.state().text).toBe(after);
    expect(w.atoms.filter((atom) => atom.field).every((atom) => atom.type === 'plain')).toBe(true);
    expect(w.scope.insertText).not.toHaveBeenCalled();
});

test('whole-document reassembly preserves Symbol glyphs and nonbreaking hyphens across polish rounds', async () => {
    const w = world([{ text: 'For pour‑type (20 ' }, { text: '°', symbol: 'F0B0' },
        { text: 'C), allow an additional 10 s. See ' }, { text: 'Table 4', field: true }, { text: ' for details.' }],
    { nativeReviewed: true, nativeControls: true, padding: true, revisedPadding: true });
    for (const after of [
        'For pour‑type (20 °C), allow an additional 10 s. See Table 4 for more details.',
        'For pour‑type (20 °C), allow an additional 10 s. See Table 4 for further details.',
    ]) {
        const parsed = await parseDocument();
        expect(parsed.paragraphs[0].text).toContain('(20 °C)');
        w.bookmarks.set('section', w.scope);
        const applied = await applyChunkResults([{ chunkId: 'c', status: 'fulfilled', amendment: after,
            chunk: { paragraphs: parsed.paragraphs, endIndex: 0 } }], new Map([['c', 'section']]), { log: jest.fn() });
        expect(applied).toMatchObject({ amendmentsApplied: 1, errors: [], uncertainChunkIds: [] });
        expect((await parseDocument()).paragraphs[0].text).toBe(after);
    }
    expect(w.atoms.filter((atom) => atom.symbol || atom.field).every((atom) => atom.type === 'plain')).toBe(true);
    expect(w.scope.insertText).not.toHaveBeenCalled();
});

test('two polish rounds edit prose around an intact equation, raw hyperlink fields and earlier revisions', async () => {
    const math = { xml: '<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">'
        + '<m:sSup><m:e><m:r><m:t>x</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup></m:oMath>',
    native: '\r𝑥\r2\r' };
    const w = world([{ text: 'Earlier result. ', type: 'del', id: 'history' }, { text: 'The old ' },
        { text: '𝑥^(2)', math }, { text: ' result refers to ' }, { text: 'dataset', field: true }, { text: ' for details.' }],
    { nativeReviewed: true, nativeControls: true, padding: true, revisedPadding: true });
    for (const after of ['The revised 𝑥^(2) result refers to dataset for more details.',
        'The final 𝑥^(2) result refers to dataset for further details.']) {
        const parsed = await parseDocument();
        expect(parsed.paragraphs[0].text).toContain('𝑥^(2)');
        expect(parsed.paragraphs[0].text).not.toContain('\n');
        w.bookmarks.set('section', w.scope);
        const applied = await applyChunkResults([{ chunkId: 'c', status: 'fulfilled', amendment: after,
            chunk: { paragraphs: parsed.paragraphs, endIndex: 0 } }], new Map([['c', 'section']]), { log: jest.fn() });
        expect(applied).toMatchObject({ amendmentsApplied: 1, errors: [], uncertainChunkIds: [] });
        expect((await parseDocument()).paragraphs[0].text).toBe(after);
    }
    expect(w.atoms.filter((atom) => atom.math || atom.field).every((atom) => atom.type === 'plain')).toBe(true);
    expect(w.atoms.filter((atom) => atom.id === 'history').map((atom) => atom.text).join('')).toBe('Earlier result. ');
    expect(w.scope.insertText).not.toHaveBeenCalled();
});

test('an equation mutation fails paragraph preflight before surrounding prose is written', async () => {
    const math = { xml: '<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:r><m:t>k</m:t></m:r></m:oMath>', native: '𝑘' };
    const w = world([{ text: 'The old ' }, { text: '𝑘', math }, { text: ' result.' }], { nativeReviewed: true });
    await expect(applyTokenMapStrategy(w.context, w.scope, w.state().text, 'The new q result.', jest.fn(),
        { paragraph: true, nativeTargets: true })).rejects.toThrow(/equation/);
    expect(w.writes).toHaveLength(0);
});

test.each(['Temperature is 20 C.', 'Temperature is 20 (C.'])('symbol deletion or replacement is rejected before any surrounding edit: %s', async (after) => {
    const w = world([{ text: 'Temperature is 20 ' }, { text: '°', symbol: 'F0B0' }, { text: 'C.' }],
        { nativeReviewed: true, nativeControls: true });
    await expect(applyCharDiffStrategy(w.context, w.scope, w.state().text, after, jest.fn(),
        { paragraph: true, nativeTargets: true })).rejects.toThrow(/native field result/);
    expect(w.writes).toHaveLength(0);
});

test('pristine native text disagrees with XML before writing rather than failing only after mutation', async () => {
    const w = world([{ text: 'Original text.' }], { nativeReviewed: true });
    w.scope.getReviewedText.mockReturnValue({ value: 'Original\u00a0text.' });
    await expect(applyCharDiffStrategy(w.context, w.scope, 'Original text.', 'Revised text.', jest.fn(),
        { paragraph: true, nativeTargets: true })).rejects.toThrow(/before writing.*U\+0020.*U\+00A0/);
    expect(w.writes).toHaveLength(0);
});

test('many repeated native search matches do not export every candidate and prefix as full OOXML packages', async () => {
    const before = 'A '.repeat(120) + 'End.';
    const after = before.slice(0, 120) + ',' + before.slice(120);
    const w = world([{ text: before }], { nativeReviewed: true });
    await applyCharDiffStrategy(w.context, w.scope, before, after, jest.fn(), { paragraph: true, nativeTargets: true });
    expect(w.xmlReadCount()).toBeLessThanOrEqual(5);
    expect(w.state().text).toBe(after);
});

test('bookmark staging compares final-view boundary text when earlier revisions remain unresolved', async () => {
    const w = world([{ text: 'Earlier words ', type: 'del' }, { text: 'Current paragraph.', type: 'ins' }], { nativeReviewed: true });
    const parsed = await parseDocument();
    const chunk = { id: 'section', startIndex: 0, endIndex: 0, paragraphs: parsed.paragraphs };
    const bookmarks = await bookmarkChunkRanges([chunk]);
    expect(bookmarks.has('section')).toBe(true);
    expect(w.writes).toHaveLength(0);
});

test('full-document prose edits preserve manual blank lines at native paragraph boundaries', async () => {
    const before = '\nOriginal line.\nSecond line.\n';
    const after = '\nRevised line.\nSecond revised line.\n';
    const w = world([{ text: before }], { nativeReviewed: true });
    const parsed = await parseDocument();
    w.bookmarks.set('_section', w.scope);
    const result = await applyChunkResults([{ chunkId: 'section', status: 'fulfilled', amendment: after,
        chunk: { id: 'section', paragraphs: parsed.paragraphs } }], new Map([['section', '_section']]),
    { trackChangesEnabled: true, log: jest.fn() });
    expect(result.errors).toEqual([]);
    expect(result.amendmentsApplied).toBe(1);
    expect(w.state().text).toBe(after);
    expect(w.paragraph.delete).not.toHaveBeenCalled();
});

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

test('preflight locates revised text without writes and the apply pass still verifies it', async () => {
    const w = world([{ text: 'Hello ' }, { text: 'and ', type: 'del', id: 'history' }, { text: 'world!' }]);
    const baseline = w.state().text;
    await tryRevisionDiff(w.context, w.scope, baseline, 'Hi!', jest.fn(), { preflightOnly: true, trackChanges: false });
    expect(w.state().text).toBe(baseline);
    expect(w.writes).toHaveLength(0);
    expect(w.context.document.changeTrackingMode).toBe('TrackMineOnly');
    await tryRevisionDiff(w.context, w.scope, baseline, 'Hi!', jest.fn());
    expect(w.state().text).toBe('Hi!');
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
