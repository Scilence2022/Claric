/** @jest-environment jsdom */
jest.mock('../src/lib/word-diff/index.js', () => ({
    applyTokenMapStrategy: jest.fn(async (_context, range, _before, after) => { range.paragraph.text = after; }),
    applySentenceDiffStrategy: jest.fn(async (_context, range, _before, after) => { range.paragraph.text = after; }),
    applyCharDiffStrategy: jest.fn(async (_context, range, _before, after) => { range.paragraph.text = after; }),
    hasCjk: (text) => /[\u4e00-\u9fff]/.test(text),
}));
jest.mock('../src/lib/document-edit-session.js', () => ({ runDocumentEditSession: jest.fn() }));
jest.mock('../src/lib/llm-client.js', () => ({ sendMessages: jest.fn() }));
jest.mock('../src/lib/file-store.js', () => ({ getFile: jest.fn(), readFile: jest.fn(), listFiles: jest.fn() }));

import { readDocumentEditSnapshot, anchorDocumentEdit, prepareDocumentEdit, applyDocumentEdit, discardDocumentEdit } from '../src/taskpane/document-edit-actions.js';
import { runDocumentEditSession } from '../src/lib/document-edit-session.js';
import { applyTokenMapStrategy, applySentenceDiffStrategy, applyCharDiffStrategy } from '../src/lib/word-diff/index.js';
import { getFile, readFile } from '../src/lib/file-store.js';
import { sendMessages } from '../src/lib/llm-client.js';

function world(texts = ['Discussion', 'Mechanism.', 'Limitations.', 'Untouched.']) {
    const bookmarks = new Map();
    const mutations = [];
    let id = 0;
    const w = { paragraphs: [], bookmarks, mutations, corruptReadback: false,
        maxOoxmlBatch: Infinity, failRangeOoxml: new Set(), failParagraphOoxml: new Set(), pendingOoxml: 0,
        failedBatchCount: 0 };
    const encode = (text) => text.replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;');
    const xml = (p) => `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:pPr><w:pStyle w:val="${p.style}"/></w:pPr><w:r><w:t>${encode(w.corruptReadback && p.added ? 'Incorrect text' : p.text)}</w:t></w:r>${p.extraXml || ''}</w:p>`;
    const collection = (items) => ({ get items() { return typeof items === 'function' ? items() : items; }, load: jest.fn(),
        getFirst() { return this.items[0]; }, getLast() { return this.items[this.items.length - 1]; } });
    function range(p) {
        return {
            paragraph: p, get text() { return p.text; }, load: jest.fn(), isNullObject: false,
            getOoxml: jest.fn(() => {
                if (w.failRangeOoxml.has(p.key)) throw new Error('GeneralException');
                w.pendingOoxml++;
                return { value: p.xmlOverride ?? xml(p) };
            }),
            compareLocationWith: (other) => ({ value: other.paragraph === p ? 'Equal' : 'Before' }),
            insertBookmark: jest.fn((name) => bookmarks.set(name, p)),
            paragraphs: collection([p]),
            expandTo: (other) => ({ paragraphs: collection(() => w.paragraphs.slice(w.paragraphs.indexOf(p), w.paragraphs.indexOf(other.paragraph) + 1)) }),
        };
    }
    function paragraph(text, added = false) {
        const p = { key: ++id, text, added, style: 'Normal', styleBuiltIn: 'Normal', isListItem: false, load: jest.fn(),
            font: { bold: false, italic: false, load: jest.fn() },
            parentTableOrNullObject: { isNullObject: true, load: jest.fn() },
            getRange: () => range(p),
            getOoxml: jest.fn(() => {
                if (w.failParagraphOoxml.has(p.key)) throw new Error('GeneralException');
                w.pendingOoxml++;
                return { value: p.xmlOverride ?? xml(p) };
            }),
            insertParagraph: jest.fn((value, location) => {
                const added = paragraph(value, true);
                mutations.push({ value, location, anchor: p.key, mode: w.document.changeTrackingMode });
                w.paragraphs.splice(w.paragraphs.indexOf(p) + (location === 'After' ? 1 : 0), 0, added);
                w.onInsert?.(added);
                return added;
            }),
        };
        return p;
    }
    w.paragraphs = texts.map((t) => paragraph(t));
    w.paragraphs[0].style = 'Heading 1';
    w.paragraphs[0].styleBuiltIn = 'Heading1';
    w.addParagraph = (text, index) => w.paragraphs.splice(index, 0, paragraph(text));
    w.document = { body: { paragraphs: collection(() => w.paragraphs) }, changeTrackingMode: 'TrackMineOnly', load: jest.fn(),
        getBookmarkRangeOrNullObject: (name) => bookmarks.has(name) ? range(bookmarks.get(name)) : { isNullObject: true, load: jest.fn() },
        deleteBookmark: jest.fn((name) => bookmarks.delete(name)),
    };
    w.context = { document: w.document, sync: jest.fn(async () => {
        const count = w.pendingOoxml;
        w.pendingOoxml = 0;
        if (count > w.maxOoxmlBatch) {
            w.failedBatchCount++;
            throw Object.assign(new Error('GeneralException'), { code: 'GeneralException' });
        }
    }) };
    global.Word = { run: async (fn) => { w.pendingOoxml = 0; return fn(w.context); }, RangeLocation: { content: 'Content' },
        InsertLocation: { before: 'Before', after: 'After' }, BuiltInStyleName: { normal: 'Normal' },
        ChangeTrackingMode: { trackAll: 'TrackAll', off: 'Off' }, LocationRelation: { equal: 'Equal' } };
    w.deps = { appState: { config: { trackChangesEnabled: true, backend: 'ollama', providers: { ollama: { model: 'test', url: 'http://localhost' } } } }, log: jest.fn() };
    return w;
}

async function prepared(w, changes = undefined) {
    const snapshot = await readDocumentEditSnapshot();
    const patch = { snapshotId: snapshot.id, changes: changes || [
        { id: 'insert-1', kind: 'insert', afterId: 'p-2', beforeId: 'p-3', paragraphs: ['First new paragraph.', 'Second new paragraph.'] },
        { id: 'replace-1', kind: 'replace', blockId: 'p-3', before: 'Limitations.', after: 'Nevertheless, limitations remain.' },
    ] };
    const anchor = await anchorDocumentEdit(snapshot, patch);
    return { snapshot, patch, anchor };
}

beforeEach(() => jest.clearAllMocks());
afterEach(() => { delete global.Word; });

test('captures addressable blocks, protects structure and does not alter document text while staging', async () => {
    const w = world();
    w.paragraphs[1].isListItem = true;
    w.paragraphs[2].extraXml = '<w:drawing/>';
    w.paragraphs[3].parentTableOrNullObject.isNullObject = false;
    const snapshot = await readDocumentEditSnapshot();
    expect(snapshot.blocks[0]).toMatchObject({ id: 'p-1', index: 0, headingLevel: 1, section: 'Discussion', readOnly: true });
    expect(snapshot.blocks.slice(1).every((b) => b.readOnly)).toBe(true);
    expect(w.mutations).toEqual([]);
});

test('retries a Word XML batch failure in bounded single-paragraph reads', async () => {
    const w = world();
    w.maxOoxmlBatch = 1;
    const snapshot = await readDocumentEditSnapshot();
    expect(snapshot.blocks.map((block) => block.text)).toEqual(['Discussion', 'Mechanism.', 'Limitations.', 'Untouched.']);
    expect(snapshot.blocks.every((block) => block.structureUnavailable === false)).toBe(true);
    expect(w.failedBatchCount).toBe(1);
});

test('one Mac Word XML failure keeps only that paragraph read-only', async () => {
    const w = world();
    const warning = jest.fn();
    w.failRangeOoxml.add(w.paragraphs[1].key);
    w.failParagraphOoxml.add(w.paragraphs[1].key);
    const snapshot = await readDocumentEditSnapshot({ onWarning: warning });
    expect(snapshot.blocks[1]).toMatchObject({ text: 'Mechanism.', ooxml: null,
        structureUnavailable: true, readOnly: true });
    expect(snapshot.blocks[2]).toMatchObject({ structureUnavailable: false, readOnly: false });
    expect(warning).toHaveBeenCalledWith(expect.stringContaining('1 paragraph(s)'));
    await expect(anchorDocumentEdit(snapshot, { snapshotId: snapshot.id, changes: [{ id: 'i', kind: 'insert',
        afterId: 'p-2', beforeId: 'p-3', paragraphs: ['New.'] }] })).rejects.toThrow(/could not be read/);
    expect(w.bookmarks.size).toBe(0);
});

test('a paragraph-level XML fallback preserves context but never permits an unsafe range anchor', async () => {
    const w = world();
    w.failRangeOoxml.add(w.paragraphs[1].key);
    const snapshot = await readDocumentEditSnapshot();
    expect(snapshot.blocks[1]).toMatchObject({ text: 'Mechanism.', structureUnavailable: true, readOnly: true });
    expect(snapshot.blocks[1].ooxml).toContain('Mechanism.');
    expect(snapshot.blocks[2].structureUnavailable).toBe(false);
});

test('a host with no readable paragraph XML reports the failing stage', async () => {
    const w = world();
    for (const paragraph of w.paragraphs) {
        w.failRangeOoxml.add(paragraph.key);
        w.failParagraphOoxml.add(paragraph.key);
    }
    await expect(readDocumentEditSnapshot()).rejects.toThrow(/Word could not safely anchor any paragraph: GeneralException/);
    expect(w.bookmarks.size).toBe(0);
});

test('Word host errors identify the snapshot stage and location', async () => {
    const w = world();
    w.context.sync.mockRejectedValueOnce(Object.assign(new Error('GeneralException'), {
        code: 'GeneralException', debugInfo: { errorLocation: 'ParagraphCollection.load' },
    }));
    await expect(readDocumentEditSnapshot()).rejects.toThrow(
        'Word could not enumerate paragraphs: GeneralException (ParagraphCollection.load)'
    );
    expect(w.bookmarks.size).toBe(0);
});

test('applies only planned paragraphs at the verified gap, reads them back and restores tracking', async () => {
    const w = world();
    const proposal = await prepared(w);
    expect(w.mutations).toEqual([]);
    const result = await applyDocumentEdit(w.deps, proposal);
    expect(result).toMatchObject({ applied: true, verified: true, partial: false, appliedOperationIds: ['replace-1', 'insert-1'] });
    expect(w.paragraphs.map((p) => p.text)).toEqual(['Discussion', 'Mechanism.', 'First new paragraph.', 'Second new paragraph.', 'Nevertheless, limitations remain.', 'Untouched.']);
    expect(w.mutations.every((m) => m.mode === 'TrackAll')).toBe(true);
    expect(w.document.changeTrackingMode).toBe('TrackMineOnly');
    expect(applyTokenMapStrategy).toHaveBeenCalledTimes(1);
    await expect(applyDocumentEdit(w.deps, proposal)).rejects.toThrow(/already attempted/);
    await discardDocumentEdit(w.deps, proposal);
    await discardDocumentEdit(w.deps, proposal);
    expect(w.bookmarks.size).toBe(0);
});

test('inserts and verifies bold and italic on new paragraphs in one proposal', async () => {
    const w = world();
    const proposal = await prepared(w, [{ id: 'i', kind: 'insert', afterId: 'p-2', beforeId: 'p-3',
        paragraphs: ['First.', 'Second.'], paragraphFormats: [{ bold: true }, { italic: true }] }]);
    const result = await applyDocumentEdit(w.deps, proposal);
    expect(result).toMatchObject({ applied: true, verified: true, partial: false });
    expect(w.paragraphs[2].font.bold).toBe(true);
    expect(w.paragraphs[3].font.italic).toBe(true);
    expect(w.paragraphs[2].font.load).toHaveBeenCalledWith('bold');
});

test('rejects unsupported paragraph formatting before creating anchors', async () => {
    world();
    const snapshot = await readDocumentEditSnapshot();
    await expect(anchorDocumentEdit(snapshot, { snapshotId: snapshot.id, changes: [{ id: 'i', kind: 'insert',
        afterId: 'p-2', beforeId: 'p-3', paragraphs: ['New.'], paragraphFormats: [{ color: '#ff0000' }] }] }))
        .rejects.toThrow(/format/);
});

test.each(['text', 'style', 'missing', 'gap'])('rejects %s drift before any write', async (kind) => {
    const w = world();
    const proposal = await prepared(w);
    if (kind === 'text') w.paragraphs[1].text = 'User edit';
    if (kind === 'style') w.paragraphs[1].style = 'Quote';
    if (kind === 'missing') w.bookmarks.clear();
    if (kind === 'gap') w.addParagraph('User paragraph', 2);
    await expect(applyDocumentEdit(w.deps, proposal)).rejects.toThrow(/changed|disappeared/);
    expect(w.mutations).toHaveLength(0);
    expect(applyTokenMapStrategy).not.toHaveBeenCalled();
    expect(proposal.anchor.attempted).toBe(false);
});

test('bookmarks keep repeated text identities after an unrelated earlier insertion', async () => {
    const w = world(['Same.', 'Mechanism.', 'Same.', 'Untouched.']);
    const proposal = await prepared(w, [{ id: 'r', kind: 'replace', blockId: 'p-3', before: 'Same.', after: 'Only this occurrence.' }]);
    w.addParagraph('Unrelated.', 0);
    const result = await applyDocumentEdit(w.deps, proposal);
    expect(result.verified).toBe(true);
    expect(w.paragraphs.map((p) => p.text)).toEqual(['Unrelated.', 'Same.', 'Mechanism.', 'Only this occurrence.', 'Untouched.']);
});

test.each(['count', 'text', 'xml'])('draft-time %s drift rejects anchoring without leaving bookmarks', async (kind) => {
    const w = world();
    const snapshot = await readDocumentEditSnapshot();
    if (kind === 'count') w.addParagraph('New.', 0);
    if (kind === 'text') w.paragraphs[0].text = 'Changed';
    if (kind === 'xml') w.paragraphs[1].extraXml = '<w:r><w:rPr><w:b/></w:rPr></w:r>';
    await expect(anchorDocumentEdit(snapshot, { snapshotId: snapshot.id, changes: [{ id: 'i', kind: 'insert', afterId: 'p-2', beforeId: 'p-3', paragraphs: ['New.'] }] })).rejects.toThrow(/changed/);
    expect(w.bookmarks.size).toBe(0);
});

test('Word proofing markup appearing during drafting does not invalidate an unchanged paragraph', async () => {
    const w = world();
    const snapshot = await readDocumentEditSnapshot();
    w.paragraphs[1].extraXml = '<w:proofErr w:type="spellStart"/>';
    const anchor = await anchorDocumentEdit(snapshot, { snapshotId: snapshot.id, changes: [{
        id: 'i', kind: 'insert', afterId: 'p-2', beforeId: 'p-3', paragraphs: ['New.'],
    }] });
    expect(Object.keys(anchor.anchors)).toEqual(['p-2', 'p-3']);
    await discardDocumentEdit(w.deps, { anchor });
});

test('Word proofing markup appearing after anchoring does not block a verified insert', async () => {
    const w = world();
    const proposal = await prepared(w, [{ id: 'i', kind: 'insert', afterId: 'p-2', beforeId: 'p-3', paragraphs: ['New.'] }]);
    w.paragraphs[1].extraXml = '<w:proofErr w:type="spellStart"/>';
    const result = await applyDocumentEdit(w.deps, proposal);
    expect(result).toMatchObject({ applied: true, verified: true, partial: false });
});

test('a real paragraph formatting change after anchoring still blocks the write', async () => {
    const w = world();
    const proposal = await prepared(w, [{ id: 'i', kind: 'insert', afterId: 'p-2', beforeId: 'p-3', paragraphs: ['New.'] }]);
    w.paragraphs[1].extraXml = '<w:r><w:rPr><w:b/></w:rPr></w:r>';
    await expect(applyDocumentEdit(w.deps, proposal)).rejects.toThrow(/target text, formatting or insertion gap changed/);
    expect(w.mutations).toHaveLength(0);
});

test('cancellation after an insertion produces a partial result and prevents replay', async () => {
    const w = world();
    const proposal = await prepared(w);
    const controller = new AbortController();
    w.onInsert = () => controller.abort();
    const result = await applyDocumentEdit(w.deps, proposal, { signal: controller.signal });
    expect(result).toMatchObject({ partial: true, interrupted: true, verified: false });
    expect(w.document.changeTrackingMode).toBe('TrackMineOnly');
    await expect(applyDocumentEdit(w.deps, proposal)).rejects.toThrow(/already attempted/);
});

test('read-back mismatch is not reported as successful application', async () => {
    const w = world();
    const proposal = await prepared(w);
    w.onInsert = () => { w.corruptReadback = true; };
    const result = await applyDocumentEdit(w.deps, proposal);
    expect(result).toMatchObject({ partial: true, verified: false });
    expect(result.warnings[0]).toMatch(/read-back/);
});

test('read-back also catches an unexpected change to an adjacent original paragraph', async () => {
    const w = world();
    const proposal = await prepared(w, [{ id: 'i', kind: 'insert', afterId: 'p-2', beforeId: 'p-3', paragraphs: ['New.'] }]);
    w.onInsert = () => { w.paragraphs.find((p) => p.text === 'Limitations.').text = 'Unexpected neighbor edit.'; };
    const result = await applyDocumentEdit(w.deps, proposal);
    expect(result).toMatchObject({ partial: true, verified: false });
    expect(result.warnings[0]).toMatch(/read-back/);
});

test.each(['line', 'cjk', 'untracked'])('uses the existing %s diff/tracking policy', async (mode) => {
    const w = world();
    if (mode === 'line') w.deps.appState.config.lineDiffEnabled = true;
    if (mode === 'untracked') w.deps.appState.config.trackChangesEnabled = false;
    const after = mode === 'cjk' ? '仍然存在局限。' : 'A limitation remains.';
    const proposal = await prepared(w, [{ id: 'r', kind: 'replace', blockId: 'p-3', before: 'Limitations.', after }]);
    await applyDocumentEdit(w.deps, proposal);
    expect(mode === 'line' ? applySentenceDiffStrategy : mode === 'cjk' ? applyCharDiffStrategy : applyTokenMapStrategy).toHaveBeenCalledTimes(1);
    expect(w.document.changeTrackingMode).toBe('TrackMineOnly');
});

test('insertion at the beginning keeps order and uses body style rather than inheriting a heading', async () => {
    const w = world();
    const proposal = await prepared(w, [{ id: 'i', kind: 'insert', afterId: null, beforeId: 'p-1', paragraphs: ['First.', 'Second.'] }]);
    await applyDocumentEdit(w.deps, proposal);
    expect(w.paragraphs.slice(0, 3).map((p) => p.text)).toEqual(['First.', 'Second.', 'Discussion']);
    expect(w.paragraphs[0].styleBuiltIn).toBe('Normal');
});

test.each(['start', 'end'])('rejects an unrelated paragraph added at the %s after staging', async (edge) => {
    const w = world();
    const change = edge === 'start'
        ? { id: 'i', kind: 'insert', afterId: null, beforeId: 'p-1', paragraphs: ['New.'] }
        : { id: 'i', kind: 'insert', afterId: 'p-4', beforeId: null, paragraphs: ['New.'] };
    const proposal = await prepared(w, [change]);
    w.addParagraph('User paragraph.', edge === 'start' ? 0 : w.paragraphs.length);
    await expect(applyDocumentEdit(w.deps, proposal)).rejects.toThrow(/insertion gap changed/);
    expect(w.mutations).toHaveLength(0);
    expect(proposal.anchor.attempted).toBe(false);
});

test('host synchronization failure after the first write reports a partial edit', async () => {
    const w = world();
    const proposal = await prepared(w, [{ id: 'i', kind: 'insert', afterId: 'p-2', beforeId: 'p-3', paragraphs: ['New.'] }]);
    let failed = false;
    w.context.sync = jest.fn(async () => {
        if (w.mutations.length && !failed) { failed = true; throw new Error('Host sync failed'); }
    });
    const result = await applyDocumentEdit(w.deps, proposal);
    expect(result).toMatchObject({ partial: true, verified: false, warnings: ['Host sync failed'] });
    expect(w.mutations).toHaveLength(1);
    await expect(applyDocumentEdit(w.deps, proposal)).rejects.toThrow(/already attempted/);
});

test('aborted reads and invalid snapshots never write bookmarks or content', async () => {
    const w = world();
    const controller = new AbortController(); controller.abort();
    await expect(readDocumentEditSnapshot({ signal: controller.signal })).rejects.toMatchObject({ name: 'AbortError' });
    await expect(anchorDocumentEdit({ id: 's' }, { snapshotId: 'different' })).rejects.toThrow(/Invalid/);
    w.paragraphs[0].xmlOverride = 'not XML';
    const snapshot = await readDocumentEditSnapshot();
    expect(snapshot.blocks[0]).toMatchObject({ structureUnavailable: true, readOnly: true, ooxml: null });
    expect(w.bookmarks.size).toBe(0);
});

test('heading text cannot be replaced through a prose patch', async () => {
    world();
    const snapshot = await readDocumentEditSnapshot();
    await expect(anchorDocumentEdit(snapshot, { snapshotId: snapshot.id, changes: [
        { kind: 'replace', blockId: 'p-1', before: 'Discussion', after: 'Different heading' },
    ] })).rejects.toThrow(/Invalid replacement target/);
});

test('prepare connects the session, pins attached versions and anchors only verified drafts', async () => {
    const w = world();
    getFile.mockResolvedValue({ fileId: 'f', versionId: 'v', name: 'Reference' });
    readFile.mockResolvedValue({ text: 'Source excerpt' });
    w.deps.fileReferences = [{ fileId: 'f', versionId: 'v', name: 'Reference' }];
    runDocumentEditSession.mockImplementation(async (args) => {
        expect(args.sourceTools.some((t) => t.name === 'file_read')).toBe(true);
        await args.send([{ role: 'user', content: 'request' }]);
        expect(await args.executeSource('file_read', { fileId: 'f', versionId: 'wrong' })).toMatchObject({ ok: true });
        return { status: 'staged', patch: { snapshotId: args.snapshot.id, changes: [{ id: 'i', kind: 'insert', afterId: 'p-2', beforeId: 'p-3', paragraphs: ['New.'] }] } };
    });
    const proposal = await prepareDocumentEdit(w.deps, { instruction: 'Integrate source' });
    expect(proposal.anchor).toBeTruthy();
    expect(readFile).toHaveBeenCalledWith('f', expect.objectContaining({ versionId: 'v' }));
    expect(sendMessages).toHaveBeenCalled();
    expect(w.mutations).toHaveLength(0);
});

test('temporary attachments are bounded source tools rather than appended instructions', async () => {
    const w = world();
    runDocumentEditSession.mockImplementation(async (args) => {
        expect(args.instruction).toBe('Integrate the evidence');
        expect(args.sourceContext).not.toContain('UNTRUSTED BODY');
        expect(args.sourceTools.map((tool) => tool.name)).toEqual(['temporary_source_list', 'temporary_source_read']);
        expect(await args.executeSource('temporary_source_list', {})).toMatchObject({ ok: true,
            result: [{ sourceId: 'temporary-1', name: 'notes.txt', chars: 14 }] });
        expect(await args.executeSource('temporary_source_read', { sourceId: 'temporary-1', offset: 10, limit: 8 }))
            .toMatchObject({ ok: true, result: { text: 'BODY', nextOffset: null } });
        expect(await args.executeSource('temporary_source_read', { sourceId: 'temporary-1', offset: -1 })).toMatchObject({ ok: false });
        return { status: 'no_op' };
    });
    await expect(prepareDocumentEdit(w.deps, { instruction: 'Integrate the evidence',
        temporaryAttachments: [{ name: 'notes.txt', kind: 'text', text: 'UNTRUSTED BODY' }] }))
        .resolves.toMatchObject({ status: 'no_op' });
    expect(w.bookmarks.size).toBe(0);
});

test('no-op and failed sessions do not leave anchors', async () => {
    const w = world();
    runDocumentEditSession.mockResolvedValueOnce({ status: 'no_op' }).mockRejectedValueOnce(new Error('Review failed'));
    expect(await prepareDocumentEdit(w.deps, { instruction: 'Check discussion' })).toEqual({ status: 'no_op' });
    await expect(prepareDocumentEdit(w.deps, { instruction: 'Insert discussion' })).rejects.toThrow('Review failed');
    expect(w.bookmarks.size).toBe(0);
});
