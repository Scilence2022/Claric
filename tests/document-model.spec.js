import { createDocumentModel, DOCUMENT_EDIT_LIMITS } from '../src/lib/document-model.js';

const snapshot = () => ({ id: 's1', blocks: [
    { id: 'p1', text: 'Results', headingLevel: 1, section: 'Results' },
    { id: 'p2', text: 'The measured result.', section: 'Results' },
    { id: 'p3', text: 'Discussion', headingLevel: 1, section: 'Discussion' },
    { id: 'p4', text: 'The mechanism explains this finding.', section: 'Discussion' },
    { id: 'p5', text: 'Limitations remain.', section: 'Discussion' },
] });
function ready(source = snapshot()) {
    const model = createDocumentModel(source);
    model.setContract({ goal: 'Integrate a discussion', requirements: ['Explain XXX at a relevant location'] });
    model.read({ ids: source.blocks.map((b) => b.id) });
    return model;
}
const insert = { kind: 'insert', afterId: 'p4', paragraphs: ['XXX adds another interpretation.'], reason: 'Before limitations.' };

test('inserts at a chosen gap, edits the inserted draft and compiles against original identities', () => {
    const model = ready();
    model.stage({ operations: [insert] });
    expect(model.read({ ids: ['draft-1'] }).blocks[0]).toMatchObject({ previousId: 'p4', nextId: 'p5' });
    model.stage({ operations: [
        { kind: 'replace', blockId: 'draft-1', text: 'XXX complements the mechanism.', reason: 'Clarify the relationship.' },
        { kind: 'replace', blockId: 'p5', text: 'Despite this interpretation, limitations remain.', reason: 'Repair the transition.' },
    ] });
    expect(model.compile()).toMatchObject({ snapshotId: 's1', revision: 2, changes: [
        { kind: 'insert', afterId: 'p4', beforeId: 'p5', paragraphs: ['XXX complements the mechanism.'] },
        { kind: 'replace', blockId: 'p5', before: 'Limitations remain.', after: 'Despite this interpretation, limitations remain.' },
    ] });
    expect(model.preview().after.some((b) => b.id === 'draft-1')).toBe(true);
    expect(model.preview().before.some((b) => b.id === 'p2')).toBe(false);
    expect(snapshot().blocks[4].text).toBe('Limitations remain.');
});

test('new prose and its requested formatting share one validated draft', () => {
    const model = ready();
    model.stage({ operations: [insert] });
    model.stage({ operations: [{ kind: 'format_new', blockId: 'draft-1', format: { bold: true, italic: true },
        reason: 'Emphasize the new discussion.' }] });
    expect(model.compile().changes[0]).toMatchObject({ kind: 'insert',
        paragraphs: ['XXX adds another interpretation.'], paragraphFormats: [{ bold: true, italic: true }] });
    expect(model.preview().after.find((item) => item.id === 'draft-1').format).toEqual({ bold: true, italic: true });
    expect(() => model.stage({ operations: [{ kind: 'format_new', blockId: 'p4', format: { bold: true }, reason: 'x' }] }))
        .toThrow(/newly inserted/);
    expect(() => model.stage({ operations: [{ kind: 'format_new', blockId: 'draft-1', format: { color: 'red' }, reason: 'x' }] }))
        .toThrow(/bold and italic/);
});

test('outline and search are paged previews and do not grant unread write targets', () => {
    const model = createDocumentModel(snapshot());
    expect(model.outline({ offset: 1, limit: 2 })).toMatchObject({ total: 5, nextOffset: 3 });
    expect(model.outline({ offset: 4 })).toMatchObject({ nextOffset: null });
    expect(model.search({ query: 'MECHANISM' }).blocks[0].id).toBe('p4');
    expect(model.search({ query: 'missing' }).total).toBe(0);
    model.setContract({ goal: 'Add discussion', requirements: ['Keep structure'] });
    expect(() => model.stage({ operations: [insert] })).toThrow(/Read the complete block p4/);
    model.read({ ids: ['p4'] });
    expect(() => model.stage({ operations: [insert] })).toThrow(/p5/);
    model.read({ ids: ['p5'], offset: 5, limit: 100 });
    expect(() => model.stage({ operations: [insert] })).toThrow(/p5/);
    model.read({ ids: ['p5'], offset: 0, limit: 5 });
    model.stage({ operations: [insert] });
});

test('one invalid operation leaves the entire staged batch unchanged, including ID allocation', () => {
    const model = ready();
    expect(() => model.stage({ operations: [insert, { kind: 'replace', blockId: 'missing', text: 'x', reason: 'x' }] })).toThrow(/Unknown block/);
    expect(model.revision).toBe(0);
    expect(model.compile().changes).toEqual([]);
    model.stage({ operations: [insert] });
    expect(model.read({ ids: ['draft-1'] }).blocks[0].text).toContain('XXX');
});

test('discard restores original text and removes inserted content without deleting originals', () => {
    const model = ready();
    model.stage({ operations: [insert, { kind: 'replace', blockId: 'p4', text: 'Edited.', reason: 'Tighten.' }] });
    model.stage({ operations: [{ kind: 'discard', blockId: 'draft-1' }, { kind: 'discard', blockId: 'p4' }] });
    expect(model.compile().changes).toEqual([]);
    expect(() => model.setContract({ goal: 'Different', requirements: ['x'] })).toThrow(/cannot change/);
});

test('scope constraints and protected text are enforced, while adjacent prose insertion remains possible', () => {
    const source = snapshot();
    source.blocks[3].readOnly = true;
    const model = ready(source);
    model.setContract({ goal: 'Edit discussion', requirements: ['Scope'], targetIds: ['p4'] });
    expect(() => model.stage({ operations: [{ kind: 'replace', blockId: 'p4', text: 'x', reason: 'x' }] })).toThrow(/protected/);
    expect(() => model.stage({ operations: [{ kind: 'replace', blockId: 'p2', text: 'x', reason: 'x' }] })).toThrow(/scope/);
    model.stage({ operations: [insert] });
    expect(model.compile().changes).toHaveLength(1);
});

test('table interiors cannot be insertion gaps', () => {
    const source = { id: 's', blocks: [{ id: 'a', text: 'Cell', inTable: true }] };
    const model = ready(source);
    expect(() => model.stage({ operations: [{ ...insert, afterId: 'a' }] })).toThrow(/inside a table/);
});

test('a failed Word structure read cannot become an insertion boundary', () => {
    const source = snapshot();
    source.blocks[4].structureUnavailable = true;
    const model = ready(source);
    expect(() => model.stage({ operations: [insert] })).toThrow(/Word structure could not be read/);
    expect(model.compile().changes).toEqual([]);
});

test('document boundaries and repeated inserts retain their order', () => {
    const model = ready();
    model.stage({ operations: [
        { kind: 'insert', beforeId: 'p1', paragraphs: ['Opening.'], reason: 'Start.' },
        { kind: 'insert', afterId: 'p5', paragraphs: ['First.', 'Last.'], reason: 'End.' },
        { kind: 'insert', beforeId: 'draft-3', paragraphs: ['Middle.'], reason: 'Between draft paragraphs.' },
    ] });
    expect(model.compile().changes).toMatchObject([
        { afterId: null, beforeId: 'p1', paragraphs: ['Opening.'] },
        { afterId: 'p5', beforeId: null, paragraphs: ['First.', 'Middle.', 'Last.'] },
    ]);
});

test.each([
    { operations: [] }, { operations: [null] },
    { operations: [{ ...insert, afterId: undefined }] },
    { operations: [{ ...insert, beforeId: 'p5' }] },
    { operations: [{ ...insert, paragraphs: [] }] },
    { operations: [{ ...insert, paragraphs: ['a\nb'] }] },
    { operations: [{ ...insert, reason: '' }] },
    { operations: [{ ...insert, paragraphs: ['x'.repeat(24001)] }] },
    { operations: [{ kind: 'replace', blockId: 'p4', text: 'a\nb', reason: 'x' }] },
    { operations: [{ kind: 'delete', blockId: 'p4' }] },
])('invalid operations fail without silently altering the draft: %j', (args) => {
    const model = ready();
    expect(() => model.stage(args)).toThrow();
    expect(model.compile().changes).toEqual([]);
});

test.each([null, { id: 's', blocks: [] }, { id: 's', blocks: [{ id: 'x', text: '' }, { id: 'x', text: '' }] },
    { id: 's', blocks: [{ id: '', text: 'x' }] }, { id: 's', blocks: [{ id: 'x', text: 5 }] }])('invalid snapshots are refused', (source) => {
    expect(() => createDocumentModel(source)).toThrow();
});

test('requires a complete bounded contract and valid paging', () => {
    const model = createDocumentModel(snapshot());
    expect(() => model.stage({ operations: [insert] })).toThrow(/contract/);
    for (const args of [{ goal: '', requirements: ['x'] }, { goal: 'x', requirements: [] },
        { goal: 'x', requirements: [''] }, { goal: 'x', requirements: ['x'], targetIds: ['missing'] },
        { goal: 'x', requirements: ['x'], targetIds: 'p4' }]) expect(() => model.setContract(args)).toThrow();
    expect(() => model.read({ ids: [] })).toThrow();
    expect(() => model.read({ ids: ['p1'], offset: -1 })).toThrow();
    expect(() => model.outline({ limit: 101 })).toThrow();
    expect(() => model.search({ query: '' })).toThrow();
});

test('read and patch budgets expose truncation and reject oversized drafts without changing content', () => {
    const source = { id: 's', blocks: [1, 2, 3].map((n) => ({ id: `p${n}`, text: 'x'.repeat(24000) })) };
    const model = ready(source);
    const result = model.read({ ids: ['p1', 'p2', 'p3'], limit: 24000 });
    expect(result.blocks.reduce((n, b) => n + b.text.length, 0)).toBe(DOCUMENT_EDIT_LIMITS.readChars);
    expect(result.blocks[2].nextOffset).toBe(0);
    model.read({ ids: ['p2', 'p3'], offset: 8000, limit: 24000 });
    expect(() => model.stage({ operations: [{ ...insert, afterId: 'p1', paragraphs: ['a'.repeat(24000), 'b'.repeat(24000)] }] })).toThrow(/too large/);
    expect(model.revision).toBe(0);
});

test('draft ID allocation avoids collisions and search pages remain bounded', () => {
    const source = { id: 's', blocks: Array.from({ length: 22 }, (_, i) => ({ id: i ? `p${i}` : 'draft-1', text: 'same' })) };
    const model = createDocumentModel(source);
    expect(model.search({ query: 'same' })).toMatchObject({ total: 22, nextOffset: 20 });
    expect(model.search({ query: 'same', offset: 20 }).blocks).toHaveLength(2);
    model.setContract({ goal: 'x', requirements: ['x'] });
    model.read({ ids: ['draft-1', 'p1'] });
    model.stage({ operations: [{ ...insert, afterId: 'draft-1' }] });
    expect(model.read({ ids: ['draft-2'] }).blocks[0].text).toContain('XXX');
});

test('capabilities distinguish protected headings from unverified insertion boundaries and write scope', () => {
    const source = snapshot();
    source.blocks[4].structureUnavailable = true;
    const model = createDocumentModel(source);
    const outline = model.outline().blocks;
    expect(outline[2]).toMatchObject({ canReplace: false, canInsertBefore: true, canInsertAfter: true,
        readOnlyReasons: ['heading'] });
    expect(outline[3]).toMatchObject({ canReplace: true, canInsertBefore: true, canInsertAfter: false });
    expect(outline[4]).toMatchObject({ canReplace: false, canInsertBefore: false, canInsertAfter: false,
        readOnlyReasons: ['structure_unavailable'] });
    model.setContract({ goal: 'Focus', requirements: ['Focus'], targetIds: ['p4'] });
    expect(model.outline().blocks[1]).toMatchObject({ canReplace: false, canInsertBefore: false, canInsertAfter: false });
});

test('a headings-only outline paginates headings without granting full paragraph reads', () => {
    const model = createDocumentModel(snapshot());
    expect(model.outline({ headingsOnly: true, limit: 1 })).toMatchObject({ total: 2, nextOffset: 1,
        blocks: [{ id: 'p1' }] });
    expect(model.outline({ headingsOnly: true, offset: 1, limit: 1 }).blocks[0].id).toBe('p3');
    expect(() => model.outline({ limit: 260 })).toThrow(/outline limit.*1 to 100/);
    expect(() => model.outline({ limit: 0 })).toThrow(/outline limit/);
    expect(() => model.outline({ headingsOnly: 'yes' })).toThrow(/boolean/);
});

test('replacement argument errors identify the correct field before changing the draft', () => {
    const model = ready();
    model.stage({ operations: [insert] });
    expect(() => model.stage({ operations: [{ kind: 'replace', id: 'draft-1', text: 'Revised.', reason: 'Fix.' }] }))
        .toThrow(/blockId.*not id/);
    expect(model.read({ ids: ['draft-1'] }).blocks[0]).toMatchObject({ text: insert.paragraphs[0], originalText: false });
});

describe('review preview with very long unchanged neighbours', () => {
    const long = (label, chars) => `${label} start. ${'x'.repeat(chars)} ${label} end.`;
    function readyLong(blocks) {
        const model = createDocumentModel({ id: 'long', blocks });
        model.setContract({ goal: 'Extend the log discussion', requirements: ['Keep the log intact'] });
        for (const b of blocks) {
            for (let offset = 0; offset < b.text.length; offset += 24000) model.read({ ids: [b.id], offset, limit: 24000 });
        }
        return model;
    }
    const contextOf = (list, id) => list.find((b) => b.id === id);

    test('an insertion after a 30,000-character paragraph keeps the change complete and excerpts the neighbour tail', () => {
        const log = long('Log', 30000);
        const model = readyLong([{ id: 'p1', text: 'Data quality' }, { id: 'p2', text: log }, { id: 'p3', text: 'Next steps.' }]);
        const paragraph = 'The log shows three recurring sensor faults.';
        model.stage({ operations: [{ kind: 'insert', afterId: 'p2', paragraphs: [paragraph], reason: 'Summarize the log.' }] });
        const preview = model.preview();
        expect(JSON.stringify(preview).length).toBeLessThanOrEqual(60000);
        expect(contextOf(preview.after, 'draft-1').text).toBe(paragraph);
        for (const list of [preview.before, preview.after]) {
            const context = contextOf(list, 'p2');
            const { start, end } = context.elided;
            expect(context.elided).toEqual({ start: 0, end: log.length - 1500, chars: log.length });
            expect(context.text).toBe(`[… unchanged text elided: characters ${start}–${end} of ${log.length} …]${log.slice(end)}`);
            expect(context.text.endsWith('Log end.')).toBe(true);
            expect(contextOf(list, 'p3')).toMatchObject({ text: 'Next steps.' });
            expect(contextOf(list, 'p3').elided).toBeUndefined();
        }
        expect(preview.changes).toMatchObject([{ kind: 'insert', afterId: 'p2', beforeId: 'p3', paragraphs: [paragraph] }]);
    });

    test('a replacement between two long paragraphs shows the facing end of each and the full change', () => {
        const left = long('Left', 30000);
        const right = long('Right', 30000);
        const model = readyLong([{ id: 'p1', text: left }, { id: 'p2', text: 'Short bridge.' }, { id: 'p3', text: right }]);
        model.stage({ operations: [{ kind: 'replace', blockId: 'p2', text: 'A clearer bridge.', reason: 'Clarify.' }] });
        const preview = model.preview();
        expect(JSON.stringify(preview).length).toBeLessThanOrEqual(60000);
        expect(contextOf(preview.before, 'p2')).toMatchObject({ text: 'Short bridge.' });
        expect(contextOf(preview.after, 'p2')).toMatchObject({ text: 'A clearer bridge.' });
        for (const list of [preview.before, preview.after]) {
            expect(contextOf(list, 'p1').elided).toEqual({ start: 0, end: left.length - 1500, chars: left.length });
            expect(contextOf(list, 'p1').text.endsWith('Left end.')).toBe(true);
            expect(contextOf(list, 'p3').elided).toEqual({ start: 1500, end: right.length, chars: right.length });
            expect(contextOf(list, 'p3').text.startsWith('Right start.')).toBe(true);
            expect(contextOf(list, 'p3').text).toMatch(/\[… unchanged text elided: characters 1500–\d+ of \d+ …\]$/);
        }
    });

    test('a long unchanged block between two changes keeps both facing ends', () => {
        const middle = long('Middle', 30000);
        const model = readyLong([{ id: 'p1', text: 'Before.' }, { id: 'p2', text: middle }, { id: 'p3', text: 'After.' }]);
        model.stage({ operations: [
            { kind: 'replace', blockId: 'p1', text: 'Revised before.', reason: 'Clarify.' },
            { kind: 'replace', blockId: 'p3', text: 'Revised after.', reason: 'Clarify.' },
        ] });
        const context = contextOf(model.preview().after, 'p2');
        expect(context.elided).toEqual({ start: 1500, end: middle.length - 1500, chars: middle.length });
        expect(context.text.startsWith('Middle start.')).toBe(true);
        expect(context.text.endsWith('Middle end.')).toBe(true);
    });

    test('many long neighbours shrink their excerpts instead of failing', () => {
        const blocks = Array.from({ length: 24 }, (_, i) => ({ id: `p${i}`, text: long(`B${i}`, 10000) }));
        const model = readyLong(blocks);
        model.stage({ operations: blocks.filter((_, i) => i % 2).slice(0, 12)
            .map((b) => ({ kind: 'insert', afterId: b.id, paragraphs: [`Note after ${b.id}.`], reason: 'Annotate.' })) });
        const preview = model.preview();
        expect(JSON.stringify(preview).length).toBeLessThanOrEqual(60000);
        expect(preview.after.filter((b) => !b.elided && b.id.startsWith('draft-'))).toHaveLength(12);
        const excerpts = preview.after.filter((b) => b.elided);
        expect(excerpts).toHaveLength(24);
        // The 1,500-character default no longer fits, so every block keeps at most 600 per facing side.
        const kept = excerpts.map(({ elided }) => elided.chars - (elided.end - elided.start));
        expect(Math.max(...kept)).toBeLessThanOrEqual(1200);
        expect(Math.min(...kept)).toBeGreaterThan(0);
    });

    test('changed content that alone exceeds the review cap still fails clearly', () => {
        const model = readyLong([{ id: 'p1', text: 'Short.' }, { id: 'p2', text: 'Also short.' }]);
        model.stage({ operations: [{ kind: 'insert', afterId: 'p1', paragraphs: ['a'.repeat(23000), 'b'.repeat(23000)], reason: 'Large addition.' }] });
        expect(() => model.preview()).toThrow(/Review context is too large: the proposed changes alone exceed 60000 characters/);
    });
});
