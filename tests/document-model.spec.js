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
