import { createDocumentEvidenceStore } from '../src/lib/document-evidence.js';

function observe(store, blocks) {
    store.add(blocks.map((block) => ({ ...block, offset: 0, nextOffset: null })));
}

test('late relevant sources survive many early reads and repeated reads are deduplicated', () => {
    const blocks = Array.from({ length: 80 }, (_, index) => ({ id: `p${index}`, text: `Unrelated observation ${index}. ` + 'ordinary '.repeat(100) }));
    blocks.push({ id: 'license', text: 'The BSD release permits independent source inspection.' });
    const store = createDocumentEvidenceStore(blocks);
    observe(store, blocks);
    observe(store, [blocks[80]]);
    const selected = store.select({ query: 'Discuss BSD source inspection', maxChars: 1200 });
    expect(selected.blocks.filter((item) => item.id === 'license')).toHaveLength(1);
    expect(selected.coverage).toMatchObject({ observedBlocks: 81, omittedBlocks: 79 });
    expect(JSON.stringify(selected.blocks).length).toBeLessThanOrEqual(1200);
    expect(selected.coverage.note).toContain('omitted');
});

test('overlapping pages merge but unread gaps are never filled in for review', () => {
    const store = createDocumentEvidenceStore([{ id: 'p', text: 'abcdefghijkl' }]);
    store.add([{ id: 'p', text: 'abcd', offset: 0 }, { id: 'p', text: 'cdef', offset: 2 },
        { id: 'p', text: 'ijkl', offset: 8 }]);
    expect(store.select().blocks).toEqual([
        { id: 'p', text: 'abcdef', offset: 0, nextOffset: 6 },
        { id: 'p', text: 'ijkl', offset: 8, nextOffset: null },
    ]);
    expect(() => store.select({ evidenceIds: ['p'] })).toThrow(/Read the complete original/);
    store.add([{ id: 'p', text: 'gh', offset: 6 }]);
    expect(store.select({ evidenceIds: ['p'] }).blocks).toEqual([
        { id: 'p', text: 'abcdefghijkl', offset: 0, nextOffset: null },
    ]);
});

test('draft assertions, unknown IDs and mismatched offsets cannot become original evidence', () => {
    const store = createDocumentEvidenceStore([{ id: 'p', text: 'Original fact.' }]);
    store.add([{ id: 'draft-1', text: 'Invented claim.', offset: 0 },
        { id: 'p', text: 'Invented claim.', offset: 0 }, { id: 'p', text: 'Original fact.', offset: -1 },
        { id: 'p', text: 'Original fact.', offset: 0, originalText: false }, { id: 'p', text: '', offset: 0 }]);
    expect(store.size).toBe(0);
    expect(store.select().blocks).toEqual([]);
    expect(() => store.select({ evidenceIds: ['p'] })).toThrow(/Read the complete original/);
});

test('pinned complete observations outrank relevance and recency without disabling the budget', () => {
    const blocks = [{ id: 'old', text: 'An important original fact.' },
        { id: 'new', text: 'More recent matching keywords.' }];
    const store = createDocumentEvidenceStore(blocks);
    observe(store, blocks);
    const budget = JSON.stringify([{ ...blocks[0], offset: 0, nextOffset: null }]).length;
    expect(store.select({ query: 'matching keywords', evidenceIds: ['old'], maxChars: budget }).blocks[0].id).toBe('old');
    expect(() => store.select({ evidenceIds: ['old', 'new'], maxChars: budget })).toThrow(/Pinned review evidence exceeds/);
});

test('evidence budgets count serialized escaping and Unicode rather than just text length', () => {
    const blocks = [{ id: 'a', text: '"\\\n'.repeat(50) }, { id: 'b', text: '中文 evidence.' }];
    const store = createDocumentEvidenceStore(blocks);
    observe(store, blocks);
    for (const maxChars of [2, 80, 200, 1000]) {
        expect(JSON.stringify(store.select({ maxChars }).blocks).length).toBeLessThanOrEqual(maxChars);
    }
});

test.each([null, ['missing'], Array(25).fill('p')])('invalid evidence pinning is actionable: %j', (evidenceIds) => {
    const store = createDocumentEvidenceStore([{ id: 'p', text: 'fact' }]);
    expect(() => store.select({ evidenceIds })).toThrow(/up to 24 original block IDs/);
});

test('rereading promotes an observed block when relevance is tied', () => {
    const blocks = [{ id: 'a', text: 'First original fact.' }, { id: 'b', text: 'Second original fact.' }];
    const store = createDocumentEvidenceStore(blocks);
    observe(store, blocks);
    observe(store, [blocks[0]]);
    expect(store.select({ maxChars: 100 }).blocks.map((item) => item.id)).toEqual(['a']);
    expect(() => store.select({ maxChars: 0 })).toThrow(/budget/);
});
