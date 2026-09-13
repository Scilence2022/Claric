import {
    createProposalRecord,
    createProposalRuntime,
    reduceProposalRecord,
    serializeProposalRecord,
} from '../src/taskpane/proposal-runtime.js';

function input(overrides = {}) {
    return {
        proposalId: 'p-1', graphId: 'g-1', taskId: 't-1', attemptId: 'a-1',
        source: { workspaceId: 'w', documentId: 'd', instanceId: 'source-i' },
        target: { workspaceId: 'w', documentId: 'd', instanceId: 'target-i' },
        kind: 'replace', scope: 'selection', title: 'Title', summary: 'Summary',
        items: [{ id: 'one', label: 'One' }, { id: 'two', label: 'Two' }],
        baseRevision: 'r1', reviewRequired: true, createdAt: 1, expiresAt: 9999999999999,
        ...overrides,
    };
}

describe('proposal record/runtime separation', () => {
    test('creates a strict bounded serializable record', () => {
        expect(() => createProposalRecord(input({ callback: () => {} }))).toThrow();
        const record = createProposalRecord(input({ runtime: { nope: true } }));
        expect(Object.keys(record)).toEqual(expect.arrayContaining([
            'proposalId', 'graphId', 'taskId', 'attemptId', 'source', 'target', 'kind', 'scope',
            'title', 'summary', 'items', 'baseRevision', 'state', 'reviewRequired', 'createdAt', 'expiresAt',
        ]));
        expect(record.runtime).toBeUndefined();
        expect(JSON.parse(JSON.stringify(record))).toEqual(record);
        expect(() => createProposalRecord(input({ title: 'x'.repeat(4001) }))).toThrow();
        expect(() => createProposalRecord(input({ summary: 'data:image/png;base64,AAAA' }))).toThrow();
        expect(() => createProposalRecord(input({ items: [{ id: 'x', range: { load() {} } }] }))).toThrow();
        expect(() => createProposalRecord(input({ items: [{ id: 'x', value: () => {} }] }))).toThrow();
    });

    test('serializes only the record contract and reducer is idempotent', () => {
        const record = createProposalRecord(input());
        expect(() => serializeProposalRecord({ ...record, apply: () => {} })).toThrow();
        const serialized = serializeProposalRecord(record);
        expect(serialized.apply).toBeUndefined();
        const partial = reduceProposalRecord(record, 'apply', ['one']);
        expect(partial.state).toBe('applying');
        expect(partial.items).toEqual([
            { id: 'one', label: 'One', status: 'applied' }, { id: 'two', label: 'Two' },
        ]);
        expect(reduceProposalRecord(partial, 'apply', ['two']).state).toBe('applied');
        expect(reduceProposalRecord(partial, 'reject').state).toBe('rejected');
    });

    test('enforces target ownership, review gate, expiry, and stale revision', async () => {
        const apply = jest.fn();
        const base = input();
        const runtime = createProposalRuntime(base, apply, {
            identity: { workspaceId: 'w', documentId: 'd', instanceId: 'other-i' },
            getRevision: jest.fn(() => 'r1'),
        });
        expect((await runtime.apply(['one'])).conflict.code).toBe('target-mismatch');
        const review = createProposalRuntime(input({ reviewRequired: false }), apply, { identity: base.target, getRevision: () => 'r1' });
        expect((await review.apply(['one'])).conflict.code).toBe('review-required');
        const expired = createProposalRuntime(input({ expiresAt: 0 }), apply, { identity: base.target, getRevision: () => 'r1', clock: () => 1 });
        expect((await expired.apply(['one'])).conflict.code).toBe('expired');
        const stale = createProposalRuntime(base, apply, { identity: base.target, getRevision: () => 'r2' });
        expect((await stale.apply(['one'])).conflict).toMatchObject({ code: 'stale', expected: 'r1', actual: 'r2' });
        expect(apply).not.toHaveBeenCalled();
    });

    test('applies selected items partially, then completes without leaking callback', async () => {
        const record = createProposalRecord(input());
        const apply = jest.fn(async (ids, snapshot) => {
            expect(snapshot.apply).toBeUndefined();
            expect(ids).toEqual(apply.mock.calls.length === 1 ? ['one'] : ['two']);
        });
        const runtime = createProposalRuntime(record, apply, { identity: record.target, getRevision: () => 'r1' });
        const first = await runtime.apply(['one']);
        expect(first).toMatchObject({ ok: true, appliedItemIds: ['one'], record: { state: 'applying' } });
        expect((await runtime.apply(['one'])).conflict.code).toBe('no-items');
        const second = await runtime.apply(['two']);
        expect(second.record.state).toBe('applied');
        expect(runtime.getRecord().apply).toBeUndefined();
        expect(runtime.reject().conflict.code).toBe('terminal');
    });
});
