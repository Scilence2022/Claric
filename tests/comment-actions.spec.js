import { parseCommentDeletionRequest, prepareCommentDeletion, applyCommentDeletion, stageCommentDeletion } from '../src/taskpane/comment-actions.js';
import { CAPABILITY_BY_TYPE } from '../src/lib/capability-catalog.js';

function world(count = 3) {
    const w = { comments: [], selectionIds: ['c1'], selectionText: 'Selected passage', deleted: [],
        pending: [], failAfter: null, ignoreDeletes: false, failReadback: false, wrote: false };
    const collection = (items) => ({ get items() { return items(); }, load: jest.fn() });
    w.addComment = (id, replies = []) => {
        const comment = { id, content: `Comment ${id}`, authorName: 'Reviewer', resolved: id === 'c2', load: jest.fn(),
            replies: collection(() => replies.map((r) => ({ id: r, content: 'Reply text', authorName: 'Author', load: jest.fn() }))),
            delete: jest.fn(() => w.pending.push(id)) };
        w.comments.push(comment);
        return comment;
    };
    for (let i = 1; i <= count; i++) w.addComment(`c${i}`, i === 1 ? ['r1', 'r2'] : []);
    w.body = { text: 'Original document text.', getComments: jest.fn(() => collection(() => w.comments)) };
    w.document = { body: w.body, comments: collection(() => w.comments), changeTrackingMode: 'TrackMineOnly',
        getSelection: jest.fn(() => ({ text: w.selectionText, load: jest.fn(),
            getComments: () => collection(() => w.comments.filter((c) => w.selectionIds.includes(c.id))) })) };
    w.context = { document: w.document, sync: jest.fn(async () => {
        const queued = w.pending.splice(0);
        if (!queued.length && w.wrote && w.failReadback) throw new Error('Readback failed');
        for (let i = 0; i < queued.length; i++) {
            if (w.failAfter === i) throw new Error('Host delete sync failed');
            w.wrote = true;
            if (!w.ignoreDeletes) {
                w.comments = w.comments.filter((c) => c.id !== queued[i]);
                w.deleted.push(queued[i]);
            }
        }
        if (queued.length) w.onWrite?.();
    }) };
    global.Word = { run: async (fn) => { w.pending = []; return fn(w.context); } };
    w.deps = { appState: { supportsComments: true }, log: jest.fn() };
    return w;
}
afterEach(() => { delete global.Word; delete global.Office; });

test.each(['delete all comments', 'Please remove all the comments.', 'Can you delete all comments?',
    'delete all comments including their replies from the document', '删除所有批注', '请删除文档中的所有批注', '把全部批注删除'])(
'recognizes explicit document-wide deletion: %s', (instruction) => {
    expect(parseCommentDeletionRequest(instruction)).toEqual({ scope: 'document' });
});
test.each(['remove all comments in the selection', 'Delete all comments from selected text', '清除选区内的全部批注'])(
'keeps explicit selection scope: %s', (instruction) => {
    expect(parseCommentDeletionRequest(instruction)).toEqual({ scope: 'selection' });
});
test.each(['delete all resolved comments', 'delete all comments by Alice', 'delete all comments except the first',
    '不要删除所有批注', 'How do I delete all comments?', 'delete all comments and polish the article',
    'delete all code comments', 'delete all comments in chapter 2', 'resolve all comments', '删除所有已解决的批注'])(
'never broadens filtered or unsupported instructions: %s', (instruction) => {
    expect(parseCommentDeletionRequest(instruction)).toBeNull();
});

// Wording the capability planner actually produced (offline pilot, 2026-10-11) plus close variants.
test.each([
    'Delete all Word comment threads and replies in the entire document.',
    'Delete all comments in the entire document.',
    'Delete all comments in the whole document',
    'Delete all comments in this entire Word document.',
    'Delete all comments in the document, including all comment threads and replies.',
    'Delete all comments, including all comment threads and replies, in the document.',
    'Delete all comment threads and replies in the document. Do not alter table content, structure, or formatting.',
    'Delete all comments in the document. Keep the tables unchanged.',
    'Delete all comments in the document and keep the tables unchanged.',
    'Delete all comments in the document while leaving all other content intact.',
    'Delete all comments in the document without modifying the text or formatting.',
    'Delete all comments. Preserve the document structure and layout.',
    "Delete all comments. Don’t change anything else.",
    'Delete all comments (including replies) in the document.',
    'Please delete every Word comment thread and reply throughout the document. Leave everything else unchanged.',
])('accepts harmless planner variants as document-wide deletion: %s', (instruction) => {
    expect(parseCommentDeletionRequest(instruction)).toEqual({ scope: 'document' });
});
test.each([
    'Delete all Word comment threads and replies in the selection.',
    'Delete all comments in the current selection. Keep the tables unchanged.',
    'Delete all comments in the selection, including all replies.',
])('keeps selection scope for harmless planner variants: %s', (instruction) => {
    expect(parseCommentDeletionRequest(instruction)).toEqual({ scope: 'selection' });
});
test.each([
    // filters on the comment set
    'Delete all comments by Alice', 'Delete all resolved comments', 'Delete all Word comment threads and replies by Alice in the entire document.',
    'Delete all unresolved comments in the document.', 'Delete all comments containing "TODO" in the document.',
    'Delete all comments that mention the budget.', 'Delete all comments with replies.', 'Delete all comments older than 2024.',
    'Delete all comments in the document except the first.', 'Delete all comments in the document, but only resolved ones.',
    'Delete all comments and replies from Alice.', 'Delete all comments in the document by reviewer Bob.',
    'Delete all comments in the document and resolve the rest.', 'Delete all comments in the document and polish the text.',
    'Delete all comments in the document and then summarize it.',
    // filters smuggled in as "preservation" clauses
    'Delete all comments in the document. Do not delete resolved comments.',
    'Delete all comments in the document. Do not delete comments by Alice.',
    'Delete all comments in the document. Keep the resolved comments unchanged.',
    'Delete all comments in the document. Keep Alice’s comments.',
    'Delete all comments in the document. Keep the first comment intact.',
    'Delete all comments in the document. Keep the replies unchanged.',
    'Delete all comments in the document. Leave the comment text untouched.',
    'Delete all comments in the document. Preserve the comment threads.',
    'Delete all comments in the document, without deleting the resolved threads.',
    'Delete all comments in the document. Do not alter table content, but keep comments from Alice.',
    'Delete all comments in the document. Keep the tables.',
    // not a deletion request at all
    'Do not delete all comments.', "Don't delete all comments in the document.", 'Keep all comments in the document.',
    'Reply to all comments in the document.', 'Resolve all comments in the document.', 'Delete all the tables in the document.',
    'Delete all comments in the document and the selection.', 'Delete all comments in the selection from the document.',
])('never widens filtered or unsupported planner wording to deletion of everything: %s', (instruction) => {
    expect(parseCommentDeletionRequest(instruction)).toBeNull();
});
test('rejects pathological input quickly instead of backtracking', () => {
    const started = Date.now();
    for (const tail of [' and replies', ' and keep the tables unchanged', ' do not alter the text and']) {
        expect(parseCommentDeletionRequest(`Delete all comments${tail.repeat(400)} by Alice`)).toBeNull();
    }
    expect(Date.now() - started).toBeLessThan(1000);
});
test('every instruction the capability catalog tells the planner to write is executable', () => {
    const { description } = CAPABILITY_BY_TYPE.get('comment_management');
    const examples = [...description.matchAll(/"([^"]+)"/g)].map((match) => match[1]);
    expect(examples).toEqual(expect.arrayContaining(['Delete all comments in the document.', 'Delete all comments in the selection.']));
    for (const example of examples) expect(parseCommentDeletionRequest(example)).not.toBeNull();
    expect(parseCommentDeletionRequest('Delete all comments in the selection.')).toEqual({ scope: 'selection' });
    expect(parseCommentDeletionRequest('删除选区内的所有批注')).toEqual({ scope: 'selection' });
});

test('prepares threads and replies without changing Word, then deletes and verifies every thread', async () => {
    const w = world();
    const proposal = await prepareCommentDeletion(w.deps, { instruction: 'delete all comments' });
    expect(proposal).toMatchObject({ scope: 'document', replies: 2, attempted: false });
    expect(proposal.comments[1].resolved).toBe(true);
    expect(w.deleted).toEqual([]);
    expect(await applyCommentDeletion(w.deps, proposal)).toMatchObject({ deleted: 3, repliesDeleted: 2, remaining: 0, verified: true, partial: false });
    expect(w.comments).toHaveLength(0);
    expect(w.body.text).toBe('Original document text.');
    expect(w.document.changeTrackingMode).toBe('TrackMineOnly');
    await expect(applyCommentDeletion(w.deps, proposal)).rejects.toThrow(/already attempted/);
});

test('uses the full document collection when the desktop API is available', async () => {
    const w = world();
    global.Office = { context: { requirements: { isSetSupported: () => true } } };
    const proposal = await prepareCommentDeletion(w.deps, { instruction: 'delete all comments' });
    expect(proposal.comments).toHaveLength(3);
    expect(w.body.getComments).not.toHaveBeenCalled();
});

test('selected comment IDs remain fixed when the user moves the selection before apply', async () => {
    const w = world();
    const proposal = await prepareCommentDeletion(w.deps, { instruction: 'delete all comments in the selection' });
    w.selectionIds = ['c2'];
    expect(await applyCommentDeletion(w.deps, proposal)).toMatchObject({ deleted: 1, verified: true });
    expect(w.comments.map((c) => c.id)).toEqual(['c2', 'c3']);
});

test.each(['added', 'text', 'reply', 'resolved', 'missing'])('refuses %s comment drift before any deletion', async (kind) => {
    const w = world();
    const proposal = await prepareCommentDeletion(w.deps, { instruction: 'delete all comments' });
    if (kind === 'added') w.addComment('c4');
    if (kind === 'text') w.comments[0].content = 'Changed';
    if (kind === 'reply') w.comments[0].replies = { items: [{ id: 'new', content: 'New reply', authorName: 'Other', load: jest.fn() }], load: jest.fn() };
    if (kind === 'resolved') w.comments[0].resolved = true;
    if (kind === 'missing') w.comments.pop();
    await expect(applyCommentDeletion(w.deps, proposal)).rejects.toThrow(/Comments changed/);
    expect(w.deleted).toEqual([]);
    expect(proposal.attempted).toBe(false);
});

test('a failed delete sync is read back without replaying the deletion', async () => {
    const w = world();
    const proposal = await prepareCommentDeletion(w.deps, { instruction: 'delete all comments' });
    w.failAfter = 1;
    expect(await applyCommentDeletion(w.deps, proposal)).toMatchObject({ deleted: 1, remaining: 2, verified: false, partial: true });
    await expect(applyCommentDeletion(w.deps, proposal)).rejects.toThrow(/already attempted/);
});

test('a host that ignores deletion never produces a success result', async () => {
    const w = world();
    const proposal = await prepareCommentDeletion(w.deps, { instruction: 'delete all comments' });
    w.ignoreDeletes = true;
    expect(await applyCommentDeletion(w.deps, proposal)).toMatchObject({ deleted: 0, remaining: 3, verified: false });
});

test('new comments added during deletion prevent an all-comments success claim', async () => {
    const w = world();
    const proposal = await prepareCommentDeletion(w.deps, { instruction: 'delete all comments' });
    w.onWrite = () => w.addComment('new');
    expect(await applyCommentDeletion(w.deps, proposal)).toMatchObject({ deleted: 3, remaining: 1, verified: false });
});

test('cancellation stops after a bounded batch and verifies the partial outcome', async () => {
    const w = world(65);
    const proposal = await prepareCommentDeletion(w.deps, { instruction: 'delete all comments' });
    const controller = new AbortController();
    w.onWrite = () => controller.abort();
    expect(await applyCommentDeletion(w.deps, proposal, { signal: controller.signal }))
        .toMatchObject({ deleted: 32, remaining: 33, verified: false, interrupted: true });
});

test('readback failure reports uncertainty after a write', async () => {
    const w = world();
    const proposal = await prepareCommentDeletion(w.deps, { instruction: 'delete all comments' });
    w.failReadback = true;
    const result = await applyCommentDeletion(w.deps, proposal);
    expect(result).toMatchObject({ verified: false, partial: true, remaining: null });
    expect(result.warnings[0]).toContain('Could not verify');
});

test('unsupported hosts, empty selections, unsupported filters and preflight cancellation do not write', async () => {
    const w = world();
    await expect(prepareCommentDeletion({ appState: { supportsComments: false } }, { instruction: 'delete all comments' })).rejects.toThrow(/WordApi 1.4/);
    await expect(prepareCommentDeletion(w.deps, { instruction: 'delete all comments by Alice' })).rejects.toThrow(/not supported/);
    w.selectionText = '';
    await expect(prepareCommentDeletion(w.deps, { instruction: 'delete all comments in the selection' })).rejects.toThrow(/Select the text/);
    const controller = new AbortController(); controller.abort();
    await expect(prepareCommentDeletion(w.deps, { instruction: 'delete all comments', signal: controller.signal })).rejects.toMatchObject({ name: 'AbortError' });
    expect(w.deleted).toEqual([]);
});

describe('planned comment_management tasks reach the real deletion path', () => {
    const stage = (w, instruction) => {
        const msg = { setStatus: jest.fn(), setText: jest.fn(), attachProposal: jest.fn() };
        const makeProposalCard = jest.fn(() => ({ markApplied: jest.fn(), markWarning: jest.fn(), markError: jest.fn() }));
        const turnDeps = { ...w.deps, isCurrentSession: () => true };
        return { msg, makeProposalCard,
            run: () => stageCommentDeletion({ turn: { instruction, planned: true }, msg, turnDeps, actions: {}, makeProposalCard }) };
    };

    test.each([
        'Delete all Word comment threads and replies in the entire document.',
        'Delete all comment threads and replies in the document. Do not alter table content, structure, or formatting.',
    ])('stages the full-document deletion for the planner wording: %s', async (instruction) => {
        const w = world();
        const staged = stage(w, instruction);
        expect(await staged.run()).toEqual({ status: 'staged' });
        expect(staged.makeProposalCard).toHaveBeenCalledWith(expect.objectContaining({ countsText: '3 comment thread(s), 2 reply/replies' }));
        expect(w.deleted).toEqual([]);
    });

    test('a planned task that carries a filter stays unsupported and reads nothing from Word', async () => {
        const w = world();
        for (const instruction of ['Delete all resolved Word comment threads and replies in the entire document.',
            'Delete all comments in the document. Do not delete comments by Alice.']) {
            const staged = stage(w, instruction);
            await expect(staged.run()).rejects.toThrow(/not supported/);
            expect(staged.makeProposalCard).not.toHaveBeenCalled();
        }
        expect(w.document.comments.load).not.toHaveBeenCalled();
        expect(w.deleted).toEqual([]);
    });
});
