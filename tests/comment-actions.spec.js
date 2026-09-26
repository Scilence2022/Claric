import { parseCommentDeletionRequest, prepareCommentDeletion, applyCommentDeletion } from '../src/taskpane/comment-actions.js';

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
