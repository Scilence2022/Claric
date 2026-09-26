/** Native Word comment deletion. Text-edit pipelines never participate. */

const BATCH_SIZE = 32;
const DELETE_VERBS = '删除|移除|清除|清空|去掉|去除';
const LOCATIONS = '全文|文档|文章|选区|所选内容|选中的文本';

/** Only unqualified bulk deletion is executable; filters must never broaden to all. */
export function parseCommentDeletionRequest(instruction) {
    const text = String(instruction || '').trim().replace(/[.!?。！？]+$/, '').trim()
        .replace(/^(?:(?:please|(?:can|could|would) you)\s+)+/i, '')
        .replace(/^(?:请(?:帮我)?|麻烦(?:你)?(?:帮我)?)/, '');
    const en = text.match(/^(?:delete|remove|clear|erase)\s+(?:all(?:\s+of)?(?:\s+the)?|every)\s+(?:comments?|comment threads?)(?:\s+(?:and|including)\s+(?:all\s+)?(?:their\s+)?replies)?(?:\s+(?:in|from|on|throughout)\s+(?:(?:the|this|entire|whole)\s+)?(document|article|selection|selected text|selected passage))?$/i);
    if (en) return { scope: /selection|selected/i.test(en[1] || '') ? 'selection' : 'document' };
    const compact = text.replace(/\s+/g, '');
    const location = `(?:(${LOCATIONS})(?:中|内|里)?的?)?`;
    const zh = compact.match(new RegExp(`^(?:${DELETE_VERBS})${location}(?:所有|全部)(?:批注|评论|comments?)$`, 'i'))
        || compact.match(new RegExp(`^(?:把|将)?${location}(?:所有|全部)(?:批注|评论|comments?)(?:全部)?(?:${DELETE_VERBS})$`, 'i'));
    return zh ? { scope: /选/.test(zh[1] || '') ? 'selection' : 'document' } : null;
}

function check(signal) { if (signal?.aborted) throw new DOMException('Comment deletion cancelled.', 'AbortError'); }
function checkHost(deps) {
    if (deps.appState?.supportsComments === false || typeof Word === 'undefined') {
        throw new Error('This Word host cannot manage comments. WordApi 1.4 is required.');
    }
}
function documentComments(context) {
    // Use the complete document collection on newer desktop hosts. The
    // cross-platform WordApi 1.4 path exposes comments through the body.
    if (globalThis.Office?.context?.requirements?.isSetSupported('WordApiDesktop', '1.4')) return context.document.comments;
    if (typeof context.document.body.getComments !== 'function') throw new Error('This Word host cannot read comments.');
    return context.document.body.getComments();
}
function snapshot(comment) {
    return { id: comment.id, content: comment.content, author: comment.authorName, resolved: comment.resolved,
        replies: comment.replies.items.map((reply) => ({ id: reply.id, content: reply.content, author: reply.authorName }))
            .sort((a, b) => a.id.localeCompare(b.id)) };
}
async function readComments(context, scope, signal) {
    let collection;
    if (scope === 'selection') {
        const range = context.document.getSelection();
        range.load('text');
        await context.sync();
        check(signal);
        if (!range.text?.trim()) throw new Error('Select the text whose comments you want to delete.');
        collection = range.getComments();
    } else collection = documentComments(context);
    collection.load('items');
    await context.sync();
    check(signal);
    const entries = [];
    for (let start = 0; start < collection.items.length; start += BATCH_SIZE) {
        const batch = collection.items.slice(start, start + BATCH_SIZE);
        for (const comment of batch) {
            comment.load('id,content,authorName,resolved');
            comment.replies.load('items');
        }
        await context.sync();
        check(signal);
        for (const comment of batch) for (const reply of comment.replies.items) reply.load('id,content,authorName');
        await context.sync();
        check(signal);
        entries.push(...batch.map((comment) => ({ comment, data: snapshot(comment) })));
    }
    if (entries.some(({ data }) => typeof data.id !== 'string' || !data.id)
        || new Set(entries.map(({ data }) => data.id)).size !== entries.length) throw new Error('Word returned unverifiable comment identities.');
    return entries;
}

export async function prepareCommentDeletion(deps, { instruction, signal } = {}) {
    checkHost(deps);
    check(signal);
    const request = parseCommentDeletionRequest(instruction);
    if (!request) throw new Error('This comment operation is not supported. Specify deletion of all comments in the document or selected text.');
    const comments = await Word.run(async (context) => (await readComments(context, request.scope, signal)).map(({ data }) => data));
    return { ...request, comments, attempted: false, replies: comments.reduce((sum, comment) => sum + comment.replies.length, 0) };
}

/** Delete only the reviewed IDs; a fresh read verifies the actual outcome. */
export async function applyCommentDeletion(deps, proposal, { signal } = {}) {
    checkHost(deps);
    check(signal);
    if (!proposal || !['document', 'selection'].includes(proposal.scope) || !proposal.comments?.length) throw new Error('No comment deletion proposal is available.');
    if (proposal.attempted) throw new Error('Comment deletion was already attempted. Inspect Word and prepare a fresh proposal.');
    const ids = new Set(proposal.comments.map((comment) => comment.id));
    const result = { deleted: 0, repliesDeleted: 0, remaining: null, verified: false, partial: false, interrupted: false, warnings: [] };
    try {
        await Word.run(async (context) => {
            const current = await readComments(context, 'document', signal);
            const byId = new Map(current.map((entry) => [entry.data.id, entry]));
            if ((proposal.scope === 'document' && current.length !== ids.size)
                || proposal.comments.some((comment) => JSON.stringify(byId.get(comment.id)?.data) !== JSON.stringify(comment))) {
                throw new Error('Comments changed since preparation. Generate a fresh deletion proposal.');
            }
            const targets = proposal.comments.map((comment) => byId.get(comment.id).comment);
            if (targets.some((comment) => typeof comment.delete !== 'function')) throw new Error('This Word host cannot delete comments.');
            for (let start = 0; start < targets.length; start += BATCH_SIZE) {
                check(signal);
                proposal.attempted = true;
                for (const comment of targets.slice(start, start + BATCH_SIZE)) comment.delete();
                await context.sync();
                check(signal);
            }
        });
    } catch (error) {
        if (!proposal.attempted) throw error;
        result.interrupted = signal?.aborted || error.name === 'AbortError';
        result.warnings.push(error.message);
    }
    // A failed sync may already have deleted some comments. Never replay it;
    // verify using a new request context, even after cancellation.
    try {
        const remaining = await Word.run(async (context) => (await readComments(context, 'document')).map(({ data }) => data.id));
        const remainingIds = new Set(remaining);
        const deleted = proposal.comments.filter((comment) => !remainingIds.has(comment.id));
        result.deleted = deleted.length;
        result.repliesDeleted = deleted.reduce((sum, comment) => sum + comment.replies.length, 0);
        result.remaining = proposal.scope === 'document' ? remaining.length : remaining.filter((id) => ids.has(id)).length;
        result.verified = result.deleted === ids.size && result.remaining === 0;
    } catch (error) { result.warnings.push(`Could not verify comment deletion: ${error.message}`); }
    result.partial = !result.verified;
    return result;
}

/** The native-action card uses explicit deletion wording, never a text diff. */
export async function stageCommentDeletion({ turn, msg, turnDeps, actions, makeProposalCard, signal }) {
    msg.setStatus('Reading comments...');
    const proposal = await (actions.prepareCommentDeletion || prepareCommentDeletion)(turnDeps, { instruction: turn.instruction, signal });
    check(signal);
    if (!turnDeps.isCurrentSession()) return;
    if (!proposal.comments.length) {
        msg.setStatus('');
        msg.setText(`No comments found in the ${proposal.scope === 'selection' ? 'selected text' : 'document'}.`);
        return { status: 'no_op', satisfied: true };
    }
    const countsText = `${proposal.comments.length} comment thread(s), ${proposal.replies} reply/replies`;
    const item = { id: 'delete-comments', label: `Delete ${countsText} from the ${proposal.scope === 'selection' ? 'selected text' : 'document'}` };
    const title = 'Delete comments';
    const card = makeProposalCard({
        title, countsText, items: [item], applyLabel: 'Delete comments',
        comment: 'Includes resolved threads and their replies. Word deletes comments directly; this is not a tracked text edit.',
        onApply: async (selectedIds, applyCtx) => {
            if (!selectedIds?.includes(item.id)) return;
            try {
                const result = await (actions.applyCommentDeletion || applyCommentDeletion)(turnDeps, proposal, applyCtx);
                if (result.verified) card.markApplied(`Deleted ${result.deleted} comment thread(s) and ${result.repliesDeleted} reply/replies.`);
                else card.markWarning(`Comment deletion was not fully verified. ${result.deleted} thread(s) confirmed removed. ${result.warnings.join(' ')} Review Word before retrying.`);
            } catch (error) { card.markError(error.message); }
        },
    });
    msg.attachProposal(card, { title, state: 'pending', countsText, items: [item] });
    msg.setStatus('Comment deletion is ready for review.');
    return { status: 'staged' };
}
