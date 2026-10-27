import { describeFormatOp } from '../lib/format-ops.js';

/** Shared, lazy proposal UI for native text formatting and section layout. */
export async function stageFormatProposal({ turn, msg, turnDeps, selectionText, actions, stagedResource, makeProposalCard, log, signal }) {
    const layout = turn.type === 'page-layout';
    const prepare = layout ? actions.preparePageLayoutProposal : actions.prepareFormatProposal;
    const discard = layout ? actions.discardPageLayoutProposal : actions.discardFormatProposal;
    const apply = layout ? actions.applyPageLayoutProposal : actions.applyFormatProposal;
    const describe = layout ? (op) => op.label : describeFormatOp;
    const title = layout ? `Proposed page layout (${turn.scope === 'selection' ? 'whole sections containing selection' : 'document'} scope)` : `Proposed changes (${turn.scope} scope)`;
    msg.setStatus('Planning document changes...');
    const proposal = await prepare(turnDeps, {
        instruction: turn.instruction,
        ...(turn.originalInstruction ? { originalInstruction: turn.originalInstruction } : {}),
        scope: turn.scope,
        selectionText,
        cleanupOnly: turn.cleanupOnly === true,
        cleanupRequested: turn.cleanupRequested,
        signal,
        onToken: (t) => msg.appendModelToken({ id: layout ? 'page-layout' : 'format' }, 'content', t),
        onReasoning: (t) => msg.appendModelToken({ id: layout ? 'page-layout' : 'format' }, 'reasoning', t),
    });
    const resource = stagedResource(() => discard?.(turnDeps, proposal));
    if (signal.aborted || !turnDeps.isCurrentSession()) {
        await resource.dispose();
        return;
    }
    if (!proposal.ops || proposal.ops.length === 0) {
        await resource.dispose();
        msg.setStatus(proposal.cleanupSummary?.preserved
            ? `${proposal.cleanupSummary.preserved} empty paragraph(s) could not be safely removed. See the activity log.`
            : layout ? proposal.alreadySatisfied ? 'The requested page layout is already satisfied.' : 'No page layout changes proposed.'
                : turn.cleanupOnly ? 'No verified empty paragraphs found in the selection.' : 'The model proposed no changes.');
        if (layout) return { status: 'no_op', satisfied: proposal.alreadySatisfied === true };
        return;
    }
    msg.setStatus('');
    const countsText = `${proposal.ops.length} change op(s)` + (proposal.cleanupSummary?.preserved
        ? ` · ${proposal.cleanupSummary.preserved} empty paragraph(s) preserved` : '')
        + (proposal.targetSummary ? ` · ${proposal.targetSummary.verifiedParagraphs} verified paragraph(s)` : '')
        + (proposal.targetSummary?.uncertainParagraphs ? ` · ${proposal.targetSummary.uncertainParagraphs} uncertain paragraph(s) preserved` : '');
    const card = makeProposalCard({
        title,
        ...(layout ? { notice: proposal.notice, applyLabel: 'Apply page layout' } : {}),
        countsText,
        comment: null,
        // One selectable change per op; Apply runs only the checked ops.
        items: proposal.ops.map((op, index) => ({
            id: index,
            label: describe(op),
            searchText: op.match ? op.match.trim().slice(0, 60) : undefined,
        })),
        onLocate: (text) => actions.revealTextSnippet(turnDeps, text),
        onApply: async (selectedIds, applyCtx = {}) => {
            try {
                const ops = proposal.ops.filter((_, index) => selectedIds.includes(index));
                const fmtResult = await resource.apply(() => apply(turnDeps, { ...proposal, ops }, applyCtx));
                if (layout) {
                    if (fmtResult?.verified) card.markApplied(`${fmtResult.verifiedSections} section(s) verified.`);
                    else card.markWarning('Page layout could not be verified. Inspect Word before retrying.');
                } else if (fmtResult?.interrupted || fmtResult?.partial) {
                    card.markWarning('Formatting stopped; changes may already be applied. Review the document and draft a new proposal.');
                } else if (proposal.targetSummary?.uncertainParagraphs) {
                    card.markWarning(`${fmtResult?.verifiedParagraphs || proposal.targetSummary.verifiedParagraphs} paragraph(s) verified; ${proposal.targetSummary.uncertainParagraphs} uncertain paragraph(s) preserved for review.`);
                } else if (fmtResult?.alreadySatisfied && fmtResult.verifiedParagraphs > 0) {
                    card.markApplied(`${fmtResult.verifiedParagraphs} paragraph(s) verified; requested formatting was already satisfied.`);
                } else if (fmtResult && fmtResult.appliedRanges === 0 && fmtResult.insertedParagraphs === 0 && !fmtResult.deletedParagraphs) {
                    card.markWarning('Nothing applied — no formatting targets matched. See the activity log.');
                } else if (proposal.cleanupSummary?.preserved) {
                    const removed = fmtResult?.deletedParagraphs ? ` ${fmtResult.deletedParagraphs} empty paragraph(s) removed;` : '';
                    const verified = proposal.cleanupSummary.unverifiable === 0;
                    const message = `Formatting applied;${removed} ${proposal.cleanupSummary.preserved} `
                        + (verified ? 'protected empty paragraph(s) preserved.' : 'empty paragraph(s) could not be safely removed.')
                        + ' See the activity log.';
                    if (verified) card.markApplied(message); else card.markWarning(message);
                } else {
                    card.markApplied();
                }
            } catch (error) {
                log(`Apply failed: ${error.message}`, 'error');
                card.markError(error.message, { retryable: false });
            }
        },
        onReject: async () => {
            await resource.dispose();
            log('Proposal rejected by user.', 'info');
        },
    });
    msg.attachProposal(card, {
        title,
        state: 'pending',
        countsText,
        items: proposal.ops.map((op, index) => ({
            id: index,
            label: describe(op),
            searchText: op.match ? op.match.trim().slice(0, 60) : undefined,
        })),
    });
}
