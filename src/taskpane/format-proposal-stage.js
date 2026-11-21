import { describeFormatOp } from '../lib/format-ops.js';

/** Shared, lazy proposal UI for native formatting, section layout and header/footer content. */
export async function stageFormatProposal({ turn, msg, turnDeps, selectionText, actions, stagedResource, makeProposalCard, log, signal }) {
    const layout = turn.type === 'page-layout';
    const headerFooter = turn.type === 'header-footer';
    const native = layout || headerFooter;
    const [prepare, discard, apply] = headerFooter
        ? [actions.prepareHeaderFooterProposal, actions.discardHeaderFooterProposal, actions.applyHeaderFooterProposal]
        : layout ? [actions.preparePageLayoutProposal, actions.discardPageLayoutProposal, actions.applyPageLayoutProposal]
            : [actions.prepareFormatProposal, actions.discardFormatProposal, actions.applyFormatProposal];
    const describe = native ? (op) => op.label : describeFormatOp;
    const title = headerFooter ? 'Proposed header and footer changes'
        : layout ? `Proposed page layout (${turn.scope === 'selection' ? 'whole sections containing selection' : 'document'} scope)`
            : `Proposed changes (${turn.scope} scope)`;
    const label = headerFooter ? 'header/footer' : 'page layout';
    msg.setStatus('Planning document changes...');
    const proposal = await prepare(turnDeps, {
        instruction: turn.instruction,
        ...(turn.originalInstruction ? { originalInstruction: turn.originalInstruction } : {}),
        scope: turn.scope,
        selectionText,
        ...(headerFooter && turn.selectionOnly ? { selectionOnly: true } : {}),
        cleanupOnly: turn.cleanupOnly === true,
        cleanupRequested: turn.cleanupRequested,
        signal,
        onToken: (t) => msg.appendModelToken({ id: turn.type }, 'content', t),
        onReasoning: (t) => msg.appendModelToken({ id: turn.type }, 'reasoning', t),
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
            : native ? proposal.alreadySatisfied ? `The requested ${label} is already satisfied.` : `No ${label} changes proposed.`
                : turn.cleanupOnly ? 'No verified empty paragraphs found in the selection.' : 'The model proposed no changes.');
        if (native) return { status: 'no_op', satisfied: proposal.alreadySatisfied === true };
        return;
    }
    msg.setStatus('');
    const countsText = `${proposal.ops.length} change op(s)` + (proposal.cleanupSummary?.preserved
        ? ` · ${proposal.cleanupSummary.preserved} empty paragraph(s) preserved` : '')
        + (proposal.targetSummary ? ` · ${proposal.targetSummary.verifiedParagraphs} verified paragraph(s)` : '')
        + (proposal.targetSummary?.uncertainParagraphs ? ` · ${proposal.targetSummary.uncertainParagraphs} uncertain paragraph(s) preserved` : '');
    const items = proposal.ops.map((op, index) => ({
        id: index,
        label: describe(op),
        searchText: op.match ? op.match.trim().slice(0, 60) : undefined,
        ...(headerFooter && op.before !== undefined ? { before: op.before, after: op.after } : {}),
    }));
    const card = makeProposalCard({
        title,
        ...(native ? { notice: proposal.notice, applyLabel: `Apply ${label}` } : {}),
        countsText,
        comment: null,
        // One selectable change per op; Apply runs only the checked ops.
        items,
        onLocate: (text) => actions.revealTextSnippet(turnDeps, text),
        onApply: async (selectedIds, applyCtx = {}) => {
            try {
                const ops = proposal.ops.filter((_, index) => selectedIds.includes(index));
                const fmtResult = await resource.apply(() => apply(turnDeps, { ...proposal, ops }, applyCtx));
                if (headerFooter) {
                    if (fmtResult?.verified) card.markApplied(`${fmtResult.verifiedStories} header/footer change(s) verified.`);
                    else card.markWarning('Header/footer content could not be verified. Inspect Word before retrying.');
                } else if (layout) {
                    if (fmtResult?.verified) card.markApplied(`${fmtResult.verifiedSections} section(s) verified.`, { noChange: fmtResult.alreadySatisfied === true });
                    else card.markWarning('Page layout could not be verified. Inspect Word before retrying.');
                } else if (fmtResult?.interrupted || fmtResult?.partial) {
                    card.markWarning('Formatting stopped; changes may already be applied. Review the document and draft a new proposal.');
                } else if (proposal.targetSummary?.uncertainParagraphs) {
                    card.markWarning(`${fmtResult?.verifiedParagraphs || proposal.targetSummary.verifiedParagraphs} paragraph(s) verified; ${proposal.targetSummary.uncertainParagraphs} uncertain paragraph(s) preserved for review.`);
                } else if (fmtResult?.alreadySatisfied && fmtResult.verifiedParagraphs > 0) {
                    card.markApplied(`${fmtResult.verifiedParagraphs} paragraph(s) verified; requested formatting was already satisfied.`, { noChange: true });
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
                if (headerFooter && error.mutationAttempted) card.markWarning(error.message);
                else card.markError(error.message, { retryable: false });
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
        items,
    });
}
