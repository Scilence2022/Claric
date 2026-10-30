/** Describe one explicit recovery attempt using its verified Word result. */
export function describeSectionRecovery(id, mode, result = null, error = null) {
    const action = { force: 'Force apply', retry: 'Reapply', regenerate: 'Regenerate and apply' }[mode] || 'Recovery';
    const section = `section ${id}`;
    const verified = Math.max(0, Number(result?.appliedParagraphs) || 0);
    const count = `${verified} paragraph edit(s) verified.`;
    const failed = result?.failedChunkIds?.includes(id) || result?.errors?.length;
    const uncertain = result?.uncertainChunkIds?.includes(id) || error?.mutationAttempted;
    const partial = result?.partialChunkIds?.includes(id) || verified > 0;
    const reason = [error?.message || (error ? String(error) : ''), ...(result?.errors || [])].filter(Boolean).join(' ');
    let outcome;
    let message;
    if (!failed && !uncertain && result?.amendmentsApplied > 0) {
        outcome = 'applied';
        message = `${action} succeeded for ${section}: the proposal was applied and verified.` + (verified ? ` ${count}` : '');
    } else if (!failed && !uncertain && result?.noChangeCount > 0) {
        outcome = 'no-change';
        message = `${action} completed for ${section}: no changes were needed; the current text already matches the proposal.`;
    } else if (partial) {
        outcome = uncertain ? 'partial-uncertain' : 'partial';
        message = `${action} partially completed for ${section}. ${count} The section is incomplete; inspect Word before retrying.`;
    } else if (uncertain) {
        outcome = 'uncertain';
        message = `${action} could not be verified for ${section}. ${count} Changes may already be present in Word; inspect the section before retrying.`;
    } else if (result?.interrupted || error?.name === 'AbortError') {
        outcome = 'paused';
        message = `${action} was paused for ${section}. ${count} No successful completion was verified.`;
    } else {
        outcome = 'error';
        message = `${action} failed for ${section}. ${count} No successful completion was verified.`;
    }
    if (reason) message += ` ${['applied', 'no-change'].includes(outcome) ? 'Follow-up error' : 'Reason'}: ${reason}`;
    return { outcome, message, level: ['applied', 'no-change'].includes(outcome) && !error ? 'success' : 'warning' };
}

/** Keep the activity log and visible/persisted card report in agreement. */
export function reportSectionRecovery(id, mode, result, error, log) {
    const report = describeSectionRecovery(id, mode, result, error);
    log(report.message, report.level);
    if (error) throw Object.assign(error instanceof Error ? error : new Error(String(error)), { recoveryReport: report.message });
    return { ...result, recoveryReport: report.message };
}
