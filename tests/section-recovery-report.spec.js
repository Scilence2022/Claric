import { describeSectionRecovery } from '../src/lib/section-recovery-report.js';

test.each([
    ['force', 'Force apply'], ['retry', 'Reapply'], ['regenerate', 'Regenerate and apply'],
])('%s identifies the requested action and the verified section result', (mode, action) => {
    const report = describeSectionRecovery('chunk-2', mode, { amendmentsApplied: 1, appliedParagraphs: 4, errors: [] });
    expect(report).toMatchObject({ outcome: 'applied', level: 'success' });
    expect(report.message).toContain(`${action} succeeded for section chunk-2`);
    expect(report.message).toContain('4 paragraph edit(s) verified');
});

test.each([
    [{ noChangeCount: 1 }, 'no-change', 'no changes were needed'],
    [{ partialChunkIds: ['c'], appliedParagraphs: 2, errors: ['Later paragraph failed.'] }, 'partial', '2 paragraph edit(s) verified'],
    [{ partialChunkIds: ['c'], uncertainChunkIds: ['c'], appliedParagraphs: 2, errors: ['Read-back differs.'] }, 'partial-uncertain', 'inspect Word'],
    [{ uncertainChunkIds: ['c'], errors: ['Read-back differs.'] }, 'uncertain', 'Changes may already be present'],
    [{ failedChunkIds: ['c'], errors: ['Protected formula changed.'] }, 'error', 'Protected formula changed'],
    [{ interrupted: true }, 'paused', 'was paused'],
    [{}, 'error', 'No successful completion was verified'],
])('reports incomplete or unchanged outcomes without claiming success: %j', (result, outcome, detail) => {
    const report = describeSectionRecovery('c', 'force', result);
    expect(report.outcome).toBe(outcome);
    expect(report.message).toContain(detail);
    expect(report.message).not.toContain('succeeded');
});

test('an exception after a possible write reports uncertainty rather than an untouched document', () => {
    const error = Object.assign(new Error('Word sync failed.'), { mutationAttempted: true });
    const report = describeSectionRecovery('c', 'force', null, error);
    expect(report).toMatchObject({ outcome: 'uncertain', level: 'warning' });
    expect(report.message).toContain('Word sync failed');
    expect(report.message).toContain('Changes may already be present');
});

test('bookmark cleanup failure preserves the verified application result and reports the follow-up error', () => {
    const report = describeSectionRecovery('c', 'force', { amendmentsApplied: 1, appliedParagraphs: 2 }, new Error('Bookmark cleanup failed.'));
    expect(report).toMatchObject({ outcome: 'applied', level: 'warning' });
    expect(report.message).toContain('applied and verified');
    expect(report.message).toContain('Follow-up error: Bookmark cleanup failed');
});
