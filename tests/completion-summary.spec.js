const { buildCompletionSummary, buildDocumentCompletionSummary } = require('../src/taskpane/ui/completion-summary.js');

test('ignores missing or malformed proposals', () => {
    expect(buildCompletionSummary()).toBe('');
    expect(buildCompletionSummary({ proposals: [null, 'bad'] })).toBe('');
});

test('summarizes mixed outcomes without claiming full application', () => {
    const text = buildCompletionSummary({ proposals: [
        { state: 'applied', items: [{}, {}] },
        { state: 'pending', items: [{}] },
        { state: 'warning' },
    ] });
    expect(text).toContain('3 proposals with 3 proposed changes');
    expect(text).toContain('1 applied; 1 awaiting review; 1 requiring review');
    expect(text).toContain('Inspect the proposal details');
    expect(text).not.toContain('selected changes were applied');
});

test.each(['unexpected', 'toString', '__proto__'])('bounds descriptions and handles unknown state %s conservatively', (state) => {
    const text = buildCompletionSummary({ proposals: [{ state, summary: 'x'.repeat(5000) }] });
    expect(text).toContain('requiring review');
    expect(text).toContain('…');
    expect(text.length).toBeLessThan(400);
});

test('summarizes non-proposal outcomes and recovers legacy document checks without treating progress as success', () => {
    expect(buildCompletionSummary({ completionSummary: 'Reviewed the document.' })).toBe('Reviewed the document.');
    const status = 'Done: 0 amendment(s), 16 comment(s) across 16 section(s).';
    expect(buildCompletionSummary({ status })).toBe(status);
    expect(buildCompletionSummary({ status: 'Processing document (check-doc)...' })).toBe('');
    expect(buildCompletionSummary({ completionSummary: {}, status: 'Reading the document...' })).toBe('');
});

test('failure and cancellation override stale successful summaries without proposals', () => {
    const completionSummary = 'Document review completed.';
    expect(buildCompletionSummary({ completionSummary, error: 'Word disconnected' })).toContain('error: Word disconnected');
    expect(buildCompletionSummary({ completionSummary, status: 'Cancelled.' })).not.toContain('completed');
    expect(buildCompletionSummary({ completionSummary, status: 'Canceled.' })).toContain('cancelled');
    const cancelled = 'Document review cancelled. Inserted 2 comments.';
    expect(buildCompletionSummary({ completionSummary: cancelled, status: 'Cancelled.' })).toBe(cancelled);
});

test('preserves the initial document review when a later retry adds proposals', () => {
    const text = buildCompletionSummary({ completionSummary: 'Document review incomplete.', proposals: [{ state: 'pending', items: [{}] }] });
    expect(text).toContain('Document review incomplete.');
    expect(text).toContain('Ready for review. The proposed changes have not been applied.');
});

describe('document completion summaries', () => {
    const chunks = [{ id: 'a' }, { id: 'b' }];
    const results = [{ status: 'fulfilled', comment: 'Fix the reference to Section 9.' },
        { status: 'fulfilled', comment: 'Use the defined term consistently.' }];

    test('reports actual coverage and writes with concise review notes, not an invented issue count', () => {
        const text = buildDocumentCompletionSummary({ chunks, results,
            applicationResult: { commentsInserted: 2, amendmentsApplied: 0 } });
        expect(text).toContain('Document review completed. Reviewed 2 of 2 section(s).');
        expect(text).toContain('Inserted 2 comment(s) in Word; applied 0 amendment(s).');
        expect(text).toContain('Review notes: Fix the reference to Section 9. Use the defined term consistently.');
        expect(text).not.toContain('2 issues');
    });

    test('zero comments is a completed review only when all sections were processed', () => {
        const text = buildDocumentCompletionSummary({ chunks, results: results.map((r) => ({ ...r, comment: 'NO ISSUES FOUND.' })) });
        expect(text).toContain('completed');
        expect(text).toContain('Inserted 0 comment(s)');
        expect(text).not.toContain('Review notes:');
        const failed = buildDocumentCompletionSummary({ chunks, results: [{ status: 'rejected', error: 'offline' }, { status: 'rejected' }] });
        expect(failed).toContain('incomplete. Reviewed 0 of 2');
        expect(failed).toContain('2 section(s) failed during processing.');
        expect(failed).not.toContain('completed');
    });

    test.each([
        { errors: ['Word insert failed'] }, { failedChunkIds: ['b'] },
        { partialChunkIds: ['b'] }, { uncertainChunkIds: ['b'] }, { halted: true },
    ])('does not claim success for incomplete application %j', (applicationResult) => {
        const text = buildDocumentCompletionSummary({ chunks, results, applicationResult });
        expect(text).toContain('Document review incomplete.');
        expect(text).not.toContain('review completed');
    });

    test('retains verified writes when cancelled and distinguishes processing from application', () => {
        const text = buildDocumentCompletionSummary({ chunks,
            results: [results[0], { status: 'cancelled' }],
            applicationResult: { commentsInserted: 1, interrupted: true } });
        expect(text).toContain('cancelled. Reviewed 1 of 2');
        expect(text).toContain('Inserted 1 comment(s)');
        expect(text).toContain('already applied remain in the document');
    });

    test('handles empty documents and bounds long model notes', () => {
        expect(buildDocumentCompletionSummary()).toContain('no sections to review');
        const text = buildDocumentCompletionSummary({ chunks: [{}], results: [{ status: 'fulfilled', comment: 'x'.repeat(10000) }] });
        expect(text).toContain('…');
        expect(text.length).toBeLessThan(600);
    });
});

test('document evidence summary separates confirmed issues from withheld claims', () => {
  const { buildDocumentCompletionSummary } = require('../src/taskpane/ui/completion-summary.js');
  const summary = buildDocumentCompletionSummary({
    results: [{ status: 'fulfilled', comment: 'A verified misspelling.', review: { confirmedCount: 1, withheldCount: 4 } },
      { status: 'fulfilled', comment: null, review: { confirmedCount: 0, withheldCount: 2 } }],
    chunks: [{}, {}], applicationResult: { commentsInserted: 1 },
  });
  expect(summary).toContain('confirmed 1 issue(s)');
  expect(summary).toContain('6 candidate issue(s) were withheld');
  expect(summary).toContain('they are not correction requests');
});
