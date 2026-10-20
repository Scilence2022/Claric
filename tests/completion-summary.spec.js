const { buildCompletionSummary } = require('../src/taskpane/ui/completion-summary.js');

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
