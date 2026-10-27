/** @jest-environment jsdom */
jest.mock('../src/lib/word-diff/revision-diff.js', () => { throw new Error('Native edit asset unavailable.'); });

import { applyTokenMapStrategy, applyCharDiffStrategy, applySentenceDiffStrategy } from '../src/lib/word-diff/index.js';

test.each([applyTokenMapStrategy, applyCharDiffStrategy, applySentenceDiffStrategy])('an unavailable native locator stops before host actions: %p', async (strategy) => {
    const range = { getOoxml: jest.fn(), insertText: jest.fn(), delete: jest.fn() };
    const context = { sync: jest.fn(), document: { changeTrackingMode: 'Off' } };
    await expect(strategy(context, range, 'Before.', 'After.', jest.fn())).rejects.toThrow(/Native edit asset unavailable/);
    expect(range.getOoxml).not.toHaveBeenCalled();
    expect(range.insertText).not.toHaveBeenCalled();
    expect(range.delete).not.toHaveBeenCalled();
    expect(context.sync).not.toHaveBeenCalled();
    expect(context.document.changeTrackingMode).toBe('Off');
});
