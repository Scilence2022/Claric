/** @jest-environment jsdom */
jest.mock('../src/lib/word-mutation-verification.js', () => { throw new Error('Verification asset could not be loaded.'); });

import { applyTokenMapStrategy, applyCharDiffStrategy, applySentenceDiffStrategy, applyBlockReplaceStrategy } from '../src/lib/word-diff/index.js';

afterEach(() => { delete global.Word; });

test.each([applyTokenMapStrategy, applyCharDiffStrategy, applySentenceDiffStrategy, applyBlockReplaceStrategy])('a missing verification asset fails before Word writes or tracking changes: %p', async (strategy) => {
    const range = { insertText: jest.fn(), delete: jest.fn(), getOoxml: jest.fn() };
    const context = { sync: jest.fn(), document: { changeTrackingMode: 'Off' } };
    global.Word = { ChangeTrackingMode: { off: 'Off', trackAll: 'TrackAll' } };
    const application = strategy === applyBlockReplaceStrategy
        ? strategy(context, range, 'Amended text.', jest.fn())
        : strategy(context, range, 'Original text.', 'Amended text.', jest.fn());
    await expect(application).rejects.toThrow(/Verification asset/);
    expect(range.insertText).not.toHaveBeenCalled();
    expect(range.delete).not.toHaveBeenCalled();
    expect(range.getOoxml).not.toHaveBeenCalled();
    expect(context.document.changeTrackingMode).toBe('Off');
});
