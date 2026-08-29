/**
 * Selection-staleness guard for the flat selection-amendment apply route.
 *
 * Regression: applySelectionAmendment re-read document.getSelection() at
 * apply time but never verified it still matched proposal.selectionText. If
 * the user moved the selection between staging and Apply, granular diffing
 * threw and the old catch-fallback replaced whatever was CURRENTLY selected
 * with the old amendment text — a data-loss path. The guard must refuse to
 * write when the selection text drifted (tolerating only whitespace /
 * line-ending normalization).
 */

jest.mock('../src/lib/word-diff/index.js', () => ({
    applyTokenMapStrategy: jest.fn().mockResolvedValue(undefined),
    applySentenceDiffStrategy: jest.fn().mockResolvedValue(undefined),
}));

jest.mock('../src/lib/word-diff/char-diff.js', () => ({
    hasCjk: jest.fn(() => false),
    applyCharDiffStrategy: jest.fn().mockResolvedValue(undefined),
}));

const { applyTokenMapStrategy, applySentenceDiffStrategy } = require('../src/lib/word-diff/index.js');
const { applySelectionAmendment } = require('../src/taskpane/word-actions.js');

/**
 * Builds a Word.run mock whose selection reports `selectionText` after a
 * sync. Strategy writes are no-ops (mocked at module level); `insertText`
 * is spied so the test can assert the destructive fallback never fires.
 */
function setupWord(selectionText) {
    const selection = {
        text: selectionText,
        load: jest.fn(),
        insertText: jest.fn(),
    };
    const context = {
        document: {
            getSelection: () => selection,
            changeTrackingMode: null,
        },
        sync: jest.fn().mockResolvedValue(undefined),
    };
    global.Word = {
        run: async (fn) => fn(context),
        ChangeTrackingMode: { off: 'Off', trackAll: 'TrackAll' },
        InsertLocation: { replace: 'Replace' },
    };
    return { selection, context };
}

function makeDeps() {
    return {
        appState: {
            config: { trackChangesEnabled: true, lineDiffEnabled: false },
            supportsComments: false,
        },
        log: jest.fn(),
    };
}

function makeProposal(selectionText) {
    return {
        selectionText,
        amendedText: `${selectionText} (amended)`,
        commentText: null,
    };
}

describe('applySelectionAmendment staleness guard (flat text route)', () => {
    afterEach(() => {
        delete global.Word;
        jest.clearAllMocks();
    });

    test('applies normally when the selection text is unchanged', async () => {
        const { selection } = setupWord('The original text.');
        const deps = makeDeps();

        await applySelectionAmendment(deps, makeProposal('The original text.'));

        expect(applyTokenMapStrategy).toHaveBeenCalledTimes(1);
        expect(selection.insertText).not.toHaveBeenCalled();
        expect(deps.log).toHaveBeenCalledWith('Changes applied successfully', 'success');
    });

    test('tolerates line-ending and whitespace normalization differences', async () => {
        setupWord('Line one.\r\nLine  two.');
        const deps = makeDeps();

        const result = await applySelectionAmendment(deps, makeProposal('Line one.\nLine two.'));

        expect(result).toBeUndefined();
        expect(applyTokenMapStrategy).toHaveBeenCalledTimes(1);
    });

    test('refuses to write when the selection moved to different text', async () => {
        const { selection } = setupWord('Something else entirely selected now.');
        const deps = makeDeps();

        const result = await applySelectionAmendment(deps, makeProposal('The original text.'));

        expect(result).toEqual(expect.objectContaining({ skipped: true }));
        expect(applyTokenMapStrategy).not.toHaveBeenCalled();
        expect(applySentenceDiffStrategy).not.toHaveBeenCalled();
        // The destructive whole-selection fallback must never fire.
        expect(selection.insertText).not.toHaveBeenCalled();
        expect(deps.log).toHaveBeenCalledWith(expect.stringContaining('Selection changed'), 'warning');
    });
});
