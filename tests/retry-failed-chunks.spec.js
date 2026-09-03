/** @jest-environment jsdom */

/**
 * retryFailedChunks tests.
 *
 * Regression: retry chunks used to be rebuilt as text-only stubs
 * ({ id, text, tokenCount, overlapText }) without a `paragraphs` field, and
 * the orchestrator was handed documentContext: null. Every retry chunk
 * crashed inside the orchestrator's message composer (chunk.paragraphs.map /
 * formatContextPrefix's context.definitions) and was instantly re-marked
 * rejected, so "Click to retry failed chunks" could never succeed.
 *
 * The contract under test: the retry re-drives the ORIGINAL chunk objects
 * (paragraphs intact, same ids for bookmark re-anchoring).
 */

jest.mock('../src/lib/orchestrator.js', () => ({
    processChunksParallel: jest.fn(),
}));
jest.mock('../src/lib/reassembler.js', () => ({
    applyChunkResults: jest.fn(),
    cleanupBookmarks: jest.fn(),
}));

const { processChunksParallel } = require('../src/lib/orchestrator.js');
const { applyChunkResults, cleanupBookmarks } = require('../src/lib/reassembler.js');
const { retryFailedChunks } = require('../src/taskpane/word-actions.js');

function makeFailedResult(chunk) {
    return {
        chunkId: chunk.id,
        chunkIndex: 0,
        status: 'rejected',
        amendment: null,
        comment: null,
        error: 'HTTP 500: boom',
        reasoning: null,
        chunk,
    };
}

function makeDeps(overrides = {}) {
    return {
        appState: {
            isProcessingDoc: false,
            processDocController: null,
            config: { trackChangesEnabled: true, lineDiffEnabled: false, commentGranularity: 0 },
            ...overrides,
        },
        log: jest.fn(),
    };
}

describe('retryFailedChunks', () => {
    beforeEach(() => {
        jest.clearAllMocks();
    });

    it('re-drives the original chunk objects (paragraphs intact) instead of text-only stubs', async () => {
        const originalChunk = {
            id: 'chunk-1',
            paragraphs: [{ text: 'First paragraph.' }, { text: 'Second paragraph.' }],
            overlapBefore: '',
            tokenCount: 10,
            startIndex: 0,
            endIndex: 1,
        };
        processChunksParallel.mockResolvedValue([]);
        applyChunkResults.mockResolvedValue({
            amendmentsApplied: 1, commentsInserted: 0, noChangeCount: 0,
            errors: [], appliedChunkIds: ['chunk-1'], interrupted: false,
        });

        const deps = makeDeps();
        await retryFailedChunks(deps, {
            failedResults: [makeFailedResult(originalChunk)],
            bookmarkMap: new Map([['chunk-1', '_wdp_test1']]),
            backendConfig: { model: 'm' },
            promptShim: {},
        });

        expect(processChunksParallel).toHaveBeenCalledTimes(1);
        const [retryChunks] = processChunksParallel.mock.calls[0];
        expect(retryChunks[0]).toBe(originalChunk);
        expect(retryChunks[0].paragraphs).toEqual([
            { text: 'First paragraph.' },
            { text: 'Second paragraph.' },
        ]);
        expect(deps.appState.isProcessingDoc).toBe(false);
    });

    it('reports still-failed chunks when the orchestrator rejects them again', async () => {
        const chunk = { id: 'chunk-2', paragraphs: [{ text: 'Body.' }], overlapBefore: '' };
        processChunksParallel.mockResolvedValue([
            {
                chunkId: 'chunk-2', chunkIndex: 0, status: 'rejected',
                amendment: null, comment: null, error: 'down', reasoning: null, chunk,
            },
        ]);
        applyChunkResults.mockResolvedValue({
            amendmentsApplied: 0, commentsInserted: 0, noChangeCount: 0,
            errors: [], appliedChunkIds: [], interrupted: false,
        });

        const deps = makeDeps();
        await retryFailedChunks(deps, {
            failedResults: [makeFailedResult(chunk)],
            bookmarkMap: new Map(),
            backendConfig: { model: 'm' },
            promptShim: {},
        });

        expect(deps.log).toHaveBeenCalledWith(expect.stringContaining('still failed'), 'warning');
    });

    it('refuses to run while a document run is already processing', async () => {
        const deps = makeDeps({ isProcessingDoc: true });
        await retryFailedChunks(deps, {
            failedResults: [], bookmarkMap: new Map(), backendConfig: {}, promptShim: {},
        });
        expect(processChunksParallel).not.toHaveBeenCalled();
    });

    // The retry used to raise isProcessingDoc without ever locking the chat
    // input: the composer stayed editable with an active Send button, so a
    // long retry had no visible in-progress state at all.
    describe('busy state', () => {
        function stubSuccess() {
            processChunksParallel.mockResolvedValue([]);
            applyChunkResults.mockResolvedValue({
                amendmentsApplied: 1, commentsInserted: 0, noChangeCount: 0,
                errors: [], appliedChunkIds: ['chunk-1'], interrupted: false,
            });
        }

        it('locks the input for the run and releases it afterwards', async () => {
            stubSuccess();
            const chunk = { id: 'chunk-1', paragraphs: [{ text: 'Body.' }], overlapBefore: '' };
            const setBusy = jest.fn();

            await retryFailedChunks(makeDeps(), {
                failedResults: [makeFailedResult(chunk)],
                bookmarkMap: new Map(),
                backendConfig: { model: 'm' },
                promptShim: {},
                setBusy,
            });

            expect(setBusy.mock.calls.map((c) => c[0])).toEqual([true, false]);
        });

        it('releases the input lock even when the retry throws', async () => {
            processChunksParallel.mockRejectedValue(new Error('backend down'));
            const chunk = { id: 'chunk-1', paragraphs: [{ text: 'Body.' }], overlapBefore: '' };
            const setBusy = jest.fn();

            await retryFailedChunks(makeDeps(), {
                failedResults: [makeFailedResult(chunk)],
                bookmarkMap: new Map(),
                backendConfig: { model: 'm' },
                promptShim: {},
                setBusy,
            });

            expect(setBusy).toHaveBeenLastCalledWith(false);
        });

        it('registers its own controller so cancel() can reach the retry', async () => {
            let seenController = null;
            processChunksParallel.mockImplementation(async (chunks, opts) => {
                // Captured mid-flight: after the run settles the slot is cleared.
                seenController = opts.signal;
                return [];
            });
            applyChunkResults.mockResolvedValue({
                amendmentsApplied: 0, commentsInserted: 0, noChangeCount: 0,
                errors: [], appliedChunkIds: [], interrupted: false,
            });
            const chunk = { id: 'chunk-1', paragraphs: [{ text: 'Body.' }], overlapBefore: '' };
            const deps = makeDeps();

            await retryFailedChunks(deps, {
                failedResults: [makeFailedResult(chunk)],
                bookmarkMap: new Map(),
                backendConfig: { model: 'm' },
                promptShim: {},
            });

            expect(seenController).toBeInstanceOf(AbortSignal);
            // Ownership released on settle, so a later turn is not blocked.
            expect(deps.appState.processDocController).toBeNull();
            expect(deps.appState.isProcessingDoc).toBe(false);
        });

        it('does not clobber a controller still owned by another operation', async () => {
            // isProcessingDoc can be false while a proposal-card apply still
            // owns the controller; overwriting it would make that apply
            // uncancellable.
            const live = new AbortController();
            const deps = makeDeps({ processDocController: live });

            await retryFailedChunks(deps, {
                failedResults: [makeFailedResult({ id: 'chunk-1', paragraphs: [{ text: 'B.' }] })],
                bookmarkMap: new Map(),
                backendConfig: { model: 'm' },
                promptShim: {},
            });

            expect(processChunksParallel).not.toHaveBeenCalled();
            expect(deps.appState.processDocController).toBe(live);
        });
    });

    it('cleans up the retried chunks\' bookmarks after the retry settles', async () => {
        // The original apply() keeps FAILED chunks' bookmarks alive precisely
        // so this retry can target them; once the retry settles they must be
        // removed again (only the retried subset — unrelated bookmarks stay).
        const chunk = { id: 'chunk-3', paragraphs: [{ text: 'Body.' }], overlapBefore: '' };
        processChunksParallel.mockResolvedValue([
            {
                chunkId: 'chunk-3', chunkIndex: 0, status: 'fulfilled',
                amendment: 'New body', comment: null, error: null, chunk,
            },
        ]);
        applyChunkResults.mockResolvedValue({
            amendmentsApplied: 1, commentsInserted: 0, noChangeCount: 0,
            errors: [], appliedChunkIds: ['chunk-3'], interrupted: false,
        });

        const deps = makeDeps();
        await retryFailedChunks(deps, {
            failedResults: [makeFailedResult(chunk)],
            bookmarkMap: new Map([['chunk-3', '_wdp_keep3'], ['chunk-other', '_wdp_other']]),
            backendConfig: { model: 'm' },
            promptShim: {},
        });

        expect(cleanupBookmarks).toHaveBeenCalledTimes(1);
        const [cleanedMap] = cleanupBookmarks.mock.calls[0];
        expect([...cleanedMap.keys()]).toEqual(['chunk-3']);
        expect(cleanedMap.get('chunk-3')).toBe('_wdp_keep3');
    });
});
