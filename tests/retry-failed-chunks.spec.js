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

    it('stages the retry without writing and cleans only confirmed applied bookmarks', async () => {
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
        deps.stageRetryProposal = jest.fn();
        const outcome = await retryFailedChunks(deps, {
            failedResults: [makeFailedResult(chunk)],
            bookmarkMap: new Map([['chunk-3', '_wdp_keep3'], ['chunk-other', '_wdp_other']]),
            backendConfig: { model: 'm' },
            promptShim: {},
        });
        expect(deps.stageRetryProposal).toHaveBeenCalledWith(outcome);
        expect(applyChunkResults).not.toHaveBeenCalled();
        expect(cleanupBookmarks).not.toHaveBeenCalled();
        const signal = new AbortController().signal;
        await outcome.apply(['chunk-3'], { signal });
        expect(applyChunkResults.mock.calls[0][2].signal).toBe(signal);
        expect(cleanupBookmarks).toHaveBeenCalledTimes(1);
        const [cleanedMap] = cleanupBookmarks.mock.calls[0];
        expect([...cleanedMap.keys()]).toEqual(['chunk-3']);
        expect(cleanedMap.get('chunk-3')).toBe('_wdp_keep3');
    });

    it('retains failed anchors and original parameters through repeated retries', async () => {
        const chunk = { id: 'c', paragraphs: [{ text: 'Body' }] };
        const failed = makeFailedResult(chunk);
        processChunksParallel.mockResolvedValue([failed]);
        const documentContext = { definitions: ['term'] };
        const deps = makeDeps();
        const outcome = await retryFailedChunks(deps, { failedResults: [failed],
            bookmarkMap: new Map([['c', '_anchor']]), backendConfig: { model: 'm' }, promptShim: {},
            documentContext, commentInstructions: 'Keep citations', concurrency: 2 });
        expect(applyChunkResults).not.toHaveBeenCalled();
        expect(cleanupBookmarks).not.toHaveBeenCalled();
        await outcome.retryFailed();
        expect(processChunksParallel.mock.calls[1][1]).toMatchObject({ documentContext, commentInstructions: 'Keep citations', concurrency: 2 });
        expect(cleanupBookmarks).not.toHaveBeenCalled();
    });

    it('blocks duplicate generation until dismissal and invalidates the old revision', async () => {
        const chunk = { id: 'c', paragraphs: [{ text: 'old' }] };
        const args = { failedResults: [makeFailedResult(chunk)], bookmarkMap: new Map([['c', '_c']]) };
        const deps = makeDeps();
        let resolve;
        processChunksParallel.mockImplementationOnce(() => new Promise((done) => { resolve = done; }));
        const running = retryFailedChunks(deps, args);
        await retryFailedChunks(deps, args);
        expect(processChunksParallel).toHaveBeenCalledTimes(1);
        const fulfilled = { chunkId: 'c', chunk, status: 'fulfilled', amendment: 'new' };
        resolve([fulfilled]);
        const first = await running;
        await retryFailedChunks(deps, args);
        expect(processChunksParallel).toHaveBeenCalledTimes(1);
        await first.discard();
        expect(cleanupBookmarks).not.toHaveBeenCalled();
        processChunksParallel.mockResolvedValue([fulfilled]);
        const second = await retryFailedChunks(deps, args);
        expect(second.revision).toBeGreaterThan(first.revision);
        await expect(first.apply(['c'])).rejects.toThrow(/no longer active/);
        expect(applyChunkResults).not.toHaveBeenCalled();
    });

    it('consumes applied chunks but preserves partial failures for the original retry link', async () => {
        const chunks = ['a', 'b'].map((id) => ({ id, paragraphs: [{ text: id }] }));
        const args = { failedResults: chunks.map(makeFailedResult), bookmarkMap: new Map([['a', '_a'], ['b', '_b']]) };
        const deps = makeDeps();
        processChunksParallel.mockResolvedValue([
            { chunkId: 'a', chunk: chunks[0], status: 'fulfilled', amendment: 'new' }, makeFailedResult(chunks[1]),
        ]);
        const first = await retryFailedChunks(deps, args);
        applyChunkResults.mockResolvedValue({ appliedChunkIds: ['a'], interrupted: false });
        await first.apply(['a']);
        expect([...cleanupBookmarks.mock.calls[0][0].keys()]).toEqual(['a']);
        processChunksParallel.mockResolvedValue([makeFailedResult(chunks[1])]);
        await retryFailedChunks(deps, args);
        expect(processChunksParallel.mock.calls[1][0].map((c) => c.id)).toEqual(['b']);
        await expect(first.apply(['a'])).rejects.toThrow(/no longer active/);
    });

    it('releases a failed generation for retry without cleaning its anchors', async () => {
        const chunk = { id: 'c', paragraphs: [{ text: 'old' }] };
        const args = { failedResults: [makeFailedResult(chunk)], bookmarkMap: new Map([['c', '_c']]) };
        const deps = makeDeps();
        processChunksParallel.mockRejectedValueOnce(new Error('network failed')).mockResolvedValueOnce([makeFailedResult(chunk)]);
        await retryFailedChunks(deps, args);
        await retryFailedChunks(deps, args);
        expect(processChunksParallel).toHaveBeenCalledTimes(2);
        expect(cleanupBookmarks).not.toHaveBeenCalled();
    });

    it.each(['isProcessing', 'isProcessingSummary', 'chatController'])('refuses retry during %s ownership', async (key) => {
        const deps = makeDeps({ [key]: key === 'chatController' ? new AbortController() : true });
        await retryFailedChunks(deps, { failedResults: [], bookmarkMap: new Map() });
        expect(processChunksParallel).not.toHaveBeenCalled();
    });
});
