/**
 * Session-storage quota-hardening tests.
 *
 * Regression: staged proposals persist full before/after texts verbatim, so
 * a large document run could exceed MAX_SESSION_BYTES while the trimmer
 * only stripped illustration previewSrc — writeSession then threw
 * QuotaExceededError and the whole session silently vanished from history.
 * The trimmer now degrades gracefully: strip previews → truncate proposal
 * diffs → drop proposals → truncate pathological message text.
 *
 * Also covered: a failed index write must surface (the caller logs it)
 * instead of silently leaving the index and blobs inconsistent.
 */

// localStorage mock (node test environment has no DOM)
const localStorageMock = (() => {
    let store = {};
    return {
        getItem: (key) => (key in store ? store[key] : null),
        setItem: (key, value) => { store[key] = String(value); },
        removeItem: (key) => { delete store[key]; },
        clear: () => { store = {}; },
        get length() { return Object.keys(store).length; },
        key: (i) => Object.keys(store)[i] || null,
    };
})();
global.localStorage = localStorageMock;

const { saveSession, __testing } = require('../src/taskpane/sessions.js');
const { MAX_SESSION_BYTES, SESSION_KEY_PREFIX } = __testing;

function hugeProposal(beforeChars, afterChars) {
    return {
        title: 'Proposed edits',
        state: 'pending',
        previewSrc: null,
        items: [{
            id: 'item-1',
            label: 'Section one',
            before: 'a'.repeat(beforeChars),
            after: 'b'.repeat(afterChars),
            searchText: 'a'.repeat(Math.min(beforeChars, 60)),
        }],
    };
}

describe('session quota hardening', () => {
    beforeEach(() => {
        localStorage.clear();
    });

    test('oversized proposal diffs are trimmed until the session fits the per-session cap', () => {
        const messages = [
            { role: 'user', text: 'Polish the whole document' },
            { role: 'assistant', text: 'Staged 1 section(s).', proposals: [hugeProposal(900_000, 900_000)] },
        ];

        const session = saveSession(messages, { id: 's-quota-1' });
        const raw = localStorage.getItem(`${SESSION_KEY_PREFIX}${session.id}`);

        expect(raw.length).toBeLessThanOrEqual(MAX_SESSION_BYTES);
        // The message survives (metadata preserved) even though the diffs
        // were truncated.
        const stored = JSON.parse(raw);
        expect(stored.messages).toHaveLength(2);
    });

    test('illustration previews are stripped first, before any diff truncation', () => {
        const bigPreview = 'data:image/png;base64,' + 'x'.repeat(1_600_000);
        const messages = [
            { role: 'user', text: 'Design an illustration' },
            { role: 'assistant', text: 'Staged illustration.', proposals: [{ previewSrc: bigPreview, items: [] }] },
        ];

        const session = saveSession(messages, { id: 's-quota-2' });
        const stored = JSON.parse(localStorage.getItem(`${SESSION_KEY_PREFIX}${session.id}`));

        expect(stored.messages[1].proposals[0].previewSrc).toBeNull();
        // No diff truncation needed — the proposal items array is untouched.
        expect(stored.messages[1].proposals[0].items).toEqual([]);
    });

    test('a pathological single message text is truncated rather than losing the session', () => {
        const messages = [
            { role: 'user', text: 'x'.repeat(2_000_000) },
            { role: 'assistant', text: 'done' },
        ];

        const session = saveSession(messages, { id: 's-quota-3' });
        const raw = localStorage.getItem(`${SESSION_KEY_PREFIX}${session.id}`);

        expect(raw).not.toBeNull();
        expect(raw.length).toBeLessThanOrEqual(MAX_SESSION_BYTES);
    });

    test('a failed index write surfaces as a saveSession error (index/blob consistency)', () => {
        const original = localStorage.setItem.bind(localStorage);
        localStorage.setItem = (key, value) => {
            if (key === 'wordAI.sessions.index') {
                const err = new Error('QuotaExceededError');
                err.name = 'QuotaExceededError';
                throw err;
            }
            return original(key, value);
        };
        try {
            expect(() => saveSession([{ role: 'user', text: 'hi' }], { id: 's-quota-4' }))
                .toThrow(/index could not be updated/i);
        } finally {
            localStorage.setItem = original;
        }
    });
});
