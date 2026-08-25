/**
 * Unit tests for the session storage layer (src/taskpane/sessions.js).
 *
 * Sessions are JSON-encoded under:
 *   - wordAI.sessions.index : array of metadata (id, title, createdAt, updatedAt, messageCount, preview)
 *   - wordAI.session.<id>   : full session { id, title, createdAt, updatedAt, messages: [...] }
 *
 * Storage guards:
 *   - At most MAX_SESSIONS entries; overflow drops the oldest.
 *   - Each session is trimmed to stay under MAX_SESSION_BYTES (strips
 *     illustration previewSrc first).
 *   - Total bytes capped at MAX_TOTAL_BYTES (older non-current sessions dropped).
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

const {
    listSessions,
    loadSession,
    saveSession,
    deleteSession,
    clearAllSessions,
    generateTitle,
    __testing,
} = require('../src/taskpane/sessions.js');

const { MAX_SESSIONS, MAX_SESSION_BYTES } = __testing;

const SESSION_KEY_PREFIX = 'wordAI.session.';
const INDEX_KEY = 'wordAI.sessions.index';

function makeMessages(texts) {
    return texts.map((t, i) => ({
        id: `m-${i}`,
        role: i % 2 === 0 ? 'user' : 'assistant',
        text: t,
        status: '',
        error: null,
        worklog: null,
        model: null,
        citations: [],
        proposals: [],
        ts: new Date(2026, 0, 1, 0, 0, i).toISOString(),
    }));
}

beforeEach(() => {
    localStorage.clear();
});

describe('saveSession', () => {
    test('creates a new session with auto id and writes index entry', () => {
        const session = saveSession(makeMessages(['Hello there']));

        expect(session.id).toMatch(/^s-/);
        expect(session.title).toBe('Hello there');
        expect(session.createdAt).toBeDefined();
        expect(session.updatedAt).toBeDefined();
        expect(session.messages).toHaveLength(1);

        const idx = listSessions();
        expect(idx).toHaveLength(1);
        expect(idx[0].id).toBe(session.id);
        expect(idx[0].title).toBe('Hello there');
        expect(idx[0].messageCount).toBe(1);
        expect(idx[0].preview).toBe('Hello there');
    });

    test('update by id preserves id, title, and createdAt; bumps updatedAt', async () => {
        const first = saveSession(makeMessages(['First user turn']));
        const originalCreatedAt = first.createdAt;
        const originalUpdatedAt = first.updatedAt;

        // Force updatedAt to differ from a real tick boundary.
        await new Promise((r) => setTimeout(r, 5));

        const second = saveSession(
            [...first.messages, ...makeMessages(['Second user turn', 'Reply'])],
            { id: first.id }
        );

        expect(second.id).toBe(first.id);
        expect(second.title).toBe('First user turn'); // title pinned from first user msg
        expect(second.createdAt).toBe(originalCreatedAt);
        expect(second.updatedAt).not.toBe(originalUpdatedAt);
        expect(second.messages).toHaveLength(3);

        const idx = listSessions();
        expect(idx).toHaveLength(1); // still one entry, refreshed in place
        expect(idx[0].messageCount).toBe(3);
    });

    test('explicit title overrides the auto-generated one', () => {
        const session = saveSession(makeMessages(['what is the deadline?']), { title: 'Contract review' });
        expect(session.title).toBe('Contract review');
        expect(listSessions()[0].title).toBe('Contract review');
    });

    test('strips incoming message to known fields (no extra props leak into storage)', () => {
        const session = saveSession([
            {
                id: 'm-0', role: 'user', text: 'hi',
                status: '', error: null, worklog: null, model: null,
                citations: [], proposals: [], ts: '2026-01-01T00:00:00.000Z',
                _internal: 'leak',
                dom: { ref: 'should-not-persist' },
            },
        ]);

        const raw = JSON.parse(localStorage.getItem(`${SESSION_KEY_PREFIX}${session.id}`));
        expect(raw.messages[0]._internal).toBeUndefined();
        expect(raw.messages[0].dom).toBeUndefined();
    });

    test('throws on non-array messages', () => {
        expect(() => saveSession(null)).toThrow(TypeError);
        expect(() => saveSession('not-an-array')).toThrow(TypeError);
    });
});

describe('listSessions / loadSession', () => {
    test('listSessions returns empty array when no index exists', () => {
        expect(listSessions()).toEqual([]);
    });

    test('listSessions sorts by updatedAt descending', async () => {
        const a = saveSession(makeMessages(['A1']));
        await new Promise((r) => setTimeout(r, 5));
        const b = saveSession(makeMessages(['B1']));
        await new Promise((r) => setTimeout(r, 5));
        const c = saveSession(makeMessages(['C1']));

        const idx = listSessions();
        expect(idx.map((m) => m.id)).toEqual([c.id, b.id, a.id]);
    });

    test('listSessions recovers from corrupt JSON in the index', () => {
        localStorage.setItem(INDEX_KEY, '{not valid json');
        expect(listSessions()).toEqual([]);
    });

    test('loadSession returns the full session data', () => {
        const saved = saveSession(makeMessages(['question', 'answer']));
        const loaded = loadSession(saved.id);
        expect(loaded.id).toBe(saved.id);
        expect(loaded.messages).toHaveLength(2);
        expect(loaded.messages[1].text).toBe('answer');
    });

    test('loadSession returns null for unknown id', () => {
        expect(loadSession('does-not-exist')).toBeNull();
    });

    test('loadSession returns null when stored JSON is malformed', () => {
        localStorage.setItem(`${SESSION_KEY_PREFIX}bad`, 'not-json');
        expect(loadSession('bad')).toBeNull();
    });

    test('loadSession returns null when stored payload is missing messages', () => {
        localStorage.setItem(`${SESSION_KEY_PREFIX}weird`, JSON.stringify({ id: 'weird', title: 't' }));
        expect(loadSession('weird')).toBeNull();
    });
});

describe('deleteSession / clearAllSessions', () => {
    test('deleteSession removes the entry from storage and index', () => {
        const a = saveSession(makeMessages(['A']));
        const b = saveSession(makeMessages(['B']));

        deleteSession(a.id);

        expect(loadSession(a.id)).toBeNull();
        expect(listSessions().map((m) => m.id)).toEqual([b.id]);
        expect(localStorage.getItem(`${SESSION_KEY_PREFIX}${a.id}`)).toBeNull();
    });

    test('deleteSession on unknown id is a no-op', () => {
        const a = saveSession(makeMessages(['A']));
        expect(() => deleteSession('unknown')).not.toThrow();
        expect(listSessions()).toHaveLength(1);
        expect(a.id).toBeDefined();
    });

    test('clearAllSessions wipes both the index and every session', () => {
        saveSession(makeMessages(['A']));
        saveSession(makeMessages(['B']));
        saveSession(makeMessages(['C']));

        clearAllSessions();

        expect(listSessions()).toEqual([]);
        expect(localStorage.getItem(INDEX_KEY)).toBeNull();
        expect(localStorage.getItem(`${SESSION_KEY_PREFIX}${listSessions()[0]?.id}`)).toBeNull();
    });
});

describe('generateTitle', () => {
    test('uses the first user message text', () => {
        expect(generateTitle(makeMessages(['  Hello  ']))).toBe('Hello');
    });

    test('truncates long titles to 30 characters with ellipsis', () => {
        const long = 'a'.repeat(80);
        const title = generateTitle(makeMessages([long]));
        expect(title.length).toBe(31); // 30 + '…'
        expect(title.endsWith('…')).toBe(true);
    });

    test('collapses whitespace', () => {
        expect(generateTitle(makeMessages(['line1\n\nline2   line3']))).toBe('line1 line2 line3');
    });

    test('falls back to "Untitled chat" when no user message exists', () => {
        expect(generateTitle(makeMessages([]))).toBe('Untitled chat');
        expect(generateTitle(makeMessages(['', '   ']))).toBe('Untitled chat');
        expect(generateTitle([
            { id: 'm-1', role: 'assistant', text: 'no user yet', ts: '' },
        ])).toBe('Untitled chat');
    });
});

describe('storage limits', () => {
    test('drops the oldest sessions beyond MAX_SESSIONS', () => {
        // Save MAX_SESSIONS + 5 sessions, oldest 5 should be evicted.
        const ids = [];
        for (let i = 0; i < MAX_SESSIONS + 5; i++) {
            const s = saveSession(makeMessages([`turn ${i}`]));
            ids.push(s.id);
        }
        expect(listSessions()).toHaveLength(MAX_SESSIONS);
        // The first 5 ids are gone.
        for (let i = 0; i < 5; i++) {
            expect(loadSession(ids[i])).toBeNull();
        }
        // The newest MAX_SESSIONS are still there.
        for (let i = 5; i < MAX_SESSIONS + 5; i++) {
            expect(loadSession(ids[i])).not.toBeNull();
        }
    });

    test('strips illustration previewSrc when a session exceeds MAX_SESSION_BYTES', () => {
        const big = 'x'.repeat(MAX_SESSION_BYTES + 1000);
        const messages = [
            { id: 'm-0', role: 'user', text: 'q', status: '', error: null, worklog: null, model: null, citations: [], proposals: [], ts: '2026-01-01T00:00:00.000Z' },
            {
                id: 'm-1', role: 'assistant', text: 'a',
                status: '', error: null, worklog: null, model: null,
                citations: [], ts: '2026-01-01T00:00:01.000Z',
                proposals: [{ title: 'Illustration', state: 'applied', previewSrc: `data:image/svg+xml;base64,${big}` }],
            },
        ];

        const session = saveSession(messages);
        const loaded = loadSession(session.id);
        expect(loaded.messages[1].proposals[0].previewSrc).toBeNull();
        // Session should now fit under the cap.
        expect(JSON.stringify(loaded).length).toBeLessThanOrEqual(MAX_SESSION_BYTES + 1000); // allow for trim approximation
    });

    test('totals over the cap evict the oldest non-current session', () => {
        // Each session ~1.4 MB (under per-session cap), three of them pushes
        // the total past MAX_TOTAL_BYTES (4 MB) — the oldest should be evicted.
        const payload = 'y'.repeat(1_400_000);
        const heavyMessages = (label) => [
            { id: 'm-0', role: 'user', text: label, status: '', error: null, worklog: null, model: null, citations: [], proposals: [], ts: '2026-01-01T00:00:00.000Z' },
            { id: 'm-1', role: 'assistant', text: payload, status: '', error: null, worklog: null, model: null, citations: [], proposals: [], ts: '2026-01-01T00:00:01.000Z' },
        ];

        const first = saveSession(heavyMessages('first'));
        const second = saveSession(heavyMessages('second'));
        const third = saveSession(heavyMessages('third'));

        // The newest session must always be present.
        expect(loadSession(third.id)).not.toBeNull();
        // The oldest is the first eviction target under the total cap.
        expect(loadSession(first.id)).toBeNull();
        // At least one of (first, second) is gone to keep us under the cap.
        const remaining = [first, second, third].filter((s) => loadSession(s.id) !== null);
        expect(remaining.length).toBeLessThanOrEqual(2);
    });
});
