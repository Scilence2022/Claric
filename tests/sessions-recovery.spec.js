const { listSessions, loadSession, saveSession, clearAllSessions } = require('../src/taskpane/sessions.js');

const INDEX = 'wordAI.sessions.index';
const PREFIX = 'wordAI.session.';
let store;

beforeEach(() => {
    store = new Map();
    global.localStorage = {
        getItem: jest.fn((key) => store.get(key) ?? null),
        setItem: jest.fn((key, value) => store.set(key, String(value))),
        removeItem: jest.fn((key) => store.delete(key)),
        key: jest.fn((i) => [...store.keys()][i] ?? null),
        get length() { return store.size; },
    };
});

function seed(id) {
    return saveSession([{ role: 'user', text: id }], { id });
}

describe('session recovery consistency', () => {
    test('quota fallback drops evicted entries along with their blobs', () => {
        seed('old-a');
        seed('old-b');
        localStorage.setItem.mockImplementationOnce(() => { throw new Error('QuotaExceededError'); });
        const session = seed('current');
        expect(listSessions().map((entry) => entry.id)).toEqual(['current']);
        expect(loadSession('current')).toEqual(session);
        expect(loadSession('old-a')).toBeNull();
        expect(loadSession('old-b')).toBeNull();
    });

    test.each([false, true])('failed retry keeps no dangling index (updating existing: %s)', (updating) => {
        seed('old');
        const prior = updating ? seed('current') : null;
        localStorage.setItem.mockImplementation((key, value) => {
            if (key === `${PREFIX}current`) throw new Error('QuotaExceededError');
            store.set(key, String(value));
        });
        expect(() => seed('current')).toThrow('QuotaExceededError');
        const entries = listSessions();
        expect(entries.map((entry) => entry.id)).toEqual(updating ? ['current'] : []);
        for (const entry of entries) expect(loadSession(entry.id)).not.toBeNull();
        expect(loadSession('current')).toEqual(prior);
        expect(loadSession('old')).toBeNull();
    });

    test('fallback index write failure is surfaced instead of reporting a saved session', () => {
        seed('old');
        localStorage.setItem.mockImplementation(() => { throw new Error('storage disabled'); });
        expect(() => seed('current')).toThrow(/history index could not be updated/);
        expect(loadSession('current')).toBeNull();
    });

    test('ordinary final index write failure remains visible', () => {
        localStorage.setItem.mockImplementation((key, value) => {
            if (key === INDEX) throw new Error('QuotaExceededError');
            store.set(key, String(value));
        });
        expect(() => seed('current')).toThrow(/history index could not be updated/);
        expect(loadSession('current')).not.toBeNull();
    });

    test('mixed bad index entries do not crash sorting or subsequent saves', () => {
        store.set(INDEX, JSON.stringify([
            null, true, 3, 'bad', [], {}, { id: null }, { id: '' },
            { id: 'bad-date', updatedAt: { localeCompare: 'not callable' } },
            { id: 'numeric-date', updatedAt: 42 },
            { id: 'valid', updatedAt: '2026-01-01T00:00:00Z' },
        ]));
        expect(listSessions().map((entry) => entry.id)).toEqual(['valid', 'bad-date', 'numeric-date']);
        expect(() => seed('current')).not.toThrow();
    });

    test('unavailable storage reads fail closed', () => {
        localStorage.getItem.mockImplementation(() => { throw new Error('SecurityError'); });
        expect(listSessions()).toEqual([]);
        expect(loadSession('missing')).toBeNull();
    });
});

describe('session namespace cleanup', () => {
    test.each([null, '{broken', 'null', '[null,3,{}]'])('clears orphans with index %p without touching other modules', (index) => {
        store.set('wordAI.config', 'config');
        store.set(`${PREFIX}orphan-a`, 'broken blob');
        store.set('wordAI.prompts', 'prompts');
        store.set(`${PREFIX}orphan-b`, '{}');
        store.set('wordAI.sessions.other', 'unrelated');
        store.set('wordAI.sessionish', 'unrelated');
        if (index !== null) store.set(INDEX, index);
        clearAllSessions();
        expect([...store.entries()]).toEqual([
            ['wordAI.config', 'config'], ['wordAI.prompts', 'prompts'],
            ['wordAI.sessions.other', 'unrelated'], ['wordAI.sessionish', 'unrelated'],
        ]);
    });

    test('one removal failure does not prevent cleaning other session keys and the index', () => {
        store.set(`${PREFIX}a`, '{}');
        store.set(`${PREFIX}b`, '{}');
        store.set(INDEX, '[]');
        localStorage.removeItem.mockImplementation((key) => {
            if (key === `${PREFIX}b`) throw new Error('storage failure');
            store.delete(key);
        });
        expect(() => clearAllSessions()).not.toThrow();
        expect([...store.keys()]).toEqual([`${PREFIX}b`]);
    });

    test('enumeration failure still attempts index cleanup', () => {
        store.set(INDEX, '[]');
        localStorage.key.mockImplementation(() => { throw new Error('storage failure'); });
        expect(() => clearAllSessions()).not.toThrow();
        expect(store.has(INDEX)).toBe(false);
    });
});
