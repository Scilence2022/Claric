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
    test.each([false, true])('failed write keeps the previous index and chats intact (updating: %s)', (updating) => {
        const older = seed('old');
        const prior = updating ? seed('current') : null;
        const index = store.get(INDEX);
        localStorage.setItem.mockImplementationOnce(() => { throw new Error('QuotaExceededError'); });

        expect(() => saveSession([{ role: 'user', text: 'new content' }], { id: 'current' })).toThrow('QuotaExceededError');
        expect(store.get(INDEX)).toBe(index);
        expect(loadSession('old')).toEqual(older);
        expect(loadSession('current')).toEqual(prior);
    });

    test('commits an updated legacy chat before dropping its old blob', () => {
        const oldSession = { id: 'old', title: 'Previous version', messages: [] };
        store.set(INDEX, JSON.stringify([{ id: 'old', title: oldSession.title }, { id: 'other' }]));
        store.set(`${PREFIX}old`, JSON.stringify(oldSession));
        store.set(`${PREFIX}other`, JSON.stringify({ id: 'other', messages: [] }));

        localStorage.setItem.mockImplementationOnce(() => { throw new Error('QuotaExceededError'); });
        expect(() => seed('old')).toThrow('QuotaExceededError');
        expect(loadSession('old')).toEqual(oldSession);

        const updated = seed('old');
        expect(updated.title).toBe(oldSession.title);
        expect(loadSession('old')).toEqual(updated);
        expect(loadSession('other')).toMatchObject({ id: 'other' });
        expect(store.has(`${PREFIX}old`)).toBe(false);
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
