/** @jest-environment jsdom */
import { setSessionScope, getSessionScope, saveSession, loadSession, listSessions, deleteSession, clearAllSessions } from '../src/taskpane/sessions.js';

const a = { workspaceId: 'workspace', documentId: 'a' };
const b = { workspaceId: 'workspace', documentId: 'b' };
const messages = (text) => [{ role: 'user', text }];
beforeEach(() => { localStorage.clear(); setSessionScope(null); });
afterEach(() => setSessionScope(null));

test('keeps legacy history separate without automatically assigning it to an open document', () => {
    saveSession(messages('Legacy'), { id: 'legacy' });
    setSessionScope(a);
    expect(listSessions()).toEqual([]);
    expect(loadSession('legacy')).toBeNull();
    clearAllSessions();
    setSessionScope(null);
    expect(loadSession('legacy').messages[0].text).toBe('Legacy');
});

test('same chat id in different documents cannot overwrite, load, delete, or clear another scope', () => {
    setSessionScope(a);
    saveSession(messages('A'), { id: 'same' });
    setSessionScope(b);
    expect(loadSession('same')).toBeNull();
    saveSession(messages('B'), { id: 'same' });
    expect(loadSession('same').scope).toEqual(b);
    deleteSession('same');
    setSessionScope(a);
    expect(loadSession('same').messages[0].text).toBe('A');
    setSessionScope(b);
    saveSession(messages('B again'), { id: 'other' });
    clearAllSessions();
    expect(listSessions()).toEqual([]);
    setSessionScope(a);
    expect(listSessions()).toHaveLength(1);
});

test('does not alias punctuation-separated workspace and document identities', () => {
    setSessionScope({ workspaceId: 'a.b', documentId: 'c' });
    saveSession(messages('one'), { id: 'same' });
    setSessionScope({ workspaceId: 'a', documentId: 'b.c' });
    expect(loadSession('same')).toBeNull();
    const scope = getSessionScope(); scope.documentId = 'changed';
    expect(getSessionScope().documentId).toBe('b.c');
    expect(() => setSessionScope({ workspaceId: '', documentId: 'b' })).toThrow();
});
