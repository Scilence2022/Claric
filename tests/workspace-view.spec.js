/** @jest-environment jsdom */
import { createWorkspaceView } from '../src/taskpane/ui/workspace-view.js';
import { createWorkspaceAccess } from '../src/taskpane/workspace-access.js';
const fs = require('fs');
const path = require('path');
const identity = { workspaceId: 'w', documentId: 'owner', instanceId: 'owner-1', displayName: 'Draft.docx' };
let view;
beforeEach(() => {
    localStorage.clear();
    document.body.innerHTML = fs.readFileSync(path.join(__dirname, '../src/taskpane/taskpane.html'), 'utf8').split('<body>')[1].split('</body>')[0];
});
afterEach(() => view?.dispose());

test('keeps settings outside chat, opens a document-first panel, and restores focus on Escape', () => {
    view = createWorkspaceView({ identity, access: createWorkspaceAccess({ identity }) });
    const button = document.getElementById('workspaceBtn');
    expect(document.getElementById('settingsPageWorkspace').contains(document.getElementById('crossDocumentUrl'))).toBe(true);
    expect(document.getElementById('workspacePanel').contains(document.getElementById('crossDocumentUrl'))).toBe(false);
    expect(document.querySelector('.app > .cross-document-settings')).toBeNull();
    button.focus(); button.click();
    expect(document.getElementById('workspaceOverlay').hidden).toBe(false);
    expect(document.getElementById('workspaceDocuments').hidden).toBe(false);
    expect(document.getElementById('workspaceEmpty').hidden).toBe(false);
    document.getElementById('workspaceTabEvidence').click();
    expect(document.getElementById('workspaceEvidence').hidden).toBe(false);
    document.getElementById('workspacePanel').dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    expect(document.getElementById('workspaceOverlay').hidden).toBe(true);
    expect(document.activeElement).toBe(button);
    button.click();
    expect(document.getElementById('workspaceDocuments').hidden).toBe(false);
});

test('lists live instances safely, displays peer restrictions and applies incoming permission changes', async () => {
    const policy = createWorkspaceAccess({ identity });
    const change = jest.fn(async (...args) => policy.update(...args));
    view = createWorkspaceView({ identity, access: policy, onAccessChange: change, clock: () => 1000 });
    view.update(new Map([
        ['peer-1', { workspaceId: 'w', documentId: 'peer', instanceId: 'peer-1', title: '<img src=x>', expiresAt: 2000, access: { readContext: false } }],
        ['old', { documentId: 'old', instanceId: 'old', expiresAt: 999 }],
    ]), true);
    expect(document.getElementById('workspaceListCount').textContent).toBe('1');
    const list = document.getElementById('workspaceDocumentList');
    expect(list.querySelector('img')).toBeNull();
    expect(list.textContent).toContain('<img src=x>');
    expect([...list.querySelectorAll('button')].find((button) => button.textContent === 'Read context').disabled).toBe(true);
    list.querySelector('input').click();
    await Promise.resolve(); await Promise.resolve();
    expect(change).toHaveBeenCalledWith('peer', 'readContext', false);
    expect(policy.allows({ workspaceId: 'w', documentId: 'peer' }, 'readContext')).toBe(false);
    view.update(new Map(), false);
    expect(document.getElementById('workspaceCount').hidden).toBe(true);
    expect(document.getElementById('workspaceEmpty').hidden).toBe(false);
});
