/** @jest-environment jsdom */
const fs = require('fs');
const path = require('path');
jest.mock('../src/lib/llm-client.js', () => ({ testConnection: jest.fn() }));
const { testConnection } = require('../src/lib/llm-client.js');
const { createSettingsLoader } = require('../src/taskpane/settings-loader.js');
const { initInputBar } = require('../src/taskpane/ui/input-bar.js');
const { appState, defaultConfig } = require('../src/taskpane/app-state.js');

test('lazy settings initializes the real panel once, probes connection, and opens from both entry controls', async () => {
    const html = fs.readFileSync(path.join(__dirname, '../src/taskpane/taskpane.html'), 'utf8');
    document.body.innerHTML = html.slice(html.indexOf('<body>') + 6, html.indexOf('</body>'));
    appState.config = defaultConfig();
    global.fetch = jest.fn(async () => ({ ok: false }));
    testConnection.mockResolvedValue({ models: [{ id: 'alpha' }] });
    const log = jest.fn();
    const settings = createSettingsLoader({ onConfigChanged: jest.fn(), log });
    const button = document.getElementById('settingsBtn');
    button.addEventListener('click', settings.open);
    initInputBar({
        onSubmit: jest.fn(), onCancel: jest.fn(), getSkills: () => [], onOpenSettings: settings.open,
    });
    const probe = settings.testConnection();
    button.focus();
    button.click();
    await probe;
    expect(testConnection).toHaveBeenCalledTimes(1);
    expect(document.getElementById('settingsOverlay').hidden).toBe(false);
    expect(document.activeElement.id).toBe('settingsCloseBtn');
    const handles = document.querySelectorAll('.settings-resize-handle').length;
    expect(handles).toBe(8);
    document.getElementById('settingsCloseBtn').click();
    expect(document.activeElement).toBe(button);
    const model = document.getElementById('modelPill');
    model.focus();
    model.click();
    await Promise.resolve();
    expect(document.activeElement.id).toBe('settingsCloseBtn');
    expect(document.querySelectorAll('.settings-resize-handle')).toHaveLength(handles);
    expect(testConnection).toHaveBeenCalledTimes(1);
    document.getElementById('settingsCloseBtn').click();
    expect(document.activeElement).toBe(model);
    expect(log).not.toHaveBeenCalled();
});
