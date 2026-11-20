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
        getModelMenuState: () => ({
            models: ['alpha'],
            currentModel: appState.config.providers[appState.config.backend].model,
            thinking: { options: [{ value: 'default', label: 'Default' }, { value: 'high', label: 'High' }], current: 'default' },
        }),
        onPickModel: (id) => settings.applyChatSelection({ model: id }),
        onPickThinkingLevel: (value) => settings.applyChatSelection({ thinkingLevel: value }),
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
    // The model pill opens the popover; picking a model commits it through
    // the lazy settings chunk without opening the panel.
    const model = document.getElementById('modelPill');
    model.focus();
    model.click();
    const menu = document.getElementById('modelMenu');
    expect(menu.hidden).toBe(false);
    menu.querySelector('[role="option"]').click();
    for (let i = 0; i < 6; i++) await Promise.resolve();
    expect(menu.hidden).toBe(true);
    const backend = appState.config.backend;
    expect(appState.config.providers[backend].model).toBe('alpha');
    expect(JSON.parse(localStorage.getItem('wordAI.config')).providers[backend].model).toBe('alpha');
    expect(document.getElementById('settingsOverlay').hidden).toBe(true);
    expect(testConnection).toHaveBeenCalledTimes(1);
    // The popover footer keeps Settings reachable from the composer.
    model.click();
    expect(menu.hidden).toBe(false);
    menu.querySelector('.model-menu-footer button').click();
    await Promise.resolve();
    expect(document.activeElement.id).toBe('settingsCloseBtn');
    expect(document.querySelectorAll('.settings-resize-handle')).toHaveLength(handles);
    expect(testConnection).toHaveBeenCalledTimes(1);
    document.getElementById('settingsCloseBtn').click();
    expect(document.activeElement).toBe(model);
    expect(log).not.toHaveBeenCalled();
});
