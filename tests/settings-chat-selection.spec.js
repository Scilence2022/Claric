/** @jest-environment jsdom */

/**
 * Specs for settings-view.applyChatSelection — the commit path behind the
 * composer's model popover: config + persisted state + settings-form sync,
 * thinking-level resolution against the model's capability profile, and no
 * connection re-probe (endpoint and key are unchanged by a pick).
 */

const fs = require('fs');
const path = require('path');
jest.mock('../src/lib/llm-client.js', () => ({ testConnection: jest.fn() }));

// Pure module — identity across jest.resetModules() is irrelevant.
const { getModelCapabilities } = require('../src/lib/model-capabilities.js');

let appState, defaultConfig, getActiveBackendConfig;
let initSettings, applyChatSelection, testConnection;
const flush = async () => { await Promise.resolve(); await Promise.resolve(); };
const el = (id) => document.getElementById(id);
function type(id, value) {
    el(id).value = value;
    el(id).dispatchEvent(new Event('input'));
}

beforeEach(() => {
    jest.useFakeTimers();
    jest.resetModules();
    ({ appState, defaultConfig, getActiveBackendConfig } = require('../src/taskpane/app-state.js'));
    ({ initSettings, applyChatSelection } = require('../src/taskpane/ui/settings-view.js'));
    ({ testConnection } = require('../src/lib/llm-client.js'));
    localStorage.clear();
    const html = fs.readFileSync(path.join(__dirname, '../src/taskpane/taskpane.html'), 'utf8');
    document.body.innerHTML = html.slice(html.indexOf('<body>') + 6, html.indexOf('</body>'));
    appState.config = defaultConfig();
    global.fetch = jest.fn(async () => ({ ok: false }));
    testConnection.mockImplementation(async (config) => ({ models: [{ id: `${config.apiKey || 'empty'}-model` }] }));
    jest.spyOn(console, 'log').mockImplementation(() => {});
});

afterEach(() => { jest.clearAllTimers(); jest.useRealTimers(); jest.restoreAllMocks(); });

test('applyChatSelection commits model + thinking level, syncs the form, persists and refreshes the pill', () => {
    const onConfigChanged = jest.fn();
    initSettings({ onConfigChanged });

    applyChatSelection({ model: 'glm-5.3-flash', thinkingLevel: 'high' });

    const backend = appState.config.backend;
    expect(appState.config.providers[backend].model).toBe('glm-5.3-flash');
    expect(appState.config.providers[backend].thinkingLevel).toBe('high');
    // The form mirrors the commit so a later form-driven save cannot
    // resurrect the previous values.
    expect(el('modelSelect').value).toBe('glm-5.3-flash');
    const levels = [...el('thinkingLevelSelect').options].map((option) => option.value);
    expect(levels).toContain('high');
    expect(el('thinkingLevelSelect').value).toBe('high');
    expect(JSON.parse(localStorage.getItem('wordAI.config')).providers[backend])
        .toMatchObject({ model: 'glm-5.3-flash', thinkingLevel: 'high' });
    expect(onConfigChanged).toHaveBeenCalled();
    // Endpoint and key did not change, so no connection probe runs.
    expect(testConnection).not.toHaveBeenCalled();
});

test('a later settings save keeps the popover-applied selection', async () => {
    initSettings();
    applyChatSelection({ model: 'kept-model', thinkingLevel: 'low' });
    // An unrelated settings edit (debounced save) must not overwrite the
    // committed model or thinking level with stale form values.
    type('apiKey', 'new-key');
    jest.advanceTimersByTime(400);
    await flush();
    expect(getActiveBackendConfig()).toMatchObject({ model: 'kept-model', thinkingLevel: 'low', apiKey: 'new-key' });
});

test('unsupported thinking levels snap to the capability profile like the settings select', () => {
    applyChatSelection({ thinkingLevel: 'xhigh' });
    const backend = appState.config.backend;
    const entry = appState.config.providers[backend];
    const profile = getModelCapabilities(backend, entry.model);
    const expected = profile.options.some((option) => option.value === 'xhigh') ? 'xhigh' : profile.defaultLevel;
    expect(entry.thinkingLevel).toBe(expected);
    expect(profile.options.some((option) => option.value === entry.thinkingLevel)).toBe(true);
});

test('re-applying the current selection changes nothing', () => {
    applyChatSelection({ model: appState.config.providers[appState.config.backend].model });
    expect(localStorage.getItem('wordAI.config')).toBe(null);
});
