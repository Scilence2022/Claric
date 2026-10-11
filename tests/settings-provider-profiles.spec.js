/** @jest-environment jsdom */
const fs = require('fs');
const path = require('path');
jest.mock('../src/lib/llm-client.js', () => ({ testConnection: jest.fn() }));

let appState, defaultConfig, normalizeConfig, getActiveBackendConfig, getActiveImageConfig;
let initSettings, updateUIFromConfig, testConnection;
const flush = async () => { await Promise.resolve(); await Promise.resolve(); };
const el = (id) => document.getElementById(id);
function change(id, value) {
    el(id).value = value;
    el(id).dispatchEvent(new Event('change'));
}
function type(id, value) {
    el(id).value = value;
    el(id).dispatchEvent(new Event('input'));
}
function models(id) {
    return [...el(id).querySelectorAll('[role="option"]')].map((item) => item.textContent);
}

beforeEach(() => {
    jest.useFakeTimers();
    jest.resetModules();
    ({ appState, defaultConfig, normalizeConfig, getActiveBackendConfig, getActiveImageConfig } = require('../src/taskpane/app-state.js'));
    ({ initSettings, updateUIFromConfig } = require('../src/taskpane/ui/settings-view.js'));
    ({ testConnection } = require('../src/lib/llm-client.js'));
    localStorage.clear();
    const html = fs.readFileSync(path.join(__dirname, '../src/taskpane/taskpane.html'), 'utf8');
    document.body.innerHTML = html.slice(html.indexOf('<body>') + 6, html.indexOf('</body>'));
    appState.config = defaultConfig();
    global.fetch = jest.fn(async () => ({ ok: false }));
    testConnection.mockImplementation(async (config) => ({ models: [{ id: `${config.apiKey || 'empty'}-model` }] }));
    jest.spyOn(console, 'log').mockImplementation(() => {});
    initSettings();
});
afterEach(() => { jest.clearAllTimers(); jest.useRealTimers(); jest.restoreAllMocks(); });

test('two keys on one chat provider retain independent settings, catalogs and request routing after reload', async () => {
    change('backendSelect', 'zhongkeyu');
    type('apiKey', 'personal');
    type('modelSelect', 'personal-model');
    type('chatProfileName', 'Personal');
    el('settingsSaveBtn').click();
    await flush();
    el('chatProfileAddBtn').click();
    expect(el('apiKey').value).toBe('');
    type('apiKey', 'team');
    type('endpointUrl', '/team');
    type('modelSelect', 'team-model');
    type('chatProfileName', 'Team');
    change('temperatureInput', '0.4');
    await flush();
    expect(getActiveBackendConfig()).toMatchObject({ provider: 'zhongkeyu', apiKey: 'team', model: 'team-model', url: '/team', temperature: 0.4 });
    el('modelSelect').focus();
    expect(models('modelDropdown')).toEqual(['team-model']);

    // Leave a probe pending to inspect the cached suggestions on switching.
    testConnection.mockReturnValueOnce(new Promise(() => {}));
    change('chatProfileSelect', 'default');
    el('modelSelect').click();
    expect(models('modelDropdown')).toEqual(['personal-model']);
    expect(el('chatProfileName').value).toBe('Personal');
    expect(el('apiKey').value).toBe('personal');
    expect(el('endpointUrl').value).toBe('/zhongkeyu');
    expect(el('temperatureInput').value).toBe('1');
    change('chatProfileSelect', 'key-2');
    await flush();
    appState.config = normalizeConfig(defaultConfig(), JSON.parse(localStorage.getItem('wordAI.config')));
    updateUIFromConfig();
    expect(el('chatProfileSelect').value).toBe('key-2');
    expect(el('chatProfileName').value).toBe('Team');
    expect(getActiveBackendConfig().apiKey).toBe('team');
    expect(testConnection.mock.calls.at(-1)[0]).not.toHaveProperty('profiles');
});

test('immediate provider switching and pending debounced saves preserve the old form', async () => {
    change('backendSelect', 'deepseek');
    type('apiKey', 'last-keystroke');
    type('modelSelect', 'custom-model');
    change('backendSelect', 'glm');
    jest.advanceTimersByTime(400);
    await flush();
    change('backendSelect', 'deepseek');
    expect(el('apiKey').value).toBe('last-keystroke');
    expect(el('modelSelect').value).toBe('custom-model');
    expect(appState.config.providers.glm.apiKey).toBe('');
});

test('credential edits immediately clear suggestions and drop the previous key response', async () => {
    change('backendSelect', 'deepseek');
    type('apiKey', 'first');
    el('refreshModelsBtn').click();
    await flush();
    el('modelSelect').focus();
    expect(models('modelDropdown')).toContain('first-model');
    let finish;
    testConnection.mockReturnValueOnce(new Promise((resolve) => { finish = resolve; }));
    el('refreshModelsBtn').click();
    type('apiKey', 'second');
    expect(models('modelDropdown')).toEqual([]);
    finish({ models: [{ id: 'stale-model' }] });
    await flush();
    expect(models('modelDropdown')).toEqual([]);
    el('refreshModelsBtn').click();
    await flush();
    el('modelSelect').click();
    expect(models('modelDropdown')).toContain('second-model');
});

test('late chat responses cannot overwrite another profile with the same model name', async () => {
    change('backendSelect', 'deepseek');
    type('apiKey', 'first');
    let finish;
    testConnection.mockReturnValueOnce(new Promise((resolve) => { finish = resolve; }));
    el('refreshModelsBtn').click();
    el('chatProfileAddBtn').click();
    type('apiKey', 'second');
    el('settingsSaveBtn').click();
    await flush();
    finish({ models: [{ id: 'stale-model' }] });
    await flush();
    el('modelSelect').focus();
    expect(models('modelDropdown')).toContain('second-model');
    expect(models('modelDropdown')).not.toContain('stale-model');
});

test('image profiles isolate keys and model lists, persist pending input on switching and remove credentials', async () => {
    change('imageProviderSelect', 'zhongkeyu');
    await flush();
    el('imageGenEnabledCheckbox').checked = true;
    el('imageGenEnabledCheckbox').dispatchEvent(new Event('change'));
    await flush();
    type('imageApiKey', 'first');
    type('imageProfileName', 'Personal images');
    testConnection.mockResolvedValueOnce({ models: [{ id: 'first-image' }] });
    el('refreshImageModelsBtn').click();
    await flush();
    el('imageProfileAddBtn').click();
    await flush();
    type('imageApiKey', 'second');
    type('imageModelInput', 'second-image');
    type('imageProfileName', 'Team images');
    let finish;
    testConnection.mockReturnValueOnce(new Promise((resolve) => { finish = resolve; }));
    el('refreshImageModelsBtn').click();
    change('imageProfileSelect', 'default'); // before debounced save fires
    await flush();
    finish({ models: [{ id: 'stale-image' }] });
    await flush();
    el('imageModelInput').focus();
    expect(models('imageModelDropdown')).toContain('first-image');
    expect(models('imageModelDropdown')).not.toContain('stale-image');
    expect(getActiveImageConfig().apiKey).toBe('first');
    change('imageProfileSelect', 'key-2');
    expect(getActiveImageConfig()).toMatchObject({ apiKey: 'second', model: 'second-image' });
    expect(el('imageProfileName').value).toBe('Team images');
    el('imageProfileRemoveBtn').click();
    expect(getActiveImageConfig().apiKey).toBe('first');
    expect(localStorage.getItem('wordAI.config')).not.toContain('second');
    expect(el('imageProfileRemoveBtn').disabled).toBe(true);
});

test('image model refresh drops superseded responses and invalidates catalogs after endpoint edits', async () => {
    change('imageProviderSelect', 'zhongkeyu');
    await flush();
    let finish;
    testConnection.mockReturnValueOnce(new Promise((resolve) => { finish = resolve; }));
    el('refreshImageModelsBtn').click();
    testConnection.mockResolvedValueOnce({ models: [{ id: 'fresh-image' }] });
    el('refreshImageModelsBtn').click();
    await flush();
    finish({ models: [{ id: 'stale-image' }] });
    await flush();
    el('imageModelInput').focus();
    expect(models('imageModelDropdown')).toContain('fresh-image');
    expect(models('imageModelDropdown')).not.toContain('stale-image');
    type('imageEndpointUrl', '/new-endpoint');
    expect(models('imageModelDropdown')).toEqual([]);
});
