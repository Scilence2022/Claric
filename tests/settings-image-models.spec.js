/** @jest-environment jsdom */

/**
 * Image model Refresh + combobox. The button queries the SELECTED IMAGE
 * provider's own endpoint ({url}{apiPath}/models — not the chat backend's),
 * filters multi-model gateways' catalogs down to image-ish ids (falling back
 * to the full list when the heuristic matches nothing), caches per provider,
 * and never lets a stale response land on the wrong provider.
 *
 * settings-view keeps provider/model caches at module scope, so each test
 * resets the module registry for a clean slate. Note: changing the image
 * provider auto-saves, which fires a CHAT connection probe — it consumes
 * whatever the llm-client mock has queued, so tests flush after switching
 * before queueing a refresh response.
 */

const fs = require('fs');
const path = require('path');
jest.mock('../src/lib/llm-client.js', () => ({ testConnection: jest.fn() }));

let testConnection;
let initSettings;
let appState;
let defaultConfig;

function loadImageApp() {
    jest.resetModules();
    ({ testConnection } = require('../src/lib/llm-client.js'));
    ({ initSettings } = require('../src/taskpane/ui/settings-view.js'));
    ({ appState, defaultConfig } = require('../src/taskpane/app-state.js'));

    const html = fs.readFileSync(path.join(__dirname, '../src/taskpane/taskpane.html'), 'utf8');
    document.body.innerHTML = html.slice(html.indexOf('<body>') + 6, html.indexOf('</body>'));
    appState.config = defaultConfig();
    global.fetch = jest.fn(async () => ({ ok: false }));
    // Default for the chat re-probe that saveSettings() triggers on provider
    // switches and picks; each refresh overrides with mockResolvedValueOnce.
    testConnection.mockResolvedValue({ connected: true, models: [{ id: 'chat-model' }] });
    initSettings();
}

async function selectImageProvider(id) {
    const select = document.getElementById('imageProviderSelect');
    select.value = id;
    select.dispatchEvent(new Event('change'));
    await flush(); // let the switch's auto-save chat probe consume the default mock
}

const flush = async () => { await Promise.resolve(); await Promise.resolve(); };

function dropdownIds() {
    return [...document.getElementById('imageModelDropdown').querySelectorAll('[role="option"]')]
        .map((o) => o.textContent);
}

beforeEach(loadImageApp);

describe('image model refresh button', () => {
    test('queries the image entry endpoint and filters a gateway catalog to image models', async () => {
        await selectImageProvider('zhongkeyu');
        // Type through the real input path: sync commits the visible form to
        // the provider entry, exactly like a user pasting a key.
        const keyInput = document.getElementById('imageApiKey');
        keyInput.value = 'zk-key';
        keyInput.dispatchEvent(new Event('input'));
        testConnection.mockResolvedValueOnce({
            connected: true,
            models: [{ id: 'gpt-5.1' }, { id: 'gpt-image-1' }, { id: 'glm-5.3-flash' }, { id: 'Kwai-Kolors/Kolors' }],
        });

        document.getElementById('refreshImageModelsBtn').click();
        await flush();

        expect(testConnection).toHaveBeenCalledWith({ url: '/zhongkeyu', apiPath: '/v1', apiKey: 'zk-key' });
        document.getElementById('imageModelInput').focus();
        expect(dropdownIds()).toEqual(['gpt-image-1', 'Kwai-Kolors/Kolors']);
        expect(document.getElementById('imageModelInput').getAttribute('aria-expanded')).toBe('true');
    });

    test('falls back to the full list when nothing looks image-capable', async () => {
        await selectImageProvider('zhongkeyu');
        testConnection.mockResolvedValueOnce({ connected: true, models: [{ id: 'foo-model' }, { id: 'bar-model' }] });

        document.getElementById('refreshImageModelsBtn').click();
        await flush();

        document.getElementById('imageModelInput').focus();
        // The configured preset model is prepended even when unlisted.
        expect(dropdownIds()).toEqual(['gpt-image-1', 'foo-model', 'bar-model']);
    });

    test('keeps the configured model selectable when the provider no longer lists it', async () => {
        await selectImageProvider('zhongkeyu');
        const modelInput = document.getElementById('imageModelInput');
        modelInput.value = 'zzz-custom-image';
        modelInput.dispatchEvent(new Event('input'));
        testConnection.mockResolvedValueOnce({ connected: true, models: [{ id: 'dall-e-3' }] });

        document.getElementById('refreshImageModelsBtn').click();
        await flush();

        document.getElementById('imageModelInput').focus();
        const options = [...document.getElementById('imageModelDropdown').querySelectorAll('[role="option"]')];
        expect(options.map((o) => o.textContent)).toEqual(['zzz-custom-image', 'dall-e-3']);
        expect(options[0].classList.contains('current')).toBe(true);
    });

    test('keyboard pick commits to the selected provider entry and persists', async () => {
        await selectImageProvider('zhongkeyu');
        testConnection.mockResolvedValueOnce({ connected: true, models: [{ id: 'gpt-image-1' }, { id: 'dall-e-3' }] });

        document.getElementById('refreshImageModelsBtn').click();
        await flush();

        const input = document.getElementById('imageModelInput');
        input.focus();
        input.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowDown', bubbles: true }));
        expect(input.getAttribute('aria-activedescendant')).toBe('image-model-option-0');
        input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
        await flush();

        expect(input.value).toBe('gpt-image-1');
        expect(input.getAttribute('aria-expanded')).toBe('false');
        const saved = JSON.parse(localStorage.getItem('wordAI.config'));
        expect(saved.imageGeneration.providers.zhongkeyu.model).toBe('gpt-image-1');
        expect(saved.imageGeneration.provider).toBe('zhongkeyu');
    });

    test('caches per provider and shows the cached list after switching back', async () => {
        await selectImageProvider('zhongkeyu');
        testConnection.mockResolvedValueOnce({ connected: true, models: [{ id: 'gpt-image-1' }] });
        document.getElementById('refreshImageModelsBtn').click();
        await flush();

        await selectImageProvider('openai');
        testConnection.mockResolvedValueOnce({ connected: true, models: [{ id: 'gpt-image-1' }, { id: 'dall-e-3' }] });
        document.getElementById('refreshImageModelsBtn').click();
        await flush();

        // Back to zhongkeyu without a third fetch: the cached list returns.
        await selectImageProvider('zhongkeyu');
        document.getElementById('imageModelInput').focus();
        expect(dropdownIds()).toEqual(['gpt-image-1']);
        const callsPerProvider = testConnection.mock.calls.map((c) => c[0].url);
        expect(callsPerProvider.filter((url) => url === '/zhongkeyu')).toHaveLength(1);
        expect(callsPerProvider.filter((url) => url === '/openai')).toHaveLength(1);
    });

    test('drops a response that lands after the provider switched', async () => {
        await selectImageProvider('zhongkeyu');
        let resolveStale;
        testConnection.mockReturnValueOnce(new Promise((resolve) => { resolveStale = resolve; }));
        document.getElementById('refreshImageModelsBtn').click();

        await selectImageProvider('openai');
        resolveStale({ connected: true, models: [{ id: 'stale-image-model' }] });
        await flush();

        document.getElementById('imageModelInput').focus();
        const dropdown = document.getElementById('imageModelDropdown');
        expect(dropdown.querySelectorAll('[role="option"]')).toHaveLength(0);
        expect(dropdown.textContent).toContain('No models loaded');
    });

    test('surfaces failures in the activity log instead of throwing', async () => {
        await selectImageProvider('zhongkeyu');
        testConnection.mockRejectedValueOnce(new Error('HTTP 401: bad key'));

        document.getElementById('refreshImageModelsBtn').click();
        await flush();

        const logs = document.getElementById('logs');
        expect(logs.lastElementChild.className).toBe('log-error');
        expect(logs.lastElementChild.textContent).toContain('image model refresh failed');
    });

    test('asks for an endpoint URL before fetching when none is set', async () => {
        await selectImageProvider('zhongkeyu');
        // Clear through the real input path: sync writes the visible value
        // into the provider entry, exactly like a user deleting the URL.
        const endpoint = document.getElementById('imageEndpointUrl');
        endpoint.value = '';
        endpoint.dispatchEvent(new Event('input'));
        testConnection.mockClear();

        document.getElementById('refreshImageModelsBtn').click();
        await flush();

        expect(testConnection).not.toHaveBeenCalled();
        expect(document.getElementById('logs').lastElementChild.className).toBe('log-warning');
    });
});
