/**
 * Settings View Module
 *
 * The settings slide-over: a tabbed panel (General, Prompts, Skills, MCP
 * Servers) that can be dragged by its header and resized from any edge or
 * corner once floated. LLM provider settings auto-save (debounced for
 * free-text fields, connection re-tested on change); prompts support
 * per-category list/edit/save/delete/clear plus the save-prompt modal.
 *
 * Element IDs match the pre-refactor markup so the logic stays a straight
 * port; localStorage keys (wordAI.config, wordAI.prompts.*) are unchanged.
 *
 * @module ui/settings-view
 */

import { testConnection as llmTestConnection } from '../../lib/llm-client.js';
import { CATEGORIES } from '../../lib/prompt-manager.js';
import { getProviderPreset } from '../../lib/providers.js';
import { getImageProviderPreset, imageSizesFor, DEFAULT_IMAGE_SIZE } from '../../lib/image-providers.js';
import { testImageConnection } from '../../lib/image-client.js';
import {
    getModelCapabilities,
    resolveThinkingLevel,
    isTemperatureSupported,
    getThinkingHint,
    getTemperatureHint,
} from '../../lib/model-capabilities.js';
import { parseSkillPackage } from '../../lib/skill-package.js';
import { connectMcpServer } from '../../lib/mcp-client.js';
import { importServerPrompts } from '../../lib/mcp-tools.js';
import { TOOL_LOOP_LIMITS } from '../../lib/tool-registry.js';
import { loadImportedSkills, addImportedSkill, removeImportedSkill } from '../../lib/skill-store.js';
import { appState, getActiveBackendConfig, getActiveImageConfig, debounce, persistSettings } from '../app-state.js';
import { BUILTIN_SKILLS, RESERVED_MCP_SKILL } from '../skills.js';
import { addLog, setConnectionStatus } from './status-bar.js';

let _onConfigChanged = () => {};
let _lastFocusedElement = null;
let _availableModels = [];
const _modelsByProvider = new Map();
let _connectionSequence = 0;
// Provider whose values are currently shown in the image form. The select
// changes before the three text inputs are re-rendered, so this separate
// value lets a change handler commit the old form to the old provider first.
let _imageFormProvider = null;

/**
 * Wires the settings slide-over and prompt management UI. Called once at startup.
 *
 * @param {object} deps
 * @param {function()} [deps.onConfigChanged] - Called after config changes (model pill refresh)
 */
export function initSettings({ onConfigChanged } = {}) {
    _onConfigChanged = onConfigChanged || (() => {});

    document.getElementById('settingsBtn').addEventListener('click', openSettings);
    document.getElementById('settingsCloseBtn').addEventListener('click', closeSettings);
    document.getElementById('settingsSaveBtn').addEventListener('click', () => {
        // Explicit Save gives visible feedback (settings are auto-saved too);
        // saveSettings returns true/false so we can show the outcome inline.
        const ok = saveSettings();
        showSaveConfirmation(ok);
    });
    for (const tab of document.querySelectorAll('.settings-tab')) {
        tab.addEventListener('click', () => switchSettingsTab(tab.dataset.tab));
    }
    document.getElementById('settingsOverlay').addEventListener('click', (e) => {
        if (e.target.id === 'settingsOverlay') closeSettings();
    });

    // Escape closes the modal first, then the slide-over.
    document.addEventListener('keydown', (e) => {
        if (e.key !== 'Escape') return;
        const modal = document.getElementById('savePromptModal');
        if (modal && modal.classList.contains('active')) {
            hideSavePromptModal();
            return;
        }
        const overlay = document.getElementById('settingsOverlay');
        if (overlay && !overlay.hasAttribute('hidden')) {
            closeSettings();
        }
    });

    document.getElementById('backendSelect').addEventListener('change', handleBackendSwitch);

    // Auto-save settings on every change. Free-text inputs (URL, API key,
    // model) are debounced so a burst of keystrokes does not fire a save +
    // connection probe per character. The header Save button is an explicit
    // affordance for the same saveSettings path.
    const debouncedSaveSettings = debounce(saveSettings, 400);
    const modelInput = document.getElementById('modelSelect');
    // Model combobox: clicking/focusing shows the full list, typing filters.
    modelInput.addEventListener('input', () => {
        updateGenerationControls(modelInput.value);
        renderModelDropdown(modelInput.value);
        debouncedSaveSettings();
    });
    modelInput.addEventListener('focus', openModelDropdown);
    modelInput.addEventListener('click', openModelDropdown);
    modelInput.addEventListener('blur', () => {
        // Item clicks use mousedown + preventDefault, so blur only means the
        // user genuinely left the field; close on the next tick.
        setTimeout(closeModelDropdown, 150);
    });
    modelInput.addEventListener('keydown', (e) => {
        if (e.key === 'Escape' && !document.getElementById('modelDropdown').hidden) {
            e.stopPropagation(); // close the dropdown, not the settings panel
            closeModelDropdown();
        }
    });
    document.getElementById('refreshModelsBtn').addEventListener('click', () => {
        addLog('Refreshing model list...', 'info');
        testConnectionUI();
    });
    document.getElementById('endpointUrl').addEventListener('input', debouncedSaveSettings);
    document.getElementById('apiKey').addEventListener('input', debouncedSaveSettings);
    document.getElementById('trackChangesCheckbox').addEventListener('change', saveSettings);
    document.getElementById('lineDiffCheckbox').addEventListener('change', saveSettings);
    document.getElementById('thinkingLevelSelect').addEventListener('change', (e) => {
        // The just-picked level drives the hint/temperature refresh; reading
        // the still-unsaved config here would snap the select back.
        updateGenerationControls(undefined, e.target.value);
        saveSettings();
    });
    document.getElementById('temperatureInput').addEventListener('change', saveSettings);
    document.getElementById('docRichnessSelect').addEventListener('change', saveSettings);
    document.getElementById('trackedChangesExtraction').addEventListener('change', saveSettings);
    document.getElementById('commentGranularity').addEventListener('change', saveSettings);
    document.getElementById('includeCommentsInSelectionCheckbox').addEventListener('change', saveSettings);

    // Image generation. The whole block is guarded: reduced DOM fixtures in
    // tests render the LLM section without this one.
    const imageProviderEl = document.getElementById('imageProviderSelect');
    if (imageProviderEl) {
        imageProviderEl.addEventListener('change', () => {
            // The select changes before the text inputs are re-rendered. Commit
            // the old form to the provider it belonged to before switching;
            // otherwise the old key/URL/model would be copied to the new entry.
            syncImageFormToProvider(_imageFormProvider);
            if (appState.config.imageGeneration) {
                appState.config.imageGeneration.provider = imageProviderEl.value;
            }
            updateImageUIFromConfig();
            saveSettings();
        });
        document.getElementById('imageGenEnabledCheckbox').addEventListener('change', saveSettings);
        document.getElementById('imageSizeSelect').addEventListener('change', saveSettings);
        const onImageTextInput = () => {
            // Keep in-memory state current immediately. The debounced call is
            // only for localStorage persistence and the connection side effect;
            // it must not be the sole owner of the user's latest keystroke.
            syncImageFormToProvider();
            debouncedSaveSettings();
        };
        document.getElementById('imageEndpointUrl').addEventListener('input', onImageTextInput);
        document.getElementById('imageModelInput').addEventListener('input', onImageTextInput);
        document.getElementById('imageApiKey').addEventListener('input', onImageTextInput);
        document.getElementById('testImageConnectionBtn').addEventListener('click', handleTestImageConnection);
    }

    // Per-category prompt controls
    for (const category of CATEGORIES) {
        document.getElementById(`promptSelect-${category}`).addEventListener('change', (e) => {
            handleCategoryPromptSelect(category, e.target.value);
        });
        document.getElementById(`savePromptBtn-${category}`).addEventListener('click', () => {
            handleSavePrompt(category);
        });
        document.getElementById(`deletePromptBtn-${category}`).addEventListener('click', () => {
            handleDeletePromptConfirm(category);
        });
        document.getElementById(`resetPromptBtn-${category}`).addEventListener('click', () => {
            handleResetPrompt(category);
        });
    }

    // Modal buttons
    document.getElementById('savePromptConfirmBtn').addEventListener('click', handleSavePromptConfirm);
    document.getElementById('savePromptCancelBtn').addEventListener('click', hideSavePromptModal);

    updateUIFromConfig();
    initSkillImport();
    initMcpServers();
    initMcpStepBudget();
    initPanelGeometry();
    renderAllDropdowns();

    // Restore textarea content from the active prompt of each category.
    for (const category of CATEGORIES) {
        const activePrompt = appState.promptManager.getActivePrompt(category);
        if (activePrompt) {
            document.getElementById(`promptTextarea-${category}`).value = activePrompt.template;
            if (category === 'amendment' && activePrompt.commentInstructions) {
                document.getElementById('commentInstructions').value = activePrompt.commentInstructions;
            }
        }
    }
}

/** Opens the settings slide-over. */
export function openSettings() {
    document.getElementById('settingsOverlay').removeAttribute('hidden');
    // Re-render on every open so the lists reflect edits made elsewhere.
    renderSkillImportList();
    renderMcpServerList();
}

/** Closes the settings slide-over and restores the docked panel layout. */
export function closeSettings() {
    closeModelDropdown();
    resetPanelGeometry();
    document.getElementById('settingsOverlay').setAttribute('hidden', '');
}

/**
 * Activates a settings tab ("general", "prompts", "skills", or "mcp") and
 * shows its page.
 *
 * @param {string} name - The tab's data-tab value
 */
function switchSettingsTab(name) {
    for (const tab of document.querySelectorAll('.settings-tab')) {
        const active = tab.dataset.tab === name;
        tab.classList.toggle('active', active);
        tab.setAttribute('aria-selected', String(active));
    }
    for (const page of document.querySelectorAll('.settings-page')) {
        page.toggleAttribute('hidden', page.id !== `settingsPage${capitalize(name)}`);
    }
}

/**
 * Reads the settings form into appState.config and persists it.
 * Re-tests the connection and refreshes the model pill afterwards.
 * Returns true on success, false when persistSettings throws.
 */
function saveSettings() {
    const config = appState.config;
    const backend = document.getElementById('backendSelect').value;
    const endpointUrl = document.getElementById('endpointUrl').value.trim();
    const apiKey = document.getElementById('apiKey').value.trim();
    const selectedModel = document.getElementById('modelSelect').value;
    const thinkingLevel = document.getElementById('thinkingLevelSelect').value;
    const temperatureValue = Number(document.getElementById('temperatureInput').value);
    const profile = getModelCapabilities(backend, selectedModel || config.providers[backend].model);

    config.backend = backend;
    config.providers[backend].url = endpointUrl || config.providers[backend].url;
    config.providers[backend].apiKey = apiKey;
    // Every provider's model is editable: users can type a model id not
    // present in the refreshable list (e.g. a newly released one).
    config.providers[backend].model = selectedModel || config.providers[backend].model;
    config.providers[backend].thinkingLevel = resolveThinkingLevel(profile, thinkingLevel);
    config.providers[backend].temperature = Number.isFinite(temperatureValue) && temperatureValue >= 0 && temperatureValue <= 2
        ? temperatureValue
        : 1;
    config.trackChangesEnabled = document.getElementById('trackChangesCheckbox').checked;
    config.lineDiffEnabled = document.getElementById('lineDiffCheckbox').checked;
    config.docExtraction = {
        richness: document.getElementById('docRichnessSelect').value
    };
    config.trackedChangesExtraction = document.getElementById('trackedChangesExtraction').checked;
    config.commentGranularity = parseInt(document.getElementById('commentGranularity').value || '0', 10);
    config.includeCommentsInSelection = document.getElementById('includeCommentsInSelectionCheckbox').checked;
    saveImageSettings(config);

    try {
        persistSettings(appState);
        addLog('Settings saved.', 'success');
        _onConfigChanged();
        // Re-test connection with new settings
        testConnectionUI();
        return true;
    } catch (e) {
        addLog(`Failed to save settings: ${e.message}`, 'error');
        return false;
    }
}

/**
 * Shows a transient inline confirmation next to the Save button after an
 * explicit save attempt ("Saved ✓" / "Save failed"). Auto-hides after 2s.
 *
 * @param {boolean} ok - Whether the save succeeded
 */
function showSaveConfirmation(ok) {
    const el = document.getElementById('settingsSaveStatus');
    if (!el) return;
    el.textContent = ok ? 'Saved ✓' : 'Save failed';
    el.classList.toggle('success', ok);
    el.classList.toggle('error', !ok);
    el.hidden = false;
    clearTimeout(_saveStatusTimer);
    _saveStatusTimer = setTimeout(() => {
        el.hidden = true;
    }, 2000);
}

/** Timer handle for the transient Save confirmation. */
let _saveStatusTimer = null;

/**
 * Commits the image text fields currently shown in the form to one provider
 * entry. `provider` is explicit because a select change fires while the text
 * inputs still display the previous provider's values.
 *
 * @param {string} [provider] - Provider represented by the visible fields
 * @param {object} [config] - Config to mutate (defaults to appState.config)
 * @returns {boolean} Whether an entry was updated
 */
export function syncImageFormToProvider(provider, config = appState.config) {
    const providerEl = document.getElementById('imageProviderSelect');
    const image = config && config.imageGeneration;
    if (!providerEl || !image || !image.providers) return false;

    const formProvider = provider || _imageFormProvider || providerEl.value;
    const entry = image.providers[formProvider];
    if (!entry) return false;

    const endpointEl = document.getElementById('imageEndpointUrl');
    const modelEl = document.getElementById('imageModelInput');
    const keyEl = document.getElementById('imageApiKey');
    const sizeEl = document.getElementById('imageSizeSelect');
    // Empty strings are intentional: clearing a custom URL/model must remain
    // cleared instead of silently resurrecting the previous value.
    if (endpointEl) entry.url = endpointEl.value.trim();
    if (modelEl) entry.model = modelEl.value.trim();
    if (keyEl) entry.apiKey = keyEl.value.trim();
    if (sizeEl && imageSizesFor(formProvider).includes(sizeEl.value)) {
        entry.size = sizeEl.value;
    }
    return true;
}

/**
 * Reads the Image Generation form into `config.imageGeneration`.
 *
 * Per-provider fields are written under the SELECTED provider only, mirroring
 * how saveSettings treats `config.providers[backend]`: switching providers must
 * not carry one provider's key or endpoint onto another.
 *
 * @param {object} config - appState.config (mutated in place)
 */
function saveImageSettings(config) {
    const providerEl = document.getElementById('imageProviderSelect');
    // The section is absent in reduced test fixtures; leave config untouched.
    if (!providerEl || !config.imageGeneration) return;

    const provider = providerEl.value;
    const image = config.imageGeneration;
    const enabledEl = document.getElementById('imageGenEnabledCheckbox');
    if (enabledEl) image.enabled = enabledEl.checked;
    if (!image.providers || !image.providers[provider]) return;
    image.provider = provider;
    _imageFormProvider = provider;
    syncImageFormToProvider(provider, config);
}

/**
 * Pushes `config.imageGeneration` into the Image Generation form, including
 * rebuilding the size list for the selected provider.
 */
function updateImageUIFromConfig() {
    const providerEl = document.getElementById('imageProviderSelect');
    const image = appState.config.imageGeneration;
    if (!providerEl || !image) return;

    document.getElementById('imageGenEnabledCheckbox').checked = !!image.enabled;
    providerEl.value = image.provider;
    _imageFormProvider = image.provider;

    const entry = image.providers[image.provider] || {};
    document.getElementById('imageEndpointUrl').value = entry.url || '';
    document.getElementById('imageModelInput').value = entry.model || '';
    document.getElementById('imageApiKey').value = entry.apiKey || '';

    const sizeSelect = document.getElementById('imageSizeSelect');
    sizeSelect.innerHTML = '';
    for (const size of imageSizesFor(image.provider)) {
        const option = document.createElement('option');
        option.value = size;
        option.textContent = size === 'auto' ? 'Auto (provider default)' : size;
        sizeSelect.appendChild(option);
    }
    sizeSelect.value = entry.size || DEFAULT_IMAGE_SIZE;

    updateImageProviderHints();
}

/**
 * Refreshes the endpoint/key hints for the selected image provider, using the
 * preset's own keyHint and staticOk flags (same reasoning as
 * updateProviderHints for chat providers).
 */
function updateImageProviderHints() {
    const providerEl = document.getElementById('imageProviderSelect');
    if (!providerEl) return;
    const preset = getImageProviderPreset(providerEl.value);
    const endpointHint = document.getElementById('imageEndpointHint');
    const keyHint = document.getElementById('imageApiKeyHint');
    if (!preset || !endpointHint || !keyHint) return;

    endpointHint.textContent = preset.staticOk === false
        ? `${preset.label} base URL — the default proxy path works when this add-in is served by its own server; this API allows no direct browser calls.`
        : `${preset.label} base URL`;
    keyHint.textContent = preset.keyHint
        ? `Get a key at ${preset.keyHint}`
        : 'API key for the image endpoint';
}

/**
 * Runs a real (tiny) generation against the current image settings and reports
 * the outcome inline. There is no cheap ping on these APIs — see
 * testImageConnection.
 */
async function handleTestImageConnection() {
    const statusEl = document.getElementById('imageTestStatus');
    const button = document.getElementById('testImageConnectionBtn');
    if (!statusEl || !button) return;

    // Save first so the test uses exactly what the user sees in the form.
    saveImageSettings(appState.config);
    const config = getActiveImageConfig(appState);
    if (!config) {
        statusEl.textContent = 'Enable image generation and set an endpoint plus model first.';
        return;
    }

    button.disabled = true;
    statusEl.textContent = 'Generating a test image...';
    try {
        const result = await testImageConnection(config, addLog);
        statusEl.textContent = result.ok ? `OK — ${result.detail}` : `Failed — ${result.detail}`;
    } catch (error) {
        statusEl.textContent = `Failed — ${error.message}`;
    } finally {
        button.disabled = false;
    }
}

/**
 * Pushes appState.config into the settings form.
 */
export function updateUIFromConfig() {
    const config = appState.config;
    const backendConfig = getActiveBackendConfig(appState);
    const modelSelect = document.getElementById('modelSelect');

    document.getElementById('backendSelect').value = config.backend;
    document.getElementById('endpointUrl').value = backendConfig.url;
    document.getElementById('apiKey').value = backendConfig.apiKey;
    document.getElementById('trackChangesCheckbox').checked = config.trackChangesEnabled;
    document.getElementById('lineDiffCheckbox').checked = config.lineDiffEnabled;

    // Model field stays editable (combobox) for every provider; typed or
    // selected values are saved as-is and the list refreshes on demand.
    modelSelect.value = backendConfig.model || '';
    modelSelect.placeholder = backendConfig.model || 'Type a model name or refresh the list';
    document.getElementById('temperatureInput').value = String(Number.isFinite(backendConfig.temperature) ? backendConfig.temperature : 1);
    // Restore this provider's cached suggestions, then rebuild the
    // capability-aware thinking options before selecting the saved level.
    _availableModels = _modelsByProvider.get(config.backend) || [];
    updateGenerationControls(modelSelect.value, backendConfig.thinkingLevel);

    const richnessSelect = document.getElementById('docRichnessSelect');
    if (richnessSelect && config.docExtraction) {
        richnessSelect.value = config.docExtraction.richness || 'structured';
    }

    document.getElementById('trackedChangesExtraction').checked = !!config.trackedChangesExtraction;
    document.getElementById('commentGranularity').value = String(config.commentGranularity || 0);
    document.getElementById('includeCommentsInSelectionCheckbox').checked = !!config.includeCommentsInSelection;

    updateImageUIFromConfig();
    updateProviderHints();
}

/**
 * Rebuilds the thinking-level dropdown and Temperature state from the
 * capability profile of the currently selected provider/model.
 *
 * The persisted thinking value may name a level another provider supports
 * but this one does not; it is then normalized to the profile's default (or
 * its documented alias) so the visible selection always maps to a request
 * this model understands.
 *
 * @param {string} [modelOverride] - In-flight model input value (typing)
 * @param {string} [preferredLevel] - Level to select instead of the saved one
 */
function updateGenerationControls(modelOverride, preferredLevel) {
    const config = appState.config;
    const backendConfig = config.providers[config.backend];
    const model = (modelOverride !== undefined ? modelOverride : document.getElementById('modelSelect').value)
        || backendConfig.model;
    const profile = getModelCapabilities(config.backend, model);
    const savedLevel = preferredLevel !== undefined ? preferredLevel : backendConfig.thinkingLevel;
    const level = resolveThinkingLevel(profile, savedLevel);

    const select = document.getElementById('thinkingLevelSelect');
    if (select) {
        select.innerHTML = '';
        for (const option of profile.options) {
            const el = document.createElement('option');
            el.value = option.value;
            el.textContent = option.label;
            if (option.description) el.title = option.description;
            select.appendChild(el);
        }
        select.value = level;
    }

    const thinkingHint = document.getElementById('thinkingLevelHint');
    if (thinkingHint) {
        thinkingHint.textContent = getThinkingHint(profile);
    }

    // Temperature is disabled — never silently overwritten — when the
    // profile/level cannot accept it; the stored value survives for the
    // models and levels that do support it.
    const temperatureInput = document.getElementById('temperatureInput');
    const temperatureOk = isTemperatureSupported(profile, level);
    if (temperatureInput) {
        temperatureInput.disabled = !temperatureOk;
    }
    const temperatureHint = document.getElementById('temperatureHint');
    if (temperatureHint) {
        temperatureHint.textContent = getTemperatureHint(profile, level);
    }
}

/**
 * Handles switching between providers: restores the selected provider's saved
 * settings, refreshes hints, and re-tests the connection.
 */
function handleBackendSwitch() {
    appState.config.backend = document.getElementById('backendSelect').value;
    // Any in-flight connection probe belongs to the previous provider now;
    // its results must not repopulate this provider's model list or status.
    _connectionSequence += 1;
    closeModelDropdown();
    updateUIFromConfig();
    saveSettings();
}

/**
 * Updates the endpoint/API-key hint text to describe the selected provider
 * preset (e.g. where to get a DeepSeek key).
 */
function updateProviderHints() {
    const config = appState.config;
    const preset = getProviderPreset(config.backend);
    const endpointHint = document.getElementById('endpointHint');
    if (endpointHint) {
        const isCustom = config.backend === 'custom';
        if (isCustom) {
            endpointHint.textContent = 'Base URL of any OpenAI-compatible server (proxy path or full https:// URL)';
        } else if (preset && preset.staticOk === false && preset.staticHint) {
            // Provider-specific explanation (e.g. OpenAI's missing CORS
            // headers) instead of the local-model mixed-content wording.
            endpointHint.textContent = preset.staticHint;
        } else if (preset && preset.staticOk === false) {
            // Local-model presets default to a same-origin proxy path, which
            // only exists when this add-in is served by the docker/dev
            // server. A statically hosted install (e.g. marketplace) cannot
            // reach http://localhost at all (mixed-content blocking), so
            // point the user at the two real options.
            endpointHint.textContent = `${preset.label} base URL — the default works when this add-in is served by its local server. From a static install (e.g. marketplace), enter an HTTPS ${preset.label} endpoint (CORS enabled, e.g. via OLLAMA_ORIGINS); http://localhost is blocked by Word's WebView.`;
        } else {
            endpointHint.textContent = `Base URL for ${preset ? preset.label : config.backend} — the default adapts to how the add-in is served: direct to the provider on static installs, same-origin proxy path behind the local server`;
        }
    }
    const keyHint = document.getElementById('apiKeyHint');
    if (keyHint && preset && preset.keyHint) {
        keyHint.textContent = `Get an API key at ${preset.keyHint} (leave blank for local backends)`;
    }
}

/**
 * Tests connection to the active LLM backend and populates the model
 * suggestion list. Updates the header status dot.
 *
 * The probe is asynchronous, so it captures the provider id and a sequence
 * number up front: a response that arrives after the user switched provider
 * (or re-saved settings) belongs to the old configuration and is dropped.
 */
export async function testConnectionUI() {
    const backend = appState.config.backend;
    const backendConfig = getActiveBackendConfig(appState);
    const preset = getProviderPreset(backend);
    const backendLabel = preset ? preset.label : backend;
    const sequence = ++_connectionSequence;

    setConnectionStatus('connecting', `${backendLabel}: Connecting...`);

    // A custom provider with no URL yet cannot be probed.
    if (!backendConfig.url) {
        setConnectionStatus('error', `${backendLabel}: enter an endpoint URL`);
        return;
    }

    try {
        const result = await llmTestConnection(backendConfig);
        _modelsByProvider.set(backend, result.models.map((m) => m.id));
        if (sequence !== _connectionSequence || backend !== appState.config.backend) return;
        setConnectionStatus('connected', `${backendLabel}: Connected`);
        addLog(`Connected to ${backendLabel}! Found ${result.models.length} model(s).`, 'success');
        populateModels(result.models);
    } catch (error) {
        if (sequence !== _connectionSequence || backend !== appState.config.backend) return;
        if (error.message && (error.message.includes('401') || error.message.includes('403'))) {
            setConnectionStatus('error', `${backendLabel}: API key required`);
            addLog(`${backendLabel} authentication failed: ${error.message}`, 'error');
        } else {
            setConnectionStatus('error', `${backendLabel}: Connection Error`);
            addLog(`${backendLabel} connection failed: ${error.message}`, 'error');
        }
        console.error('Connection error:', error);
    }
}

/**
 * Populates the model combobox from the provider's models endpoint.
 * The field is a text input backed by a custom dropdown: clicking or focusing
 * the input shows the full list, typing filters it, and free-text model ids
 * remain valid (Enter/typing saves as-is). Refresh replaces the suggestions
 * without touching the current value.
 *
 * @param {Array<{id: string}>} models - OpenAI-format model objects
 */
function populateModels(models) {
    const current = getActiveBackendConfig(appState).model;
    const ids = new Set(models.map(m => m.id));

    _availableModels = models.map(m => m.id);
    // Prepend the configured model so it stays selectable even if the
    // provider no longer lists it.
    if (current && !ids.has(current)) {
        _availableModels.unshift(current);
    }
    _modelsByProvider.set(appState.config.backend, [..._availableModels]);

    const dropdown = document.getElementById('modelDropdown');
    if (dropdown && !dropdown.hidden) {
        renderModelDropdown(document.getElementById('modelSelect').value);
    }
    // A newly listed model can activate a model-specific capability profile.
    updateGenerationControls();
}

/**
 * Renders the model dropdown items, filtered by the given substring
 * (case-insensitive). An empty filter renders the full list.
 *
 * @param {string} [filter]
 */
function renderModelDropdown(filter = '') {
    const dropdown = document.getElementById('modelDropdown');
    if (!dropdown) return;

    const query = (filter || '').trim().toLowerCase();
    const current = getActiveBackendConfig(appState).model;
    const matches = query
        ? _availableModels.filter((id) => id.toLowerCase().includes(query))
        : _availableModels;

    dropdown.innerHTML = '';
    if (matches.length === 0) {
        const empty = document.createElement('div');
        empty.className = 'model-dropdown-empty';
        empty.textContent = _availableModels.length
            ? 'No matching models'
            : 'No models loaded — click Refresh';
        dropdown.appendChild(empty);
        return;
    }
    for (const id of matches) {
        const item = document.createElement('div');
        item.className = 'model-dropdown-item' + (id === current ? ' current' : '');
        item.setAttribute('role', 'option');
        item.textContent = id;
        // mousedown + preventDefault keeps focus in the input (no blur) so
        // the click reliably selects instead of being swallowed by blur.
        item.addEventListener('mousedown', (e) => {
            e.preventDefault();
            selectModel(id);
        });
        dropdown.appendChild(item);
    }
}

/** Opens the model dropdown showing the full (unfiltered) list. */
function openModelDropdown() {
    renderModelDropdown('');
    const dropdown = document.getElementById('modelDropdown');
    dropdown.hidden = false;
    document.getElementById('modelSelect').setAttribute('aria-expanded', 'true');
}

/** Closes the model dropdown. */
function closeModelDropdown() {
    const dropdown = document.getElementById('modelDropdown');
    if (!dropdown) return;
    dropdown.hidden = true;
    document.getElementById('modelSelect').setAttribute('aria-expanded', 'false');
}

/**
 * Picks a model from the dropdown: sets the input value, saves, closes.
 *
 * @param {string} id - The model id
 */
function selectModel(id) {
    document.getElementById('modelSelect').value = id;
    closeModelDropdown();
    updateGenerationControls(id);
    saveSettings();
}

// ============================================================================
// PROMPT MANAGEMENT
// ============================================================================

/**
 * Populates the dropdown for a single category from PromptManager state.
 * Selects the active prompt if one exists.
 *
 * @param {string} category - One of CATEGORIES
 */
function renderCategoryDropdown(category) {
    const select = document.getElementById(`promptSelect-${category}`);
    const prompts = appState.promptManager.getPrompts(category);
    const activePrompt = appState.promptManager.getActivePrompt(category);

    select.innerHTML = '<option value="">(None)</option>';

    const newOpt = document.createElement('option');
    newOpt.value = '__new__';
    newOpt.textContent = '+ New Prompt';
    select.appendChild(newOpt);

    prompts.forEach((prompt) => {
        const option = document.createElement('option');
        option.value = prompt.id;
        option.textContent = prompt.name;
        if (prompt.description) {
            option.title = prompt.description;
        }
        select.appendChild(option);
    });

    if (activePrompt) {
        select.value = activePrompt.id;
    }
}

/** Renders dropdowns for all categories. */
function renderAllDropdowns() {
    for (const category of CATEGORIES) {
        renderCategoryDropdown(category);
    }
}

/**
 * Handles selecting a prompt from a category's dropdown.
 * Auto-activates the selected prompt or deactivates if "(None)" is chosen.
 *
 * @param {string} category
 * @param {string} promptId - The prompt ID, '' for (None), '__new__' for new
 */
function handleCategoryPromptSelect(category, promptId) {
    const textarea = document.getElementById(`promptTextarea-${category}`);

    if (promptId === '__new__' || !promptId) {
        appState.promptManager.selectPrompt(category, null);
        textarea.value = '';
        if (category === 'amendment') {
            document.getElementById('commentInstructions').value = '';
        }
        addLog(promptId === '__new__'
            ? `${capitalize(category)}: ready for new prompt`
            : `${capitalize(category)} prompt deactivated`, 'info');
        return;
    }

    const prompt = appState.promptManager.selectPrompt(category, promptId);
    if (prompt) {
        textarea.value = prompt.template;
        if (category === 'amendment') {
            document.getElementById('commentInstructions').value = prompt.commentInstructions || '';
        }
        addLog(`Loaded ${category} prompt: ${prompt.name}`, 'info');
    }
}

/**
 * Saves the category textarea: updates the selected prompt in place, or opens
 * the save-prompt modal when no prompt is selected.
 *
 * @param {string} category
 */
function handleSavePrompt(category) {
    const select = document.getElementById(`promptSelect-${category}`);
    const selectedValue = select.value;
    const template = document.getElementById(`promptTextarea-${category}`).value.trim();

    if (!template) {
        addLog('Prompt template cannot be empty', 'warning');
        return;
    }

    if (selectedValue && selectedValue !== '__new__') {
        // Existing prompt selected -- update in-place
        const updates = { template };
        if (category === 'amendment') {
            updates.commentInstructions = document.getElementById('commentInstructions').value.trim();
        }
        appState.promptManager.updatePrompt(category, selectedValue, updates);
        addLog(`Prompt updated: ${appState.promptManager.getPrompt(category, selectedValue).name} (${category})`, 'success');
    } else {
        showSavePromptModal(category);
    }
}

/**
 * Deletes the currently selected prompt in a category.
 *
 * @param {string} category
 */
function handleDeletePromptConfirm(category) {
    const select = document.getElementById(`promptSelect-${category}`);
    const promptId = select.value;

    if (!promptId || promptId === '__new__') {
        addLog('No prompt selected to delete', 'warning');
        return;
    }

    const prompt = appState.promptManager.getPrompt(category, promptId);
    if (!prompt) return;

    appState.promptManager.deletePrompt(category, promptId);
    addLog(`Prompt deleted: ${prompt.name} (${category})`, 'success');

    renderCategoryDropdown(category);
    document.getElementById(`promptTextarea-${category}`).value = '';
    if (category === 'amendment') {
        document.getElementById('commentInstructions').value = '';
    }
}

/**
 * Clears the textarea for a category without deactivating.
 *
 * @param {string} category
 */
function handleResetPrompt(category) {
    document.getElementById(`promptTextarea-${category}`).value = '';
    addLog(`${capitalize(category)} prompt text cleared`, 'info');
}

/** Tracks which category the save-prompt modal is saving into. */
let _modalCategory = null;

/**
 * Opens the save prompt modal with category context.
 *
 * @param {string} category - The category being saved to
 */
function showSavePromptModal(category) {
    _modalCategory = category;
    _lastFocusedElement = document.activeElement;
    const modal = document.getElementById('savePromptModal');
    modal.classList.add('active');
    document.getElementById('savePromptCategory').textContent = `Saving to: ${capitalize(category)}`;
    document.getElementById('promptName').value = '';
    document.getElementById('promptDescription').value = '';
    document.getElementById('promptName').focus();
}

/** Hides the save prompt modal and restores focus to the invoking control. */
function hideSavePromptModal() {
    document.getElementById('savePromptModal').classList.remove('active');
    if (_lastFocusedElement && typeof _lastFocusedElement.focus === 'function') {
        _lastFocusedElement.focus();
    }
    _lastFocusedElement = null;
}

/** Handles the Save button in the save prompt modal. */
function handleSavePromptConfirm() {
    const category = _modalCategory;
    const name = document.getElementById('promptName').value.trim();
    const description = document.getElementById('promptDescription').value.trim();
    const template = document.getElementById(`promptTextarea-${category}`).value.trim();

    if (!name) {
        addLog('Please enter a prompt name', 'warning');
        return;
    }
    if (!template) {
        addLog('Prompt template cannot be empty', 'warning');
        return;
    }

    const promptData = { name, template, description };
    if (category === 'amendment') {
        promptData.commentInstructions = document.getElementById('commentInstructions').value.trim();
    }
    const prompt = appState.promptManager.addPrompt(category, promptData);
    addLog(`Prompt saved: ${name} (${category})`, 'success');

    renderCategoryDropdown(category);
    hideSavePromptModal();

    // Auto-select the saved prompt
    handleCategoryPromptSelect(category, prompt.id);
    document.getElementById(`promptSelect-${category}`).value = prompt.id;
}

/**
 * Capitalizes the first letter of a string.
 * @param {string} str
 * @returns {string}
 */
function capitalize(str) {
    return str.charAt(0).toUpperCase() + str.slice(1);
}

/**
 * Imported Skills (SKILL.md) management: import via paste or file, list,
 * remove. Imported packages surface as slash commands through the skill
 * registry (skills.js listSkills).
 */
function renderSkillImportList() {
    const list = document.getElementById('skillImportList');
    if (!list) return;
    list.textContent = '';

    const skills = loadImportedSkills();
    if (skills.length === 0) {
        const empty = document.createElement('p');
        empty.className = 'help-text';
        empty.textContent = 'No imported skills yet.';
        list.appendChild(empty);
        return;
    }
    for (const skill of skills) {
        const row = document.createElement('div');
        row.className = 'skill-import-row';

        const text = document.createElement('div');
        text.className = 'skill-import-text';
        const name = document.createElement('div');
        name.className = 'skill-import-name';
        name.textContent = `${skill.slash} · ${skill.category} · ${skill.scope}`;
        const desc = document.createElement('div');
        desc.className = 'skill-import-desc';
        desc.textContent = skill.description;
        text.appendChild(name);
        text.appendChild(desc);

        const removeBtn = document.createElement('button');
        removeBtn.type = 'button';
        removeBtn.className = 'btn btn-compact';
        removeBtn.textContent = 'Remove';
        removeBtn.addEventListener('click', () => {
            removeImportedSkill(skill.name);
            addLog(`Removed imported skill ${skill.slash}.`, 'info');
            renderSkillImportList();
        });

        row.appendChild(text);
        row.appendChild(removeBtn);
        list.appendChild(row);
    }
}

/**
 * Renders the read-only list of built-in slash commands (plus the reserved
 * /mcp skill) so they are discoverable from Settings, not only by typing /.
 * Static content — rendered once at init.
 */
function renderBuiltinSkillList() {
    const list = document.getElementById('builtinSkillList');
    if (!list) return;
    list.textContent = '';

    for (const skill of [...BUILTIN_SKILLS, RESERVED_MCP_SKILL]) {
        const row = document.createElement('div');
        row.className = 'skill-import-row';

        const text = document.createElement('div');
        text.className = 'skill-import-text';
        const name = document.createElement('div');
        name.className = 'skill-import-name';
        name.textContent = `${skill.slash} · ${skill.category}`;
        const desc = document.createElement('div');
        desc.className = 'skill-import-desc';
        desc.textContent = skill.description;
        text.appendChild(name);
        text.appendChild(desc);

        row.appendChild(text);
        list.appendChild(row);
    }
}

function initSkillImport() {
    const importText = document.getElementById('skillImportText');
    const importBtn = document.getElementById('skillImportBtn');
    const clearBtn = document.getElementById('skillImportClearBtn');
    const fileInput = document.getElementById('skillImportFile');
    if (!importText || !importBtn) return; // markup absent — nothing to wire

    renderBuiltinSkillList();

    fileInput.addEventListener('change', () => {
        const file = fileInput.files && fileInput.files[0];
        if (!file) return;
        file.text().then((content) => {
            importText.value = content;
        }).catch((err) => {
            addLog(`Could not read ${file.name}: ${err.message}`, 'error');
        });
    });

    importBtn.addEventListener('click', () => {
        const pkg = parseSkillPackage(importText.value, addLog);
        if (!pkg) {
            addLog('Skill import failed — see the warning above.', 'error');
            return;
        }
        const result = addImportedSkill(pkg);
        if (!result.ok) {
            addLog(`Skill import failed: ${result.error}`, 'error');
            return;
        }
        importText.value = '';
        fileInput.value = '';
        addLog(`Imported skill ${pkg.slash} — it now works as a slash command.`, 'success');
        renderSkillImportList();
    });

    clearBtn.addEventListener('click', () => {
        importText.value = '';
        fileInput.value = '';
    });

    renderSkillImportList();
}

/**
 * MCP Servers management: add/test/remove tool servers the /mcp command
 * can use. Persisted through the normal config pipeline (saveSettings →
 * normalizeConfig → localStorage).
 */
function renderMcpServerList() {
    const list = document.getElementById('mcpServerList');
    if (!list) return;
    list.textContent = '';

    const servers = (appState.config.mcpServers || []);
    if (servers.length === 0) {
        const empty = document.createElement('p');
        empty.className = 'help-text';
        empty.textContent = 'No MCP servers configured yet. /mcp will explain itself until you add one.';
        list.appendChild(empty);
        return;
    }
    servers.forEach((server, index) => {
        const row = document.createElement('div');
        row.className = 'skill-import-row';

        const text = document.createElement('div');
        text.className = 'skill-import-text';
        const name = document.createElement('div');
        name.className = 'skill-import-name';
        name.textContent = `${server.name || server.url}${server.enabled === false ? ' (disabled)' : ''}`;
        const desc = document.createElement('div');
        desc.className = 'skill-import-desc';
        desc.textContent = server.url;
        text.appendChild(name);
        text.appendChild(desc);

        const toolsBtn = document.createElement('button');
        toolsBtn.type = 'button';
        toolsBtn.className = 'btn btn-compact';
        toolsBtn.textContent = 'Test';
        toolsBtn.addEventListener('click', async () => {
            addLog(`Testing MCP server "${server.name || server.url}"...`, 'info');
            try {
                const client = await connectMcpServer({ url: server.url, token: server.token, log: addLog });
                const tools = await client.listTools();
                addLog(`MCP "${server.name || server.url}" OK: ${tools.length} tool(s) — ${tools.map((t) => t.name).join(', ') || '(none)'}.`, 'success');
            } catch (err) {
                addLog(`MCP "${server.name || server.url}" failed: ${err.message}`, 'error');
            }
        });

        const toggleBtn = document.createElement('button');
        toggleBtn.type = 'button';
        toggleBtn.className = 'btn btn-compact';
        toggleBtn.textContent = server.enabled === false ? 'Enable' : 'Disable';
        toggleBtn.addEventListener('click', () => {
            appState.config.mcpServers[index].enabled = server.enabled === false;
            saveSettings();
            renderMcpServerList();
        });

        const importBtn = document.createElement('button');
        importBtn.type = 'button';
        importBtn.className = 'btn btn-compact';
        importBtn.textContent = 'Import prompts';
        importBtn.addEventListener('click', async () => {
            addLog(`Importing prompts from MCP "${server.name || server.url}"...`, 'info');
            try {
                const client = await connectMcpServer({ url: server.url, token: server.token, log: addLog });
                const result = await importServerPrompts(server.name || 'mcp', client);
                let added = 0;
                for (const descriptor of result.imported) {
                    const addResult = addImportedSkill(descriptor);
                    if (addResult.ok) added += 1;
                    else addLog(`Skipped ${descriptor.slash}: ${addResult.error}`, 'warning');
                }
                for (const error of result.errors) addLog(`Skipped an MCP prompt: ${error}`, 'warning');
                addLog(`Imported ${added} prompt(s) as slash commands from "${server.name || server.url}".`, added ? 'success' : 'info');
                renderSkillImportList();
            } catch (err) {
                addLog(`Prompt import failed: ${err.message}`, 'error');
            }
        });

        const removeBtn = document.createElement('button');
        removeBtn.type = 'button';
        removeBtn.className = 'btn btn-compact';
        removeBtn.textContent = 'Remove';
        removeBtn.addEventListener('click', () => {
            appState.config.mcpServers.splice(index, 1);
            saveSettings();
            renderMcpServerList();
        });

        row.appendChild(text);
        row.appendChild(importBtn);
        row.appendChild(toolsBtn);
        row.appendChild(toggleBtn);
        row.appendChild(removeBtn);
        list.appendChild(row);
    });
}

function initMcpServers() {
    const addBtn = document.getElementById('mcpServerAddBtn');
    if (!addBtn) return; // markup absent — nothing to wire

    addBtn.addEventListener('click', () => {
        const nameInput = document.getElementById('mcpServerName');
        const urlInput = document.getElementById('mcpServerUrl');
        const tokenInput = document.getElementById('mcpServerToken');
        const url = (urlInput.value || '').trim();
        if (!url) {
            addLog('MCP server URL is required.', 'warning');
            return;
        }
        appState.config.mcpServers = appState.config.mcpServers || [];
        const addedName = (nameInput.value || '').trim();
        appState.config.mcpServers.push({
            name: addedName,
            url,
            token: (tokenInput.value || '').trim(),
            enabled: true,
        });
        nameInput.value = '';
        urlInput.value = '';
        tokenInput.value = '';
        saveSettings();
        addLog(`Added MCP server "${addedName || url}".`, 'success');
        renderMcpServerList();
    });

    renderMcpServerList();
}

/**
 * Per-turn step budget for MCP tool loops (default: TOOL_LOOP_LIMITS).
 */
function initMcpStepBudget() {
    const input = document.getElementById('mcpStepBudget');
    if (!input) return;
    input.value = String(appState.config.mcpStepBudget || TOOL_LOOP_LIMITS.MAX_STEPS_DEFAULT);
    input.addEventListener('change', () => {
        const value = Number(input.value);
        appState.config.mcpStepBudget = Number.isFinite(value) && value > 0
            ? Math.min(Math.round(value), 48)
            : TOOL_LOOP_LIMITS.MAX_STEPS_DEFAULT;
        input.value = String(appState.config.mcpStepBudget);
        saveSettings();
    });
}

// ============================================================================
// PANEL GEOMETRY (drag & resize)
// ============================================================================

/** Smallest usable settings panel while resizing. */
const PANEL_MIN_WIDTH = 320;
const PANEL_MIN_HEIGHT = 260;

/** Handle directions; corner handles combine two edge letters. */
const RESIZE_DIRECTIONS = ['n', 'ne', 'e', 'se', 's', 'sw', 'w', 'nw'];

/**
 * Makes the settings panel draggable by its header and resizable from any
 * edge or corner. The panel starts docked (flex right edge, full height);
 * the first drag or resize gesture floats it with explicit absolute
 * geometry inside the fixed overlay. Closing resets to the docked layout.
 */
function initPanelGeometry() {
    const overlay = document.getElementById('settingsOverlay');
    const panel = document.getElementById('settingsPanel');
    if (!overlay || !panel) return; // markup absent — nothing to wire

    const header = panel.querySelector('.settings-header');
    header.title = 'Drag to move — double-click to dock';
    header.classList.add('settings-drag-handle');

    for (const dir of RESIZE_DIRECTIONS) {
        const handle = document.createElement('div');
        handle.className = `settings-resize-handle settings-resize-${dir}`;
        handle.addEventListener('pointerdown', (e) => startPanelResize(e, overlay, panel, dir));
        panel.appendChild(handle);
    }

    header.addEventListener('pointerdown', (e) => {
        // The header chrome drags the panel; its controls still click.
        if (e.target.closest('button, input, select, textarea, a')) return;
        e.preventDefault();
        startPanelDrag(e, overlay, panel, header);
    });
    header.addEventListener('dblclick', (e) => {
        if (e.target.closest('button, input, select, textarea, a')) return;
        resetPanelGeometry();
    });
}

/** Switches the docked panel to floating mode, pinned at its current rect. */
function floatPanel(panel) {
    if (panel.classList.contains('floating')) return;
    const rect = panel.getBoundingClientRect();
    panel.classList.add('floating');
    panel.style.left = `${Math.round(rect.left)}px`;
    panel.style.top = `${Math.round(rect.top)}px`;
    panel.style.width = `${Math.round(rect.width)}px`;
    panel.style.height = `${Math.round(rect.height)}px`;
}

/** Restores the docked layout (on close and on header double-click). */
function resetPanelGeometry() {
    const panel = document.getElementById('settingsPanel');
    if (!panel) return;
    panel.classList.remove('floating');
    panel.style.left = '';
    panel.style.top = '';
    panel.style.width = '';
    panel.style.height = '';
}

/**
 * Drags the floating panel by the header, clamped to the overlay bounds.
 * Pointer capture keeps the gesture alive when the cursor leaves the header.
 */
function startPanelDrag(e, overlay, panel, header) {
    floatPanel(panel);
    const startX = e.clientX;
    const startY = e.clientY;
    const startLeft = parseFloat(panel.style.left);
    const startTop = parseFloat(panel.style.top);
    header.classList.add('dragging');
    header.setPointerCapture(e.pointerId);

    const onMove = (ev) => {
        const left = clamp(startLeft + ev.clientX - startX, 0, overlay.clientWidth - panel.offsetWidth);
        const top = clamp(startTop + ev.clientY - startY, 0, overlay.clientHeight - panel.offsetHeight);
        panel.style.left = `${left}px`;
        panel.style.top = `${top}px`;
    };
    const onUp = () => {
        header.classList.remove('dragging');
        header.removeEventListener('pointermove', onMove);
        header.removeEventListener('pointerup', onUp);
        header.removeEventListener('pointercancel', onUp);
    };
    header.addEventListener('pointermove', onMove);
    header.addEventListener('pointerup', onUp);
    header.addEventListener('pointercancel', onUp);
}

/**
 * Resizes the floating panel from the given edge/corner direction. Edges
 * being pulled from stay fixed: 'w'/'n' move the left/top edge (and origin),
 * 'e'/'s' grow the box, everything clamped to the minimum size and overlay.
 */
function startPanelResize(e, overlay, panel, dir) {
    e.preventDefault();
    e.stopPropagation();
    floatPanel(panel);
    const handle = e.currentTarget;
    const start = {
        x: e.clientX,
        y: e.clientY,
        left: parseFloat(panel.style.left),
        top: parseFloat(panel.style.top),
        width: panel.offsetWidth,
        height: panel.offsetHeight,
    };
    handle.setPointerCapture(e.pointerId);
    panel.classList.add('resizing');

    const onMove = (ev) => {
        const dx = ev.clientX - start.x;
        const dy = ev.clientY - start.y;
        let { left, top, width, height } = start;
        if (dir.includes('e')) width = clamp(start.width + dx, PANEL_MIN_WIDTH, overlay.clientWidth - start.left);
        if (dir.includes('s')) height = clamp(start.height + dy, PANEL_MIN_HEIGHT, overlay.clientHeight - start.top);
        if (dir.includes('w')) {
            width = clamp(start.width - dx, PANEL_MIN_WIDTH, start.left + start.width);
            left = start.left + start.width - width;
        }
        if (dir.includes('n')) {
            height = clamp(start.height - dy, PANEL_MIN_HEIGHT, start.top + start.height);
            top = start.top + start.height - height;
        }
        panel.style.left = `${left}px`;
        panel.style.top = `${top}px`;
        panel.style.width = `${width}px`;
        panel.style.height = `${height}px`;
    };
    const onUp = () => {
        panel.classList.remove('resizing');
        handle.removeEventListener('pointermove', onMove);
        handle.removeEventListener('pointerup', onUp);
        handle.removeEventListener('pointercancel', onUp);
    };
    handle.addEventListener('pointermove', onMove);
    handle.addEventListener('pointerup', onUp);
    handle.addEventListener('pointercancel', onUp);
}

/** Clamps value into [min, max]; tolerates a max below min (degenerate). */
function clamp(value, min, max) {
    return Math.min(Math.max(value, min), Math.max(min, max));
}
