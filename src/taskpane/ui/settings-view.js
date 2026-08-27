/**
 * Settings View Module
 *
 * The settings slide-over: LLM provider settings (auto-saved, debounced for
 * free-text fields, connection re-tested on change) and prompt management
 * (per-category list/edit/save/delete/clear plus the save-prompt modal).
 *
 * Element IDs match the pre-refactor markup so the logic stays a straight
 * port; localStorage keys (wordAI.config, wordAI.prompts.*) are unchanged.
 *
 * @module ui/settings-view
 */

import { testConnection as llmTestConnection } from '../../lib/llm-client.js';
import { CATEGORIES } from '../../lib/prompt-manager.js';
import { getProviderPreset } from '../../lib/providers.js';
import { appState, getActiveBackendConfig, debounce, persistSettings } from '../app-state.js';
import { addLog, setConnectionStatus } from './status-bar.js';

let _onConfigChanged = () => {};
let _lastFocusedElement = null;

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

    // Auto-save settings on every change. Free-text inputs (URL, API key) are
    // debounced so a burst of keystrokes does not fire a save + connection
    // probe per character. The header Save button is an explicit affordance
    // for the same saveSettings path.
    const debouncedSaveSettings = debounce(saveSettings, 400);
    const modelInput = document.getElementById('modelSelect');
    modelInput.addEventListener('input', saveSettings);
    // Model combobox: clicking/focusing shows the full list, typing filters.
    modelInput.addEventListener('input', () => renderModelDropdown(modelInput.value));
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
    document.getElementById('docRichnessSelect').addEventListener('change', saveSettings);
    document.getElementById('trackedChangesExtraction').addEventListener('change', saveSettings);
    document.getElementById('commentGranularity').addEventListener('change', saveSettings);
    document.getElementById('includeCommentsInSelectionCheckbox').addEventListener('change', saveSettings);

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
}

/** Closes the settings slide-over. */
export function closeSettings() {
    closeModelDropdown();
    document.getElementById('settingsOverlay').setAttribute('hidden', '');
}

/**
 * Activates a settings tab ("general" or "prompts") and shows its page.
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

    config.backend = backend;
    config.providers[backend].url = endpointUrl || config.providers[backend].url;
    config.providers[backend].apiKey = apiKey;
    // Every provider's model is editable: users can type a model id not
    // present in the refreshable list (e.g. a newly released one).
    config.providers[backend].model = selectedModel || config.providers[backend].model;
    config.trackChangesEnabled = document.getElementById('trackChangesCheckbox').checked;
    config.lineDiffEnabled = document.getElementById('lineDiffCheckbox').checked;
    config.docExtraction = {
        richness: document.getElementById('docRichnessSelect').value
    };
    config.trackedChangesExtraction = document.getElementById('trackedChangesExtraction').checked;
    config.commentGranularity = parseInt(document.getElementById('commentGranularity').value || '0', 10);
    config.includeCommentsInSelection = document.getElementById('includeCommentsInSelectionCheckbox').checked;

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

    const richnessSelect = document.getElementById('docRichnessSelect');
    if (richnessSelect && config.docExtraction) {
        richnessSelect.value = config.docExtraction.richness || 'structured';
    }

    document.getElementById('trackedChangesExtraction').checked = !!config.trackedChangesExtraction;
    document.getElementById('commentGranularity').value = String(config.commentGranularity || 0);
    document.getElementById('includeCommentsInSelectionCheckbox').checked = !!config.includeCommentsInSelection;

    updateProviderHints();
}

/**
 * Handles switching between providers: restores the selected provider's saved
 * settings, refreshes hints, and re-tests the connection.
 */
function handleBackendSwitch() {
    appState.config.backend = document.getElementById('backendSelect').value;
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
        endpointHint.textContent = isCustom
            ? 'Base URL of any OpenAI-compatible server (proxy path or full https:// URL)'
            : `Base URL for ${preset ? preset.label : config.backend} — leave the default proxy path unless you host the backend elsewhere`;
    }
    const keyHint = document.getElementById('apiKeyHint');
    if (keyHint && preset && preset.keyHint) {
        keyHint.textContent = `Get an API key at ${preset.keyHint} (leave blank for local backends)`;
    }
}

/**
 * Tests connection to the active LLM backend and populates the model
 * suggestion list. Updates the header status dot.
 */
export async function testConnectionUI() {
    const backendConfig = getActiveBackendConfig(appState);
    const preset = getProviderPreset(appState.config.backend);
    const backendLabel = preset ? preset.label : appState.config.backend;

    setConnectionStatus('connecting', `${backendLabel}: Connecting...`);

    // A custom provider with no URL yet cannot be probed.
    if (!backendConfig.url) {
        setConnectionStatus('error', `${backendLabel}: enter an endpoint URL`);
        return;
    }

    try {
        const result = await llmTestConnection(backendConfig);
        setConnectionStatus('connected', `${backendLabel}: Connected`);
        addLog(`Connected to ${backendLabel}! Found ${result.models.length} model(s).`, 'success');
        populateModels(result.models);
    } catch (error) {
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

    const dropdown = document.getElementById('modelDropdown');
    if (dropdown && !dropdown.hidden) {
        renderModelDropdown(document.getElementById('modelSelect').value);
    }
}

/** Model ids from the last successful refresh, plus the configured model. */
let _availableModels = [];

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
