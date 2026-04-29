/* global Word, Office */

// Import CSS for webpack to bundle
import './taskpane.css';
import { applyTokenMapStrategy, applySentenceDiffStrategy } from 'office-word-diff';
import { sendPrompt, testConnection as llmTestConnection, stripMarkdown } from '../lib/llm-client.js';
import { PromptManager, CATEGORIES } from '../lib/prompt-manager.js';
import { CommentQueue } from '../lib/comment-queue.js';
import { fireCommentRequest } from '../lib/comment-request.js';
import { extractAllComments, extractDocumentText, extractDocumentStructured, estimateTokenCount, extractTrackedChanges, extractCommentsOnRange } from '../lib/comment-extractor.js';
import { formatSelectionWithComments } from '../lib/selection-with-comments.js';
import { createSummaryDocument, buildSummaryHtml } from '../lib/document-generator.js';
import { parseDelimitedResponse, buildFallbackClassificationPrompt } from '../lib/response-parser.js';
import { parseDocument } from '../lib/document-parser.js';
import { chunkDocument } from '../lib/document-chunker.js';
import { extractContext } from '../lib/context-extractor.js';
import { processChunksParallel } from '../lib/orchestrator.js';
import { bookmarkChunkRanges, applyChunkResults, cleanupBookmarks } from '../lib/reassembler.js';
import { CATEGORY, SCOPE } from '../lib/panel-actions.js';

// Global configuration (defaults from env, overridable via UI/localStorage)
let config = {
    backend: 'ollama',
    trackChangesEnabled: true,
    lineDiffEnabled: false,
    docExtraction: {
        richness: 'structured'
    },
    trackedChangesExtraction: false,
    commentGranularity: 0,
    includeCommentsInSelection: false,
    backends: {
        ollama: {
            url: process.env.DEFAULT_OLLAMA_URL || '/ollama',
            apiKey: '',
            model: process.env.DEFAULT_MODEL || 'gpt-oss:20b'
        },
        vllm: {
            url: process.env.DEFAULT_VLLM_URL || '/vllm',
            apiKey: '',
            model: process.env.DEFAULT_VLLM_MODEL || 'qwen3.5-35b-a3b'
        }
    }
};

/**
 * Returns the config object for the currently selected backend.
 * @returns {{ url: string, apiKey: string, model: string }}
 */
function getActiveBackendConfig() {
    return config.backends[config.backend];
}

const promptManager = new PromptManager();
let currentTab = 'context';
const unsavedText = { context: '', amendment: '', comment: '', summary: '' };
let isProcessing = false;
let isProcessingDoc = false;
let processDocController = null; // AbortController for cancellation
let _inflightBtn = null; // Phase 05.1-04b: DOM ref of the panel button currently driving a doc-scope op (cancel-morph target)
let supportsComments = false;  // Set during initialize() via WordApi 1.4 check
const commentQueue = new CommentQueue(addLog);

// Token estimate cache -- avoids repeated Word API calls
let _tokenEstimateCache = { docCharCount: null, commentCount: null };
let _tokenEstimateDirty = true;  // Set true to trigger Word API re-read
let _tokenEstimateTimer = null;  // Debounce timer

// ============================================================================
// PANEL-BUTTON DISPATCHER (Phase 05.1 — decoupled from active-mode reads)
// ============================================================================
// These exports are defined at module scope (above any Office.js
// side-effects) so `require('../src/taskpane/taskpane.js')` succeeds in
// Jest node + jsdom test environments. STYLE.md: "Dispatch Over If/Else
// Chains" + "Enums for Fixed Values" + "JSDoc on Public Methods".

/**
 * Returns the panel-button enable state for a category.
 * Single source of truth for the AC-05/06/07 enable rules. Read by
 * updatePanelButtons; not used by the dispatcher itself.
 *
 * @param {string} category — one of CATEGORY.AMENDMENT/COMMENT/SUMMARY
 * @param {object} [deps]   — optional override hooks for tests
 * @returns {{enabled: boolean, withCommentEnabled: boolean}}
 */
export function getPanelButtonState(category, deps = {}) {
    const pm = deps.promptManager || promptManager;
    const doc = deps.document || (typeof document !== 'undefined' ? document : null);
    const hasActive = !!pm.getActivePrompt(category);
    const ci = doc ? doc.getElementById('commentInstructions') : null;
    const hasCommentText = !!(ci && ci.value && ci.value.trim().length > 0);
    return {
        enabled: hasActive,
        withCommentEnabled: hasActive && hasCommentText,
    };
}

/**
 * Toggles `disabled` on this category's panel buttons based on prompt
 * activation + commentInstructions text. No-op if the document is
 * unavailable (allows pure-function calls under node-env tests).
 *
 * @param {string} category — one of CATEGORY.AMENDMENT/COMMENT/SUMMARY
 * @param {object} [deps]   — optional override hooks for tests
 */
export function updatePanelButtons(category, deps = {}) {
    const doc = deps.document || (typeof document !== 'undefined' ? document : null);
    if (!doc) return;
    const { enabled, withCommentEnabled } = getPanelButtonState(category, deps);
    const btns = doc.querySelectorAll(`[data-panel="${category}"][data-action]`);
    btns.forEach((btn) => {
        const action = btn.getAttribute('data-action');
        const isCommentVariant = action.includes('amend-comment');
        if (category === CATEGORY.AMENDMENT && isCommentVariant) {
            btn.disabled = !withCommentEnabled;
        } else {
            btn.disabled = !enabled;
        }
    });
}

/**
 * Morphs the initiating panel button to 'Cancel' and disables peer buttons during a doc-scope op.
 * Stores the inflight (category, withComment) on btn.dataset so the addEventListener
 * cancel-mode branch can re-derive the args needed to fire the abort path. Also stashes
 * btn.dataset.originalLabel for restore in clearInflightButton().
 *
 * @param {HTMLButtonElement|null} btn — the initiating panel button (may be null in test/edge paths)
 * @param {{category: string, withComment: boolean}} [opts]
 */
function setInflightButton(btn, { category, withComment } = {}) {
    if (!btn) return;
    _inflightBtn = btn;
    btn.dataset.originalLabel = btn.textContent;
    btn.dataset.inflightCategory = category || '';
    btn.dataset.inflightWithComment = withComment ? 'true' : 'false';
    btn.textContent = 'Cancel';
    btn.classList.add('cancel-mode'); // matches .btn.cancel-mode (Plan 02)
    // Disable all OTHER panel buttons; the inflight button itself stays enabled (it IS the cancel button).
    if (typeof document !== 'undefined') {
        document.querySelectorAll('[data-panel][data-action]').forEach((other) => {
            if (other !== btn) other.disabled = true;
        });
    }
    btn.disabled = false;
}

/**
 * Restores the inflight button's label/state after the doc-scope op finishes or aborts.
 * Restores originalLabel, removes .cancel-mode, clears the inflight dataset entries,
 * and clears the module-scope _inflightBtn ref. Does NOT re-enable peer buttons —
 * the caller's updatePanelButtons() triplet in the finally block handles that, so the
 * re-enable state reflects current prompt activation rather than the inflight lockout.
 */
function clearInflightButton() {
    if (_inflightBtn) {
        if (_inflightBtn.dataset.originalLabel) {
            _inflightBtn.textContent = _inflightBtn.dataset.originalLabel;
            delete _inflightBtn.dataset.originalLabel;
        }
        delete _inflightBtn.dataset.inflightCategory;
        delete _inflightBtn.dataset.inflightWithComment;
        _inflightBtn.classList.remove('cancel-mode');
    }
    _inflightBtn = null;
}

/**
 * Frozen dispatch table — keys are tuple strings shaped as
 *   `${category}:${scope}:${withComment ? 'withComment' : 'plain'}`.
 * Exactly 7 entries (one per RESEARCH.md Pattern 3 row). Adding an 8th
 * tuple silently is caught by tests/dispatcher.spec.js
 * dispatch-table-completeness. STYLE.md: "Dispatch Over If/Else Chains"
 * + "Enums for Fixed Values".
 *
 * Each value is a `(deps) => Promise<void>` thunk that calls the right
 * handler with the args from the original tuple. The dispatcher
 * (createDispatcher below) looks up the route via key and invokes it.
 */
export const ROUTES = Object.freeze({
    'amendment:selection:plain':       (deps) => deps.handleReviewSelection({ category: CATEGORY.AMENDMENT, withComment: false }),
    'amendment:selection:withComment': (deps) => deps.handleReviewSelection({ category: CATEGORY.AMENDMENT, withComment: true }),
    'amendment:document:plain':        (deps) => deps.handleProcessDocument({ category: CATEGORY.AMENDMENT, withComment: false }),
    'amendment:document:withComment':  (deps) => deps.handleProcessDocument({ category: CATEGORY.AMENDMENT, withComment: true }),
    'comment:selection:plain':         (deps) => deps.handleReviewSelection({ category: CATEGORY.COMMENT,   withComment: false }),
    'comment:document:plain':          (deps) => deps.handleProcessDocument({ category: CATEGORY.COMMENT,   withComment: false }),
    'summary:document:plain':          (deps) => deps.handleSummaryGeneration(),
});

/**
 * Builds the runAction dispatcher with injected handler deps. Returns
 * the runAction function. Factory shape exists so
 * tests/dispatcher.spec.js can inject spies without going through
 * Office.js side-effects.
 *
 * @param {object} deps — handler refs + isProcessingDocRef + addLog
 * @returns {function({category: string, scope: string, withComment: boolean}): Promise<void>}
 */
export function createDispatcher(deps) {
    /**
     * Routes a panel-button click to the correct handler by explicit
     * args — never reads any active-mode global. Looks up the route in the frozen
     * ROUTES table.
     *
     * @param {{category: string, scope: string, withComment: boolean}} args
     */
    return async function runAction({ category, scope, withComment }) {
        if (deps.isProcessingDocRef && deps.isProcessingDocRef()) {
            if (deps.addLog) {
                deps.addLog('Already processing the document. Cancel first.', 'warning');
            }
            return;
        }
        const key = `${category}:${scope}:${withComment ? 'withComment' : 'plain'}`;
        const route = ROUTES[key];
        if (!route) {
            // Defensive: log and return. Should be impossible if HTML
            // data-action strings stay in sync with the ACTION enum
            // (Plan 02 grep gate).
            // eslint-disable-next-line no-console
            console.warn('runAction: no route for', key);
            return;
        }
        return route(deps);
    };
}

// ============================================================================
// END PANEL-BUTTON DISPATCHER
// ============================================================================

if (typeof Office !== 'undefined') {
    Office.onReady((info) => {
        if (info.host === Office.HostType.Word) {
            initialize();
        }
    });
}

function initialize() {
    // Load saved settings
    loadSettings();

    // Load prompt state from localStorage
    promptManager.loadState();

    // Setup event listeners -- general
    // Panel-scoped action buttons (Phase 05.1) -- dispatcher routes via
    // explicit (category, scope, withComment) args — never via active-mode.
    const runAction = createDispatcher({
        promptManager,
        handleReviewSelection,
        handleProcessDocument,
        handleSummaryGeneration,
        fireCommentRequest,
        isProcessingDocRef: () => isProcessingDoc,
        addLog,
    });
    document.querySelectorAll('[data-panel][data-action]').forEach((btn) => {
        const category = btn.getAttribute('data-panel'); // 'amendment' | 'comment' | 'summary'
        const action = btn.getAttribute('data-action');
        const scope = action.endsWith('-document') ? SCOPE.DOCUMENT : SCOPE.SELECTION;
        const withComment = action.includes('amend-comment');
        btn.addEventListener('click', () => {
            // LOCKED cancel-morph branch (Plan 04b, revision Issue 2):
            // If this button is currently morphed to Cancel, bypass the dispatcher
            // pre-flight (which would reject because isProcessingDoc === true) and
            // call handleProcessDocument directly with the original inflight args.
            // The handler's existing cancel branch
            //   `if (isProcessingDoc && processDocController) { processDocController.abort(); return; }`
            // fires. Single source of cancel-routing logic.
            if (btn.classList.contains('cancel-mode')) {
                const inflightCategory = btn.dataset.inflightCategory || category;
                const inflightWithComment = btn.dataset.inflightWithComment === 'true';
                return handleProcessDocument({ category: inflightCategory, withComment: inflightWithComment });
            }
            return runAction({ category, scope, withComment });
        });
    });

    document.getElementById("clearLogsBtn").onclick = clearLogs;
    document.getElementById("settingsToggle").onclick = toggleSettings;
    document.getElementById("runVerificationBtn").onclick = runVerification;
    document.getElementById("backendSelect").onchange = handleBackendSwitch;

    // Auto-save settings on every change (no Save button needed)
    document.getElementById("backendSelect").addEventListener('change', saveSettings);
    document.getElementById("modelSelect").addEventListener('change', saveSettings);
    document.getElementById("endpointUrl").addEventListener('input', saveSettings);
    document.getElementById("apiKey").addEventListener('input', saveSettings);
    document.getElementById("trackChangesCheckbox").addEventListener('change', saveSettings);
    document.getElementById("lineDiffCheckbox").addEventListener('change', saveSettings);
    document.getElementById("docRichnessSelect").addEventListener('change', saveSettings);
    document.getElementById("trackedChangesExtraction").addEventListener('change', saveSettings);
    document.getElementById("commentGranularity").addEventListener('change', saveSettings);
    document.getElementById("includeCommentsInSelectionCheckbox").addEventListener('change', saveSettings);

    // Tab bar -- click and keyboard navigation
    for (const category of CATEGORIES) {
        const tabBtn = document.getElementById(`tab-${category}`);
        tabBtn.addEventListener('click', () => switchTab(category));
        tabBtn.addEventListener('keydown', handleTabKeydown);
    }

    // Per-category prompt controls
    for (const category of CATEGORIES) {
        document.getElementById(`promptSelect-${category}`).onchange = (e) => {
            handleCategoryPromptSelect(category, e.target.value);
        };
        document.getElementById(`savePromptBtn-${category}`).onclick = () => {
            const select = document.getElementById(`promptSelect-${category}`);
            const selectedValue = select.value;
            const textarea = document.getElementById(`promptTextarea-${category}`);
            const template = textarea.value.trim();

            if (!template) {
                addLog('Prompt template cannot be empty', 'warning');
                return;
            }

            if (selectedValue && selectedValue !== '__new__') {
                // Existing prompt selected -- update in-place
                const updates = { template };
                // Capture commentInstructions for amendment prompts
                if (category === 'amendment') {
                    const commentField = document.getElementById('commentInstructions');
                    updates.commentInstructions = commentField ? commentField.value.trim() : '';
                }
                promptManager.updatePrompt(category, selectedValue, updates);
                unsavedText[category] = template;
                addLog(`Prompt updated: ${promptManager.getPrompt(category, selectedValue).name} (${category})`, 'success');
            } else {
                // No prompt or "+ New Prompt" selected -- show create modal
                showSavePromptModal(category);
            }
        };
        document.getElementById(`deletePromptBtn-${category}`).onclick = () => {
            handleDeletePromptConfirm(category);
        };
        document.getElementById(`resetPromptBtn-${category}`).onclick = () => {
            handleResetPrompt(category);
        };
    }

    // Modal buttons
    document.getElementById("savePromptConfirmBtn").onclick = handleSavePromptConfirm;
    document.getElementById("savePromptCancelBtn").onclick = hideSavePromptModal;

    // Comment instructions field (amendment tab) -- enable/disable amendment
    // panel buttons in lockstep with textarea content (AC-06/07).
    document.getElementById('commentInstructions').addEventListener('input', () => {
        updatePanelButtons(CATEGORY.AMENDMENT);
    });

    // Initial UI state
    updateUIFromConfig();

    // Render prompt UI from PromptManager state
    renderAllDropdowns();
    updateDotIndicators();
    // Phase 05.1: update all panel buttons after initial render
    updatePanelButtons(CATEGORY.AMENDMENT);
    updatePanelButtons(CATEGORY.COMMENT);
    updatePanelButtons(CATEGORY.SUMMARY);
    updateTokenEstimate();

    // Detect and log supported Word API version (diagnostics only)
    const apiVersions = ['1.8', '1.7', '1.6', '1.5', '1.4', '1.3', '1.2', '1.1'];
    let detectedVersion = 'unknown';
    try {
        if (typeof Office !== 'undefined' && Office.context && Office.context.requirements) {
            for (const ver of apiVersions) {
                if (Office.context.requirements.isSetSupported('WordApi', ver)) {
                    detectedVersion = ver;
                    break;
                }
            }
        }
    } catch (e) { /* detection failed */ }
    addLog(`Word API version: ${detectedVersion}`, 'info');

    // Detect WordApi 1.4 support for comment features
    if (typeof Office !== 'undefined' && Office.context && Office.context.requirements) {
        supportsComments = Office.context.requirements.isSetSupported('WordApi', '1.4');
    }

    if (!supportsComments) {
        // Hide comment-related UI elements (graceful degradation)
        const commentTab = document.getElementById('tab-comment');
        const commentPanel = document.getElementById('panel-comment');
        const commentStatusBar = document.getElementById('commentStatusBar');
        if (commentTab) commentTab.style.display = 'none';
        if (commentPanel) commentPanel.style.display = 'none';
        if (commentStatusBar) commentStatusBar.style.display = 'none';
        addLog('Comment features unavailable (requires Word API 1.4)', 'info');
    }

    // Restore unsaved text from active prompts on load
    for (const category of CATEGORIES) {
        const activePrompt = promptManager.getActivePrompt(category);
        if (activePrompt) {
            unsavedText[category] = activePrompt.template;
            document.getElementById(`promptTextarea-${category}`).value = activePrompt.template;
            // Restore commentInstructions for amendment prompts
            if (category === 'amendment' && activePrompt.commentInstructions) {
                const commentField = document.getElementById('commentInstructions');
                if (commentField) {
                    commentField.value = activePrompt.commentInstructions;
                }
            }
        }
    }

    // Auto-test connection and load models
    testConnectionUI();

    addLog("Contract Review Add-in initialized.", "info");
}

// ============================================================================
// SETTINGS & UI
// ============================================================================

function loadSettings() {
    try {
        const saved = localStorage.getItem('wordAI.config');
        if (saved) {
            const parsed = JSON.parse(saved);

            if (parsed.ollamaUrl && !parsed.backends) {
                // Old flat format detected -- migrate to nested backends structure
                config.backend = 'ollama';
                config.backends.ollama.url = parsed.ollamaUrl;
                config.backends.ollama.apiKey = parsed.apiKey || '';
                config.backends.ollama.model = parsed.selectedModel || config.backends.ollama.model;
                if (typeof parsed.trackChangesEnabled === 'boolean') {
                    config.trackChangesEnabled = parsed.trackChangesEnabled;
                }
                if (typeof parsed.lineDiffEnabled === 'boolean') {
                    config.lineDiffEnabled = parsed.lineDiffEnabled;
                }
                // Save migrated config immediately so migration only runs once
                localStorage.setItem('wordAI.config', JSON.stringify(config));
            } else {
                // New nested format -- merge normally
                config = { ...config, ...parsed };
            }

            // Ensure docExtraction defaults exist (for configs saved before this feature)
            if (!config.docExtraction) {
                config.docExtraction = { richness: 'structured' };
            }
            // Clean up legacy maxLength from older configs
            if (config.docExtraction.maxLength !== undefined) {
                delete config.docExtraction.maxLength;
            }

            // Ensure trackedChangesExtraction default (for configs saved before this feature)
            if (config.trackedChangesExtraction === undefined) {
                config.trackedChangesExtraction = false;
            }

            // Ensure commentGranularity default (for configs saved before this feature)
            if (config.commentGranularity === undefined) {
                config.commentGranularity = 0;
            }

            // Ensure includeCommentsInSelection default (Phase 05.2 — for configs saved before this feature)
            if (config.includeCommentsInSelection === undefined) {
                config.includeCommentsInSelection = false;
            }
        }
    } catch (e) {
        console.error("Failed to load settings:", e);
    }
}

function saveSettings() {
    const backend = document.getElementById("backendSelect").value;
    const endpointUrl = document.getElementById("endpointUrl").value.trim();
    const apiKey = document.getElementById("apiKey").value.trim();
    const trackChanges = document.getElementById("trackChangesCheckbox").checked;
    const lineDiff = document.getElementById("lineDiffCheckbox").checked;
    const selectedModel = document.getElementById("modelSelect").value;

    config.backend = backend;
    config.backends[backend].url = endpointUrl || config.backends[backend].url;
    config.backends[backend].apiKey = apiKey;
    // Only update model for Ollama -- vLLM model is read-only
    if (backend === 'ollama') {
        config.backends[backend].model = selectedModel || config.backends[backend].model;
    }
    config.trackChangesEnabled = trackChanges;
    config.lineDiffEnabled = lineDiff;
    config.docExtraction = {
        richness: document.getElementById('docRichnessSelect').value
    };
    config.trackedChangesExtraction = document.getElementById('trackedChangesExtraction').checked;
    config.commentGranularity = parseInt(document.getElementById('commentGranularity').value || '0', 10);
    config.includeCommentsInSelection = document.getElementById('includeCommentsInSelectionCheckbox').checked;

    try {
        localStorage.setItem('wordAI.config', JSON.stringify(config));
        addLog("Settings saved.", "success");
        invalidateTokenEstimateCache();
        updateTokenEstimate();

        // Re-test connection with new settings
        testConnectionUI();
    } catch (e) {
        addLog(`Failed to save settings: ${e.message}`, "error");
    }
}

function updateUIFromConfig() {
    const backendConfig = getActiveBackendConfig();
    const modelSelect = document.getElementById("modelSelect");

    document.getElementById("backendSelect").value = config.backend;
    document.getElementById("endpointUrl").value = backendConfig.url;
    document.getElementById("apiKey").value = backendConfig.apiKey;
    document.getElementById("trackChangesCheckbox").checked = config.trackChangesEnabled;
    document.getElementById("lineDiffCheckbox").checked = config.lineDiffEnabled;

    if (config.backend === 'vllm') {
        // vLLM: show configured model as read-only (disabled dropdown)
        modelSelect.innerHTML = '';
        const option = document.createElement('option');
        option.value = backendConfig.model;
        option.textContent = backendConfig.model;
        modelSelect.appendChild(option);
        modelSelect.disabled = true;
    } else {
        // Ollama: enable dropdown (models populated by testConnectionUI)
        modelSelect.disabled = false;
    }

    const richnessSelect = document.getElementById('docRichnessSelect');
    if (richnessSelect && config.docExtraction) {
        richnessSelect.value = config.docExtraction.richness || 'structured';
    }

    const trackedChangesCheckbox = document.getElementById('trackedChangesExtraction');
    if (trackedChangesCheckbox) {
        trackedChangesCheckbox.checked = !!config.trackedChangesExtraction;
    }

    const granularitySelect = document.getElementById('commentGranularity');
    if (granularitySelect) {
        granularitySelect.value = String(config.commentGranularity || 0);
    }

    const includeCommentsCheckbox = document.getElementById('includeCommentsInSelectionCheckbox');
    if (includeCommentsCheckbox) {
        includeCommentsCheckbox.checked = !!config.includeCommentsInSelection;
    }
}

/**
 * Handles switching between backends in the UI.
 * Restores the selected backend's saved settings and triggers a connection test.
 */
function handleBackendSwitch() {
    config.backend = document.getElementById('backendSelect').value;
    updateUIFromConfig();
    testConnectionUI();
}

function toggleSettings() {
    const content = document.getElementById("settingsContent");
    const header = document.getElementById("settingsToggle");
    content.classList.toggle("active");
    header.classList.toggle("active");
}

// ============================================================================
// PROMPT MANAGEMENT (PromptManager Integration)
// ============================================================================

/**
 * Populates the dropdown for a single category from PromptManager state.
 * Selects the active prompt if one exists.
 *
 * @param {string} category - One of 'context', 'amendment', 'comment'
 */
function renderCategoryDropdown(category) {
    const select = document.getElementById(`promptSelect-${category}`);
    const prompts = promptManager.getPrompts(category);
    const activePrompt = promptManager.getActivePrompt(category);

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

    // Select the active prompt in the dropdown
    if (activePrompt) {
        select.value = activePrompt.id;
    }
}

/**
 * Renders dropdowns for all three categories.
 */
function renderAllDropdowns() {
    for (const category of CATEGORIES) {
        renderCategoryDropdown(category);
    }
}

/**
 * Handles selecting a prompt from a category's dropdown.
 * Auto-activates the selected prompt or deactivates if "(None)" is chosen.
 *
 * @param {string} category - One of 'context', 'amendment', 'comment'
 * @param {string} promptId - The prompt ID, or empty string for "(None)"
 */
function handleCategoryPromptSelect(category, promptId) {
    const textarea = document.getElementById(`promptTextarea-${category}`);

    if (promptId === '__new__') {
        promptManager.selectPrompt(category, null);
        textarea.value = '';
        unsavedText[category] = '';
        // Clear commentInstructions when starting a new amendment prompt
        if (category === 'amendment') {
            const commentField = document.getElementById('commentInstructions');
            if (commentField) commentField.value = '';
        }
        addLog(`${capitalize(category)}: ready for new prompt`, "info");
        updateDotIndicators();
        updatePanelButtons(category);
        updateTokenEstimate();
        return;
    }

    if (!promptId) {
        // "(None)" selected -- deactivate category
        promptManager.selectPrompt(category, null);
        textarea.value = '';
        unsavedText[category] = '';
        // Clear commentInstructions when deactivating amendment prompt
        if (category === 'amendment') {
            const commentField = document.getElementById('commentInstructions');
            if (commentField) commentField.value = '';
        }
        addLog(`${capitalize(category)} prompt deactivated`, "info");
    } else {
        // Select and auto-activate prompt
        const prompt = promptManager.selectPrompt(category, promptId);
        if (prompt) {
            textarea.value = prompt.template;
            unsavedText[category] = prompt.template;
            // Restore commentInstructions for amendment prompts
            if (category === 'amendment') {
                const commentField = document.getElementById('commentInstructions');
                if (commentField) {
                    commentField.value = prompt.commentInstructions || '';
                }
            }
            addLog(`Loaded ${category} prompt: ${prompt.name}`, "info");
        }
    }

    updateDotIndicators();
    updatePanelButtons(category);
    updateTokenEstimate();
}

/**
 * Switches to a different tab, preserving unsaved textarea edits.
 *
 * @param {string} category - The category tab to switch to
 */
function switchTab(category) {
    if (category === currentTab) return;

    // Save current textarea content before switching
    const currentTextarea = document.getElementById(`promptTextarea-${currentTab}`);
    unsavedText[currentTab] = currentTextarea.value;

    // Update tab bar ARIA and styles
    for (const cat of CATEGORIES) {
        const tabBtn = document.getElementById(`tab-${cat}`);
        const panel = document.getElementById(`panel-${cat}`);
        const isTarget = (cat === category);

        tabBtn.setAttribute('aria-selected', isTarget ? 'true' : 'false');
        tabBtn.classList.toggle('active', isTarget);
        tabBtn.tabIndex = isTarget ? 0 : -1;

        if (isTarget) {
            panel.removeAttribute('hidden');
        } else {
            panel.setAttribute('hidden', '');
        }
    }

    currentTab = category;

    // Restore textarea content for new tab
    const newTextarea = document.getElementById(`promptTextarea-${category}`);
    newTextarea.value = unsavedText[category];

    // Phase 05.1: defensive update — buttons in non-visible panels are
    // hidden but cheap to update.
    updatePanelButtons(category);

    // Invalidate token estimate cache on tab switch -- document may have changed
    invalidateTokenEstimateCache();
    updateTokenEstimate();
}

/**
 * Handles arrow key navigation within the tab bar per WAI-ARIA pattern.
 *
 * @param {KeyboardEvent} e
 */
function handleTabKeydown(e) {
    const currentIndex = CATEGORIES.indexOf(currentTab);
    let newIndex = currentIndex;

    switch (e.key) {
        case 'ArrowRight':
            newIndex = (currentIndex + 1) % CATEGORIES.length;
            break;
        case 'ArrowLeft':
            newIndex = (currentIndex - 1 + CATEGORIES.length) % CATEGORIES.length;
            break;
        case 'Home':
            newIndex = 0;
            break;
        case 'End':
            newIndex = CATEGORIES.length - 1;
            break;
        default:
            return; // Don't prevent default for other keys
    }

    e.preventDefault();
    const newCategory = CATEGORIES[newIndex];
    document.getElementById(`tab-${newCategory}`).focus();
    switchTab(newCategory);
}

/**
 * Updates the dot indicators on each tab to reflect activation state.
 * Green dot = active prompt, red dot = no active prompt.
 */
function updateDotIndicators() {
    for (const category of CATEGORIES) {
        const dot = document.getElementById(`dot-${category}`);
        const isActive = promptManager.getActivePrompt(category) !== null;
        dot.classList.toggle('active', isActive);
    }
}

/**
 * Invalidates the token estimate cache, causing the next
 * updateTokenEstimate() call to re-read from the Word API.
 */
function invalidateTokenEstimateCache() {
    _tokenEstimateDirty = true;
}

/**
 * Reads document size metrics from Word API for token estimation.
 * Cached: only calls Word API when _tokenEstimateDirty is true.
 * Returns cached values on subsequent calls until invalidated.
 *
 * @param {object} options - Which metrics are needed
 * @param {boolean} options.needDocText - Whether to read body.text length
 * @param {boolean} options.needComments - Whether to count comments
 * @returns {Promise<{docCharCount: number|null, commentCount: number|null}>}
 */
async function getDocumentMetrics({ needDocText, needComments }) {
    const docCached = !needDocText || _tokenEstimateCache.docCharCount !== null;
    const commentsCached = !needComments || _tokenEstimateCache.commentCount !== null;
    if (!_tokenEstimateDirty && docCached && commentsCached) {
        return _tokenEstimateCache;
    }

    try {
        await Word.run(async (context) => {
            const body = context.document.body;

            // Read body text length if needed
            if (needDocText) {
                body.load('text');
            }

            // Read comment count if needed
            let commentCollection = null;
            if (needComments && supportsComments) {
                commentCollection = body.getComments();
                commentCollection.load('items');
            }

            await context.sync();

            if (needDocText) {
                _tokenEstimateCache.docCharCount = (body.text || '').length;
            }
            if (needComments && supportsComments && commentCollection) {
                _tokenEstimateCache.commentCount = commentCollection.items.length;
            }
        });
        _tokenEstimateDirty = false;
    } catch (e) {
        // Word API unavailable (e.g., test environment) -- leave cache as null
        console.warn('Token estimate: Word API unavailable, using prompt-only estimate', e);
    }

    return _tokenEstimateCache;
}

/**
 * Updates the token estimation display with current prompt and data sizes.
 * Reads actual document size from Word API (cached + debounced) to show
 * realistic token estimates including document text and comments.
 *
 * Shows estimated total tokens across: active context prompt + active
 * category prompt (amendment/comment/summary) + document text + comments.
 *
 * Uses estimateTokenCount (Math.ceil(text.length / 4)) heuristic.
 * Informational only -- helps users gauge LLM context window fit.
 *
 * Async: callers fire-and-forget. DOM is updated when data is ready.
 */
async function updateTokenEstimate() {
    // Debounce: cancel pending call, schedule new one after 300ms
    if (_tokenEstimateTimer) {
        clearTimeout(_tokenEstimateTimer);
    }

    await new Promise((resolve) => {
        _tokenEstimateTimer = setTimeout(resolve, 300);
    });
    _tokenEstimateTimer = null;

    const container = document.getElementById('tokenEstimate');
    const valueEl = document.getElementById('tokenEstimateValue');
    const breakdownEl = document.getElementById('tokenEstimateBreakdown');
    if (!container || !valueEl) return;

    // Phase 05.1 Plan 04a: gate on currentTab (visible panel) rather than
    // the prompt-manager's active-mode global, so the estimate matches what
    // the user is actually looking at. If no category has any active prompt
    // at all, hide the estimate entirely.
    const anyActive = CATEGORIES.some((c) => !!promptManager.getActivePrompt(c));
    if (!anyActive) {
        container.style.display = 'none';
        return;
    }

    let totalTokens = 0;
    const parts = [];
    let needDocText = false;
    let needComments = false;
    let hasTrackedChanges = false;

    // Context prompt tokens (always included if active)
    const contextPrompt = promptManager.getActivePrompt('context');
    if (contextPrompt && contextPrompt.template) {
        const ctxTokens = estimateTokenCount(contextPrompt.template);
        totalTokens += ctxTokens;
        parts.push(`ctx:~${ctxTokens.toLocaleString()}`);
    }

    if (currentTab === CATEGORY.SUMMARY) {
        // Summary panel visible: summary prompt + actual document data estimates
        const summaryPrompt = promptManager.getActivePrompt('summary');
        if (summaryPrompt && summaryPrompt.template) {
            const summTokens = estimateTokenCount(summaryPrompt.template);
            totalTokens += summTokens;
            parts.push(`prompt:~${summTokens.toLocaleString()}`);

            if (summaryPrompt.template.includes('{whole document}')) {
                needDocText = true;
            }
            if (summaryPrompt.template.includes('{comments}')) {
                needComments = true;
            }
            if (config.trackedChangesExtraction && summaryPrompt.template.includes('{tracked changes}')) {
                hasTrackedChanges = true;
            }
        }
    } else {
        // Amendment/comment mode: category prompt
        const categories = ['amendment', 'comment'];
        for (const cat of categories) {
            const prompt = promptManager.getActivePrompt(cat);
            if (prompt && prompt.template) {
                const catTokens = estimateTokenCount(prompt.template);
                totalTokens += catTokens;
                parts.push(`${cat.substring(0, 5)}:~${catTokens.toLocaleString()}`);
            }
        }
        // Selection text is variable and unknown until user selects -- show note
        parts.push('+selection');
    }

    // Fetch real document metrics from Word API (cached + debounced)
    if (needDocText || needComments) {
        const metrics = await getDocumentMetrics({ needDocText, needComments });

        if (needDocText && metrics.docCharCount !== null) {
            const docTokens = Math.ceil(metrics.docCharCount / 4);
            totalTokens += docTokens;
            parts.push(`doc:~${docTokens.toLocaleString()}`);
        } else if (needDocText) {
            // Word API failed -- show note instead of number
            parts.push('+doc text');
        }

        if (needComments && metrics.commentCount !== null) {
            // Estimate ~50 tokens per comment (author + text + associated text)
            const commentTokens = metrics.commentCount * 50;
            totalTokens += commentTokens;
            parts.push(`comments:~${commentTokens.toLocaleString()}`);
        } else if (needComments) {
            parts.push('+comments');
        }
    }

    // Tracked changes: can't cheaply estimate OOXML parsing cost, show note
    if (hasTrackedChanges) {
        parts.push('+tracked changes');
    }

    container.style.display = 'flex';
    valueEl.textContent = `~${totalTokens.toLocaleString()}`;
    breakdownEl.textContent = `(${parts.join(' | ')})`;

    // Color coding based on rough context window thresholds
    valueEl.classList.remove('warning', 'danger');
    if (totalTokens > 100000) {
        valueEl.classList.add('danger');
    } else if (totalTokens > 50000) {
        valueEl.classList.add('warning');
    }
}

/**
 * Opens the save prompt modal with category context.
 *
 * @param {string} category - The category being saved to
 */
function showSavePromptModal(category) {
    document.getElementById('savePromptModal').classList.add('active');
    document.getElementById('savePromptCategory').textContent = `Saving to: ${capitalize(category)}`;
    document.getElementById('promptName').value = '';
    document.getElementById('promptDescription').value = '';
    document.getElementById('promptName').focus();
}

/**
 * Hides the save prompt modal.
 */
function hideSavePromptModal() {
    document.getElementById('savePromptModal').classList.remove('active');
}

/**
 * Handles the Save button in the save prompt modal.
 * Creates a new prompt in the current tab's category and auto-selects it.
 */
function handleSavePromptConfirm() {
    const name = document.getElementById('promptName').value.trim();
    const description = document.getElementById('promptDescription').value.trim();
    const template = document.getElementById(`promptTextarea-${currentTab}`).value.trim();

    if (!name) {
        addLog('Please enter a prompt name', "warning");
        return;
    }

    if (!template) {
        addLog('Prompt template cannot be empty', "warning");
        return;
    }

    const promptData = { name, template, description };
    // Capture commentInstructions for amendment prompts
    if (currentTab === 'amendment') {
        const commentField = document.getElementById('commentInstructions');
        promptData.commentInstructions = commentField ? commentField.value.trim() : '';
    }
    const prompt = promptManager.addPrompt(currentTab, promptData);
    addLog(`Prompt saved: ${name} (${currentTab})`, "success");

    renderCategoryDropdown(currentTab);
    hideSavePromptModal();

    // Auto-select the saved prompt
    handleCategoryPromptSelect(currentTab, prompt.id);
    document.getElementById(`promptSelect-${currentTab}`).value = prompt.id;
}

/**
 * Handles deleting the currently selected prompt in a category.
 *
 * @param {string} category - The category to delete from
 */
function handleDeletePromptConfirm(category) {
    const select = document.getElementById(`promptSelect-${category}`);
    const promptId = select.value;

    if (!promptId) {
        addLog('No prompt selected to delete', "warning");
        return;
    }

    const prompt = promptManager.getPrompt(category, promptId);
    if (!prompt) return;

    promptManager.deletePrompt(category, promptId);
    addLog(`Prompt deleted: ${prompt.name} (${category})`, "success");

    renderCategoryDropdown(category);
    document.getElementById(`promptTextarea-${category}`).value = '';
    unsavedText[category] = '';

    updateDotIndicators();
    updatePanelButtons(category);
    updateTokenEstimate();
}

/**
 * Clears the textarea for a category without deactivating.
 * User must select "(None)" from dropdown to deactivate.
 *
 * @param {string} category - The category to clear
 */
function handleResetPrompt(category) {
    document.getElementById(`promptTextarea-${category}`).value = '';
    unsavedText[category] = '';
    addLog(`${capitalize(category)} prompt text cleared`, "info");
    updateTokenEstimate();
}

/**
 * Capitalizes the first letter of a string.
 * @param {string} str
 * @returns {string}
 */
function capitalize(str) {
    return str.charAt(0).toUpperCase() + str.slice(1);
}

// ============================================================================
// COMMENT STATUS BAR
// ============================================================================

/**
 * Updates the comment status bar visibility and pending count text.
 * Called by the integration plan (03-03) whenever the pending count changes.
 *
 * @param {number} count - Number of comments currently pending
 */
function updateCommentStatusBar(count) {
    const bar = document.getElementById('commentStatusBar');
    if (!bar) return;

    if (count === 0) {
        bar.style.display = 'none';
    } else {
        bar.style.display = 'flex';
        const text = document.getElementById('commentStatusText');
        if (text) {
            text.textContent = `${count} comment${count !== 1 ? 's' : ''} pending...`;
        }
    }
}

// ============================================================================
// CONNECTION & MODEL MANAGEMENT
// ============================================================================

/**
 * Tests connection to the active LLM backend and populates models.
 * Uses the unified llm-client.js testConnection function.
 */
async function testConnectionUI() {
    const indicator = document.getElementById("statusIndicator");
    const statusText = document.getElementById("statusText");
    const backendConfig = getActiveBackendConfig();
    const backendLabel = config.backend === 'vllm' ? 'vLLM' : 'Ollama';

    indicator.className = "status-indicator";
    statusText.textContent = "Connecting...";

    try {
        const result = await llmTestConnection(backendConfig);

        indicator.classList.add("connected");
        statusText.textContent = `${backendLabel}: Connected`;
        addLog(`Connected to ${backendLabel}! Found ${result.models.length} model(s).`, "success");

        // Populate model dropdown
        populateModels(result.models);

        if (config.backend === 'vllm') {
            // vLLM: set model to configured value as read-only
            const modelSelect = document.getElementById("modelSelect");
            modelSelect.innerHTML = '';
            const option = document.createElement('option');
            option.value = backendConfig.model;
            option.textContent = backendConfig.model;
            modelSelect.appendChild(option);
            modelSelect.disabled = true;
        }
    } catch (error) {
        indicator.classList.add("error");

        // Handle auth-specific errors
        if (error.message && (error.message.includes('401') || error.message.includes('403'))) {
            statusText.textContent = `${backendLabel}: API key required`;
            addLog(`${backendLabel} authentication failed: ${error.message}`, "error");
        } else {
            statusText.textContent = `${backendLabel}: Connection Error`;
            addLog(`${backendLabel} connection failed: ${error.message}`, "error");
        }

        console.error("Connection error:", error);
    }
}

/**
 * Populates the model dropdown from the /v1/models response.
 * Models use OpenAI format: { id: "model-name" }.
 */
function populateModels(models) {
    const select = document.getElementById("modelSelect");
    const activeModel = getActiveBackendConfig().model;
    select.innerHTML = '';

    if (models.length === 0) {
        select.innerHTML = '<option value="">No models available</option>';
        return;
    }

    models.forEach(model => {
        const option = document.createElement('option');
        option.value = model.id;
        option.textContent = model.id;
        if (model.id === activeModel) {
            option.selected = true;
        }
        select.appendChild(option);
    });

    // If selected model not in list, select first available
    const modelIds = models.map(m => m.id);
    if (!modelIds.includes(activeModel)) {
        const fallback = modelIds[0];
        config.backends[config.backend].model = fallback;
        select.value = fallback;
    }
}

// ============================================================================
// LLM INTEGRATION
// ============================================================================

/**
 * Handles the summary generation workflow.
 * Extracts all comments, sends to LLM with summary prompt, creates new document.
 * Fire-and-forget: user can switch modes immediately after triggering.
 */
async function handleSummaryGeneration() {
    // Phase 05.1 Plan 04a: legacy shared-action button IDs no longer exist
    // in HTML. Plan 04b will reintroduce inflight-button morph via a tracked
    // ref; for this intermediate state the loading visual is intentionally absent.
    try {
        addLog('Extracting document comments...', 'info');

        // 1. Extract all comments
        const comments = await extractAllComments();

        if (comments.length === 0) {
            addLog('No comments found in document. Add comments first, then generate summary.', 'warning');
            return;
        }

        addLog(`Found ${comments.length} comment(s). Sending to LLM...`, 'info');

        // 2. Extract document text if summary prompt uses {whole document} placeholder
        const summaryOpts = {};
        const activeSummaryPrompt = promptManager.getActivePrompt('summary');
        if (activeSummaryPrompt && activeSummaryPrompt.template.includes('{whole document}')) {
            const extraction = config.docExtraction || {};
            const richness = extraction.richness || 'structured';
            addLog(`Extracting document text (${richness})...`, 'info');
            summaryOpts.documentText = await extractDocumentStructured({ richness });
            addLog(`Document text extracted (${summaryOpts.documentText.length} chars, ~${estimateTokenCount(summaryOpts.documentText)} tokens)`, 'info');
        }

        // 3. Extract tracked changes if enabled and summary prompt uses {tracked changes} placeholder
        if (config.trackedChangesExtraction && activeSummaryPrompt && activeSummaryPrompt.template.includes('{tracked changes}')) {
            addLog('Extracting tracked changes (OOXML parsing)...', 'info');
            const tcResult = await extractTrackedChanges();
            addLog(`Tracked changes extracted (${tcResult.changes.length} change(s))`, 'info');

            // Format tracked changes for the prompt -- show before/after with author prominently
            let tcText = '';
            if (tcResult.changes.length > 0) {
                tcText = tcResult.changes.map((c, i) => {
                    const num = i + 1;
                    const author = c.author || 'Unknown';
                    const date = c.date || '';
                    const dateStr = date ? ` on ${date}` : '';

                    if (c.type === 'Replaced') {
                        return `[Change ${num}] REPLACED by ${author}${dateStr}:\n` +
                               `  BEFORE: "${c.beforeText}"\n` +
                               `  AFTER:  "${c.afterText}"` +
                               (c.paragraphText ? `\n  IN CLAUSE: "${c.paragraphText}"` : '');
                    } else if (c.type === 'Deleted') {
                        return `[Change ${num}] DELETED by ${author}${dateStr}:\n` +
                               `  REMOVED: "${c.text}"` +
                               (c.paragraphText ? `\n  IN CLAUSE: "${c.paragraphText}"` : '');
                    } else if (c.type === 'Added') {
                        return `[Change ${num}] ADDED by ${author}${dateStr}:\n` +
                               `  INSERTED: "${c.text}"` +
                               (c.paragraphText ? `\n  IN CLAUSE: "${c.paragraphText}"` : '');
                    } else if (c.type.startsWith('Moved')) {
                        return `[Change ${num}] ${c.type.toUpperCase()} by ${author}${dateStr}:\n` +
                               `  TEXT: "${c.text}"` +
                               (c.paragraphText ? `\n  IN CLAUSE: "${c.paragraphText}"` : '');
                    }
                    return `[Change ${num}] ${c.type} by ${author}${dateStr}: "${c.text}"`;
                }).join('\n\n');
            }

            if (tcText) {
                summaryOpts.trackedChangesText = tcText;
            } else if (tcResult.changes.length === 0) {
                summaryOpts.trackedChangesText = '(No tracked changes found in document)';
            }
        }

        // 4. Compose messages using PromptManager
        const messages = promptManager.composeSummaryMessages(comments, summaryOpts);

        if (messages.length === 0) {
            addLog('No summary prompt active. Select a Summary prompt first.', 'warning');
            return;
        }

        // 4. Send to LLM (flatten messages to single prompt for sendPrompt compatibility)
        const backendConfig = getActiveBackendConfig();
        let fullPrompt;
        if (messages.length >= 2 && messages[0].role === 'system') {
            fullPrompt = messages[0].content + '\n\n' + messages.slice(1).map(m => m.content).join('\n\n');
        } else {
            fullPrompt = messages.map(m => m.content).join('\n\n');
        }

        const llmResponse = await sendPrompt(backendConfig, fullPrompt, addLog);
        addLog(`Summary received (${llmResponse.length} chars). Creating document...`, 'info');

        // 5. Build HTML and create document
        // Get document title for the summary doc
        let docTitle = 'Comment Summary';
        try {
            await Word.run(async (context) => {
                const props = context.document.properties;
                props.load('title');
                await context.sync();
                if (props.title) {
                    docTitle = `Comment Summary - ${props.title}`;
                }
            });
        } catch (e) {
            // Title lookup failed -- use default
        }

        const html = buildSummaryHtml(llmResponse, comments, docTitle);
        await createSummaryDocument(html, docTitle, addLog);

        addLog('Summary document opened successfully.', 'success');

    } catch (error) {
        addLog(`Summary generation failed: ${error.message}`, 'error');
        console.error('Summary generation error:', error);
    } finally {
        // Plan 04b will restore inflight-button morph; for now just refresh
        // panel-button enable states.
        updatePanelButtons(CATEGORY.AMENDMENT);
        updatePanelButtons(CATEGORY.COMMENT);
        updatePanelButtons(CATEGORY.SUMMARY);
    }
}

async function handleReviewSelection({ category, withComment } = {}) {
    // Phase 05.1 Plan 04a: routes by explicit args — never reads any
    // active-mode global. STYLE.md "No Silent Failures": defensive guards
    // log with context and return.
    if (!category) {
        console.error('handleReviewSelection: missing category arg');
        addLog('Internal error: missing category', 'error');
        return;
    }
    if (!promptManager.getActivePrompt(category)) {
        addLog(`No active ${category} prompt`, 'warning');
        return;
    }

    if (isProcessing) {
        addLog("Already processing a request", "warning");
        return;
    }

    // Only block UI for amendment (synchronous) operations.
    // Comment selection is non-blocking (fire-and-forget via comment queue).
    const needsBlocking = (category === CATEGORY.AMENDMENT);

    try {
        if (needsBlocking) {
            isProcessing = true;
            // Plan 04b will morph the initiating button to a loading state
            // via the inflight-button helper; for now just refresh enable states.
            updatePanelButtons(CATEGORY.AMENDMENT);
        }

        // 1. Get Selection (Phase 05.2: gated comment enrichment via includeCommentsInSelection toggle)
        const includeComments = !!config.includeCommentsInSelection;
        let selectionText = "";
        let plainSelectionText = "";
        let enrichmentError = null;

        await Word.run(async (context) => {
            const selection = context.document.getSelection();
            selection.load("text");
            // OOXML fetch only when enrichment is requested (toggle ON) — saves a sync round-trip on the default path.
            const ooxmlResult = includeComments ? selection.getOoxml() : null;
            await context.sync();
            if (!selection.text || !selection.text.trim()) {
                throw new Error("Please select some text first.");
            }
            plainSelectionText = selection.text;

            if (!includeComments) {
                // Toggle OFF — today's behavior preserved verbatim. No extractor, no splicer.
                selectionText = plainSelectionText;
                return;
            }

            // Toggle ON: extract comments + replies whose anchors intersect the selection.
            // SAME Word.run owns selectionRange — research correction #4 / Open Q2: caller
            // controls the lifecycle, no nested Word.run, no cross-context tracking issues.
            let comments = [];
            try {
                comments = await extractCommentsOnRange(context, selection);
            } catch (err) {
                // STYLE.md No Silent Failures: log with context, fall back to plain text below.
                console.error('[handleReviewSelection] extractCommentsOnRange failed', { err });
                enrichmentError = err;
            }

            // Splice annotations using the OOXML walker. Pure function — runs after sync.
            if (!enrichmentError) {
                try {
                    selectionText = formatSelectionWithComments(ooxmlResult.value, comments);
                } catch (err) {
                    console.error('[handleReviewSelection] formatSelectionWithComments failed', {
                        err, ooxmlPrefix: String(ooxmlResult.value || '').slice(0, 200),
                    });
                    enrichmentError = err;
                }
            }
        });

        // STYLE.md No Silent Failures: surface degradation in the user-facing log.
        if (includeComments && enrichmentError) {
            addLog(`Comment enrichment failed (${enrichmentError.message}); falling back to plain selection.`, 'warning');
            selectionText = plainSelectionText;
        } else if (includeComments && selectionText.length > plainSelectionText.length) {
            const delta = selectionText.length - plainSelectionText.length;
            addLog(`Selection enriched with comment threads (+${delta} chars)`, 'info');
        }

        const activeBackend = getActiveBackendConfig();
        addLog(`Processing selection (${selectionText.length} chars) via ${activeBackend.model}...`, "info");

        // 2. Route by explicit args (never by active-mode read)
        if (category === CATEGORY.AMENDMENT) {
            if (withComment) {
                const commentInstructions = document.getElementById('commentInstructions').value.trim();
                await handleMergedAmendmentComment(selectionText, commentInstructions, activeBackend);
            } else {
                await handleAmendmentOnly(selectionText, activeBackend);
            }
        } else if (category === CATEGORY.COMMENT) {
            if (!supportsComments) {
                addLog("Comment features require Word API 1.4", "warning");
            } else {
                fireCommentRequest(selectionText, {
                    config: activeBackend,
                    sendPromptFn: sendPrompt,
                    promptManager: promptManager,
                    commentQueue: commentQueue,
                    log: addLog,
                    addLogWithRetryFn: addLogWithRetry,
                    updateStatusBarFn: updateCommentStatusBar
                });
            }
        } else {
            // Summary is never dispatched through here (the dispatcher routes
            // summary:document:plain directly to handleSummaryGeneration).
            console.warn('handleReviewSelection: unsupported category', category);
        }

    } catch (error) {
        addLog(`Error: ${error.message}`, "error");
    } finally {
        if (needsBlocking) {
            isProcessing = false;
            updatePanelButtons(CATEGORY.AMENDMENT);
            updatePanelButtons(CATEGORY.COMMENT);
            updatePanelButtons(CATEGORY.SUMMARY);
        }
    }
}

/**
 * Handles amendment-only submission (no comment instructions).
 * Sends amendment prompt to LLM and applies diff as tracked changes.
 */
async function handleAmendmentOnly(selectionText, activeBackend) {
    const messages = promptManager.composeMessages(selectionText, 'amendment');

    let fullPrompt;
    if (messages.length === 2) {
        fullPrompt = messages[0].content + '\n\n' + messages[1].content;
    } else if (messages.length === 1) {
        fullPrompt = messages[0].content;
    } else {
        throw new Error("No prompt composed -- check active prompts");
    }

    const backendConfig = getActiveBackendConfig();
    const rawResponse = await sendPrompt(backendConfig, fullPrompt, addLog);
    const response = stripMarkdown(rawResponse, addLog);

    addLog(`LLM Response received [${backendConfig.model}]`, "success");
    addLog(`Response: ${response.substring(0, 100)}${response.length > 100 ? '...' : ''}`, "info");

    addLog("Applying changes...", "info");

    await Word.run(async (context) => {
        const selection = context.document.getSelection();
        if (Word.ChangeTrackingMode) {
            context.document.changeTrackingMode = config.trackChangesEnabled
                ? Word.ChangeTrackingMode.trackAll
                : Word.ChangeTrackingMode.off;
        }
        if (config.lineDiffEnabled) {
            await applySentenceDiffStrategy(context, selection, selectionText, response, addLog);
        } else {
            await applyTokenMapStrategy(context, selection, selectionText, response, addLog);
        }
    });

    addLog("Changes applied successfully", "success");
}

/**
 * Handles merged amendment + comment submission.
 * Sends a single merged prompt to LLM, parses delimited response,
 * applies amendment as tracked changes and inserts comment on selection.
 * Falls back to a second LLM call if delimiters are missing.
 */
async function handleMergedAmendmentComment(selectionText, commentInstructions, activeBackend) {
    const messages = promptManager.composeMergedMessages(selectionText, commentInstructions);

    let fullPrompt;
    if (messages.length === 2) {
        fullPrompt = messages[0].content + '\n\n' + messages[1].content;
    } else if (messages.length === 1) {
        fullPrompt = messages[0].content;
    } else {
        throw new Error("No prompt composed -- check active prompts");
    }

    const backendConfig = getActiveBackendConfig();
    addLog(`Sending merged amendment + comment request [${backendConfig.model}]...`, "info");
    const response = await sendPrompt(backendConfig, fullPrompt, addLog);

    addLog(`LLM Response received [${backendConfig.model}]`, "success");
    addLog(`Response: ${response.substring(0, 100)}${response.length > 100 ? '...' : ''}`, "info");

    // Parse delimited response
    let parsed = parseDelimitedResponse(response);

    // Fallback: if delimiters not found, try a second LLM call to classify
    if (parsed.amendment === null) {
        addLog("Response missing delimiters, attempting to classify...", "info");
        const fallbackMessages = buildFallbackClassificationPrompt(response, selectionText);
        const fallbackPrompt = fallbackMessages[0].content + '\n\n' + fallbackMessages[1].content;

        try {
            const fallbackResponse = await sendPrompt(backendConfig, fallbackPrompt, addLog);
            parsed = parseDelimitedResponse(fallbackResponse);

            if (parsed.amendment === null) {
                // Still no delimiters -- treat entire original response as amendment (best-effort)
                addLog("Could not split response into amendment and comment", "warning");
                parsed = { amendment: response.trim(), comment: null, raw: response };
            }
        } catch (fallbackError) {
            // Fallback call failed -- use original response as amendment
            addLog(`Fallback classification failed: ${fallbackError.message}`, "warning");
            parsed = { amendment: response.trim(), comment: null, raw: response };
        }
    }

    // Apply amendment as tracked changes (strip markdown artifacts first)
    if (parsed.amendment) {
        parsed.amendment = stripMarkdown(parsed.amendment, addLog);
        addLog("Applying amendment changes...", "info");

        await Word.run(async (context) => {
            const selection = context.document.getSelection();
            if (Word.ChangeTrackingMode) {
                context.document.changeTrackingMode = config.trackChangesEnabled
                    ? Word.ChangeTrackingMode.trackAll
                    : Word.ChangeTrackingMode.off;
            }
            if (config.lineDiffEnabled) {
                await applySentenceDiffStrategy(context, selection, selectionText, parsed.amendment, addLog);
            } else {
                await applyTokenMapStrategy(context, selection, selectionText, parsed.amendment, addLog);
            }
        });

        addLog("Amendment changes applied successfully", "success");
    }

    // Insert comment on selection if available and supported
    if (parsed.comment && supportsComments) {
        addLog("Inserting comment...", "info");

        try {
            await Word.run(async (context) => {
                const selection = context.document.getSelection();
                selection.load("text");
                await context.sync();

                // Insert comment directly on the current selection range
                const contentRange = selection.getRange();
                contentRange.insertComment(parsed.comment);
                await context.sync();
            });

            addLog("Comment inserted successfully", "success");
        } catch (commentError) {
            // Comment insertion failed -- log the comment text so it is not lost
            addLog(`Comment insertion failed: ${commentError.message}. Comment text: "${parsed.comment}"`, "warning");
        }
    } else if (parsed.comment && !supportsComments) {
        addLog(`Comment generated but Word API 1.4 not available. Comment: "${parsed.comment}"`, "warning");
    }
}

// ============================================================================
// WHOLE-DOCUMENT PROCESSING
// ============================================================================

/**
 * Updates the process progress bar with current chunk progress.
 * @param {object} progress - Progress object from orchestrator
 * @param {number} progress.completed - Completed chunks
 * @param {number} progress.failed - Failed chunks
 * @param {number} progress.total - Total chunks
 * @param {number} progress.percentComplete - Percentage complete
 * @param {number} progress.estimatedSecondsRemaining - ETA in seconds
 */
function updateProcessProgress(progress) {
    const fill = document.getElementById('progressFill');
    const text = document.getElementById('progressText');
    if (fill) fill.style.width = `${progress.percentComplete}%`;
    if (text) {
        text.textContent = `Processing: ${progress.completed + progress.failed}/${progress.total} chunks`;
        if (progress.estimatedSecondsRemaining > 0) {
            text.textContent += ` (~${progress.estimatedSecondsRemaining}s remaining)`;
        }
    }
}

/**
 * Handles the full whole-document processing workflow.
 * Parses document, chunks it, extracts context, processes chunks in parallel,
 * applies results as tracked changes/comments, and shows summary.
 * Double-click acts as cancel.
 */
async function handleProcessDocument({ category, withComment } = {}) {
    // Cancel path — preserved unchanged. Plan 04b will refine the cancel-morph
    // wiring, but the AbortController contract stays the same.
    if (isProcessingDoc && processDocController) {
        processDocController.abort();
        addLog('Cancelling document processing...', 'warning');
        return;
    }

    // Phase 05.1 Plan 04a: routes by explicit args — never reads any
    // active-mode global.
    if (!category) {
        console.error('handleProcessDocument: missing category arg');
        addLog('Internal error: missing category', 'error');
        return;
    }
    if (!promptManager.getActivePrompt(category)) {
        addLog(`No active ${category} prompt`, 'warning');
        return;
    }

    // Determine which panel button initiated this run (used for cancel-morph).
    // data-action mapping: amendment -> amend-document or amend-comment-document;
    // comment -> comment-document; summary -> summary-document.
    let initiatingAction;
    if (category === CATEGORY.AMENDMENT) {
        initiatingAction = withComment ? 'amend-comment-document' : 'amend-document';
    } else if (category === CATEGORY.COMMENT) {
        initiatingAction = 'comment-document';
    } else {
        initiatingAction = 'summary-document';
    }
    const initiatingBtn = document.querySelector(
        `[data-panel="${category}"][data-action="${initiatingAction}"]`
    );

    // Block all panel buttons; morph the initiating button to "Cancel".
    isProcessingDoc = true;
    processDocController = new AbortController();
    setInflightButton(initiatingBtn, { category, withComment });

    // Show progress bar, hide comment status bar
    const progressBar = document.getElementById('processProgressBar');
    const commentBar = document.getElementById('commentStatusBar');
    progressBar.style.display = 'flex';
    commentBar.style.display = 'none';

    try {
        // Step 1: Parse document
        addLog('Parsing document...', 'info');
        const docModel = await parseDocument();
        addLog(`Found ${docModel.paragraphs.length} paragraphs (~${docModel.totalTokens} tokens)`, 'info');

        // Step 2: Chunk document
        const chunks = chunkDocument(docModel, { maxTokens: 6000 });
        addLog(`Split into ${chunks.length} chunks`, 'info');

        // Step 3: Extract context
        const documentContext = extractContext(docModel);
        addLog(`Extracted ${documentContext.definitions.length} definitions, ${documentContext.outline.length} headings`, 'info');

        // Step 4: Bookmark chunk ranges
        const bookmarkMap = await bookmarkChunkRanges(chunks);

        // Step 5: Process chunks in parallel
        const backendConfig = getActiveBackendConfig();
        const concurrency = chunks.some(c => c.tokenCount > 8000) ? 4 : 6;
        const commentInstructions = document.getElementById('commentInstructions')?.value?.trim() || '';

        const results = await processChunksParallel(chunks, {
            config: backendConfig,
            promptManager: promptManager,
            documentContext: documentContext,
            log: addLog,
            onProgress: updateProcessProgress,
            signal: processDocController.signal,
            concurrency: concurrency,
            timeoutMs: 300000,
            commentInstructions: commentInstructions
        });

        // Step 6: Apply results to document
        addLog('Applying changes to document...', 'info');
        const granularity = parseInt(document.getElementById('commentGranularity')?.value || '0', 10);
        const applicationResult = await applyChunkResults(results, bookmarkMap, {
            trackChangesEnabled: config.trackChangesEnabled,
            lineDiffEnabled: config.lineDiffEnabled,
            log: addLog,
            commentGranularity: granularity
        });

        // Step 7: Cleanup
        await cleanupBookmarks(bookmarkMap);

        // Step 8: Summary log
        const failed = results.filter(r => r.status === 'rejected').length;
        const cancelled = results.filter(r => r.status === 'cancelled').length;
        addLog(
            `Document processed: ${chunks.length} chunks, ` +
            `${applicationResult.amendmentsApplied} amendments applied, ` +
            `${applicationResult.commentsInserted} comments inserted` +
            (failed > 0 ? `, ${failed} chunks failed` : '') +
            (cancelled > 0 ? `, ${cancelled} chunks cancelled` : ''),
            failed > 0 ? 'warning' : 'success'
        );

        // Show "Retry All Failed" link if failures exist
        if (failed > 0) {
            const failedChunks = results.filter(r => r.status === 'rejected');
            addLogWithRetry(
                `${failed} chunk(s) failed. Click to retry failed chunks.`,
                'warning',
                () => retryFailedChunks(failedChunks, bookmarkMap, backendConfig)
            );
        }

    } catch (error) {
        if (error.name === 'AbortError') {
            addLog('Document processing cancelled. Already-applied changes remain in the document.', 'warning');
        } else {
            addLog(`Document processing failed: ${error.message}`, 'error');
            console.error('Process document error:', error);
        }
    } finally {
        isProcessingDoc = false;
        processDocController = null;
        clearInflightButton(); // restore label, drop .cancel-mode + dataset before re-enabling peers
        progressBar.style.display = 'none';
        commentBar.style.display = commentQueue.count > 0 ? 'flex' : 'none';
        updatePanelButtons(CATEGORY.AMENDMENT);
        updatePanelButtons(CATEGORY.COMMENT);
        updatePanelButtons(CATEGORY.SUMMARY);
    }
}

/**
 * Retries processing only the failed chunks.
 * Re-runs the orchestrator on the failed chunk subset and applies results.
 *
 * @param {Array} failedResults - Array of ChunkResult objects with status 'rejected'
 * @param {Map} bookmarkMap - Original chunkId -> bookmarkName map
 * @param {object} backendConfig - Backend configuration
 */
async function retryFailedChunks(failedResults, bookmarkMap, backendConfig) {
    addLog(`Retrying ${failedResults.length} failed chunk(s)...`, 'info');

    isProcessingDoc = true;
    processDocController = new AbortController();
    const progressBar = document.getElementById('processProgressBar');

    // Retry runs without a panel-button morph: it's triggered from the
    // activity-log retry link, not from a panel button click, so there's
    // no initiating button to morph. _inflightBtn stays null; cancel during
    // retry is currently not exposed in the UI (out of scope for Plan 04b).
    // updatePanelButtons reflects the disabled/blocked state via the
    // dispatcher's isProcessingDocRef pre-flight.
    updatePanelButtons(CATEGORY.AMENDMENT);
    updatePanelButtons(CATEGORY.COMMENT);
    updatePanelButtons(CATEGORY.SUMMARY);
    progressBar.style.display = 'flex';

    try {
        // Reconstruct chunks from failed results for re-processing
        const retryChunks = failedResults.map(r => ({
            id: r.chunkId,
            text: r.originalText || '',
            tokenCount: r.originalText ? Math.ceil(r.originalText.length / 4) : 0,
            overlapText: ''
        }));

        const commentInstructions = document.getElementById('commentInstructions')?.value?.trim() || '';

        const results = await processChunksParallel(retryChunks, {
            config: backendConfig,
            promptManager: promptManager,
            documentContext: null,
            log: addLog,
            onProgress: updateProcessProgress,
            signal: processDocController.signal,
            concurrency: 4,
            timeoutMs: 300000,
            commentInstructions: commentInstructions
        });

        const granularity = parseInt(document.getElementById('commentGranularity')?.value || '0', 10);
        const applicationResult = await applyChunkResults(results, bookmarkMap, {
            trackChangesEnabled: config.trackChangesEnabled,
            lineDiffEnabled: config.lineDiffEnabled,
            log: addLog,
            commentGranularity: granularity
        });

        const stillFailed = results.filter(r => r.status === 'rejected').length;
        addLog(
            `Retry complete: ${applicationResult.amendmentsApplied} amendments, ` +
            `${applicationResult.commentsInserted} comments` +
            (stillFailed > 0 ? `, ${stillFailed} still failed` : ''),
            stillFailed > 0 ? 'warning' : 'success'
        );

    } catch (error) {
        if (error.name === 'AbortError') {
            addLog('Retry cancelled.', 'warning');
        } else {
            addLog(`Retry failed: ${error.message}`, 'error');
        }
    } finally {
        isProcessingDoc = false;
        processDocController = null;
        clearInflightButton(); // safety: no-op if _inflightBtn === null (typical retry path)
        progressBar.style.display = 'none';
        updatePanelButtons(CATEGORY.AMENDMENT);
        updatePanelButtons(CATEGORY.COMMENT);
        updatePanelButtons(CATEGORY.SUMMARY);
    }
}

// ============================================================================
// VERIFICATION SCRIPT
// ============================================================================

async function runVerification() {
    const btn = document.getElementById("runVerificationBtn");

    try {
        btn.classList.add("loading");
        btn.disabled = true;
        addLog("Loading verification script...", "info");

        const module = await import('../scripts/verify-word-api.js');
        await module.runAllVerifications(addLog);

    } catch (error) {
        addLog(`Verification Error: ${error.message}`, "error");
        console.error(error);
    } finally {
        btn.classList.remove("loading");
        btn.disabled = false;
    }
}

// ============================================================================
// LOGGING
// ============================================================================

function addLog(message, type = "info") {
    const logsDiv = document.getElementById("logs");
    const entry = document.createElement("div");
    const timestamp = new Date().toLocaleTimeString();

    entry.className = `log-${type}`;
    entry.textContent = `[${timestamp}] ${message}`;

    logsDiv.appendChild(entry);
    logsDiv.scrollTop = logsDiv.scrollHeight;

    console.log(`[${type.toUpperCase()}] ${message}`);

    // Send to server log (best effort)
    fetch('/log', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ message, type, timestamp: new Date().toISOString() })
    }).catch(() => { });
}

/**
 * Extended version of addLog that appends a clickable "Retry" link to the log entry.
 * Used for failed comment requests where the user can retry the operation.
 *
 * @param {string} message - The log message text
 * @param {string} type - Log type: "info", "success", "warning", "error"
 * @param {Function} retryCallback - Function to call when Retry is clicked
 */
function addLogWithRetry(message, type, retryCallback) {
    const logsDiv = document.getElementById("logs");
    const entry = document.createElement("div");
    const timestamp = new Date().toLocaleTimeString();
    entry.className = `log-${type}`;

    const msgSpan = document.createElement("span");
    msgSpan.textContent = `[${timestamp}] ${message} `;
    entry.appendChild(msgSpan);

    if (retryCallback) {
        const retryLink = document.createElement("a");
        retryLink.textContent = "Retry";
        retryLink.href = "#";
        retryLink.className = "retry-link";
        retryLink.onclick = (e) => {
            e.preventDefault();
            retryCallback();
            entry.remove();  // Remove the error log entry on retry
        };
        entry.appendChild(retryLink);
    }

    logsDiv.appendChild(entry);
    logsDiv.scrollTop = logsDiv.scrollHeight;

    console.log(`[${type.toUpperCase()}] ${message}`);

    // Send to server log (best effort)
    fetch('/log', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ message, type, timestamp: new Date().toISOString() })
    }).catch(() => { });
}

function clearLogs() {
    document.getElementById("logs").innerHTML = "";
}
