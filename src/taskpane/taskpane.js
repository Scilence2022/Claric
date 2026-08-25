/**
 * Taskpane Bootstrap
 *
 * Entry point for the chat-driven taskpane. Responsibilities are limited to:
 * wiring the UI modules together, detecting Word API capabilities, and
 * kicking off the initial connection test. All logic lives in:
 *
 *   app-state.js     - config, PromptManager, processing flags
 *   skills.js        - skill registry (built-in + custom prompts)
 *   conversation.js  - turn routing and orchestration
 *   word-actions.js  - document/LLM pipelines
 *   ui/*             - chat view, input bar, welcome, settings, status bar
 *
 * `normalizeConfig` is re-exported from app-state.js as a test seam
 * (tests/config-persistence.spec.js).
 *
 * Module-level code only imports modules and registers Office.onReady, so the
 * module stays importable under Jest (node/jsdom) without Word globals.
 */

// Import CSS for webpack to bundle
import './taskpane.css';

import { appState, loadSettings, getActiveBackendConfig } from './app-state.js';
import { BUILTIN_SKILLS, listSkills } from './skills.js';
import { createConversation } from './conversation.js';
import { watchSelection } from './word-actions.js';
import * as chatView from './ui/chat-view.js';
import { renderWelcomeChips } from './ui/welcome.js';
import { initInputBar } from './ui/input-bar.js';
import { initSettings, openSettings, testConnectionUI } from './ui/settings-view.js';
import { initStatusBar, addLog, addLogWithRetry, updateCommentStatusBar } from './ui/status-bar.js';
import { getProviderPreset } from '../lib/providers.js';

export { normalizeConfig } from './app-state.js';

if (typeof Office !== 'undefined') {
    Office.onReady((info) => {
        if (info.host === Office.HostType.Word) {
            initialize();
        }
    });
}

/**
 * Wires the modules together and starts the app. Called from Office.onReady.
 */
function initialize() {
    // Load saved settings (localStorage key unchanged: wordAI.config)
    loadSettings(appState, addLog);
    appState.promptManager.loadState();
    appState.log = addLog;

    // Status bar (log drawer, comment pending bar)
    initStatusBar();

    // Chat view
    chatView.initChatView();

    // Input bar + conversation orchestration
    const input = initInputBar({
        onSubmit: (text) => conversation.submit(text),
        onCancel: () => conversation.cancel(),
        getSkills: () => listSkills(appState.promptManager),
        onOpenSettings: openSettings,
    });

    const conversation = createConversation({
        appState,
        view: {
            addUserMessage: chatView.addUserMessage,
            createAssistantMessage: chatView.createAssistantMessage,
            addSystemNote: chatView.addSystemNote,
            hideWelcome: chatView.hideWelcome,
            renderWelcome: chatView.showWelcome,
            clearChat: chatView.clearChat,
        },
        input,
        log: addLog,
        logWithRetry: addLogWithRetry,
        updateStatusBar: updateCommentStatusBar,
    });

    // Settings slide-over (provider settings + prompt management)
    initSettings({ onConfigChanged: updateModelPill });

    // Header buttons
    document.getElementById('newChatBtn').addEventListener('click', () => conversation.newChat());
    document.getElementById('infoBtn').addEventListener('click', showAbout);

    // Welcome skill chips fill the input with the slash command
    renderWelcomeChips(BUILTIN_SKILLS, (skill) => {
        input.setValue(`${skill.slash} `);
        input.focus();
    });

    updateModelPill();

    // Live selection preview above the input bar (selection text is also
    // added to the next turn's context at submit time)
    watchSelection((text) => input.setSelectionPreview(text));

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
    } catch { /* detection failed */ }
    addLog(`Word API version: ${detectedVersion}`, 'info');

    // Detect WordApi 1.4 support for comment features (graceful degradation:
    // comment-category turns report the requirement instead of failing).
    if (typeof Office !== 'undefined' && Office.context && Office.context.requirements) {
        appState.supportsComments = Office.context.requirements.isSetSupported('WordApi', '1.4');
    }
    if (!appState.supportsComments) {
        addLog('Comment features unavailable (requires Word API 1.4)', 'info');
    }

    // Auto-test connection and load models
    testConnectionUI();

    addLog('Claric initialized.', 'info');
    input.focus();
}

/**
 * Updates the model pill in the input bar from the active provider config.
 */
function updateModelPill() {
    const preset = getProviderPreset(appState.config.backend);
    const backendConfig = getActiveBackendConfig(appState);
    const label = preset ? preset.label : appState.config.backend;
    const pill = document.getElementById('modelPill');
    if (pill) {
        pill.textContent = `${label}: ${backendConfig.model || '(no model)'}`;
        pill.title = 'Open settings';
    }
}

/** Shows a short about message in the chat. */
function showAbout() {
    chatView.addSystemNote(
        'Claric — your redlining scribe for Word.\n' +
        'Type "/" for skills (/copy-edit, /check-doc, /flag-issues, /summarize-contract, ...), ' +
        'select text and type an instruction to edit it as tracked changes, or just ask a question about the document.'
    );
}
