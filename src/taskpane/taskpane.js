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
 * Module-level code only imports modules and registers Office.onReady, so the
 * module stays importable under Jest (node/jsdom) without Word globals.
 */

// Import CSS for webpack to bundle
import './taskpane.css';

import { appState, loadSettings, getActiveBackendConfig, persistSettings } from './app-state.js';
import { listSkills } from './skills.js';
import { createConversation } from './conversation.js';
import { watchSelection, revealTextSnippet } from './word-actions.js';
import { reapOrphanChunkBookmarks } from '../lib/reassembler.js';
import * as chatView from './ui/chat-view.js';
import { renderWelcomeChips, selectWelcomeSkills } from './ui/welcome.js';
import { initInputBar } from './ui/input-bar.js';
import { initSettings, openSettings, testConnectionUI } from './ui/settings-view.js';
import { initStatusBar, addLog, addLogWithRetry, updateCommentStatusBar, toggleLogDrawer } from './ui/status-bar.js';
import { initHistoryView, openHistory } from './ui/history-view.js';
import { listSessions, loadSession as loadStoredSession, saveSession, deleteSession } from './sessions.js';
import { getProviderPreset } from '../lib/providers.js';
import { getHostPlatform } from '../lib/platform.js';

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

    // Citation pills rebuilt from a restored session carry no closures — give
    // them the same reveal action the live turn uses.
    chatView.setCitationSelectHandler((searchText) => {
        revealTextSnippet({ log: addLog }, searchText);
    });

    // History slide-over (sessions saved to localStorage; survives reloads).
    initHistoryView({
        onLoadSession: (session) => {
            chatView.setCurrentSession(session);
            input.focus();
        },
        onDeleteSession: (id) => {
            try { deleteSession(id); } catch (e) { addLog(`Delete session failed: ${e.message}`, 'error'); }
        },
        onNewChat: () => conversation.newChat(),
    });

    /**
     * Persists the live session after a turn completes. Called by the
     * conversation orchestrator via createConversation's onTurnCommitted dep.
     * @param {object|null} session
     */
    function persistCurrentSession(session) {
        if (!session || !Array.isArray(session.messages) || session.messages.length === 0) return;
        try {
            saveSession(session.messages, { id: session.id, title: session.title });
        } catch (e) {
            addLog(`Save session failed: ${e.message}`, 'error');
        }
    }

    // Input bar + conversation orchestration
    const input = initInputBar({
        onSubmit: (text, attachments) => conversation.submit(text, attachments),
        onCancel: () => conversation.cancel(),
        getSkills: () => listSkills(appState.promptManager),
        onOpenSettings: openSettings,
        onLog: addLog,
        getAutoApply: () => appState.config.autoApplyChanges === true,
        setAutoApply: (value) => {
            appState.config.autoApplyChanges = value === true;
            persistSettings(appState);
            addLog(value
                ? 'Auto-apply enabled — proposed changes will be applied automatically as tracked changes.'
                : 'Auto-apply disabled — proposed changes will wait for your review.',
                'info');
        },
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
            getCurrentSession: chatView.getCurrentSession,
        },
        input,
        log: addLog,
        logWithRetry: addLogWithRetry,
        updateStatusBar: updateCommentStatusBar,
        onTurnCommitted: persistCurrentSession,
    });

    // Proposal cards settle after their turn has already been finalized
    // (Apply/Reject clicks), so late state changes must re-persist the
    // session or history would keep them at "pending" forever.
    chatView.setProposalStateChangeHandler(() => {
        persistCurrentSession(chatView.getCurrentSession());
    });

    // Settings slide-over (provider settings + prompt management)
    initSettings({ onConfigChanged: updateModelPill });

    // Header buttons
    document.getElementById('historyBtn').addEventListener('click', openHistory);
    document.getElementById('logBtn').addEventListener('click', toggleLogDrawer);
    document.getElementById('newChatBtn').addEventListener('click', () => conversation.newChat());
    document.getElementById('infoBtn').addEventListener('click', showAbout);

    // Welcome skill chips fill the input with the slash command. The
    // built-ins only (hardcoded before) hid the reserved /mcp skill and
    // imported SKILL.md packages — show the full "get started" set.
    renderWelcomeChips(selectWelcomeSkills(listSkills(appState.promptManager)), (skill) => {
        input.setValue(`${skill.slash} `);
        input.focus();
    });

    updateModelPill();

    // Restore the most recent session if one exists; otherwise stay on the
    // welcome page (the current chat-view already shows the welcome by default).
    const recent = listSessions();
    if (recent.length > 0) {
        const full = loadStoredSession(recent[0].id);
        if (full) chatView.setCurrentSession(full);
    }

    // Live selection preview above the input bar (text snippet + image
    // thumbnails; image-bearing selections enter the turn's context via the
    // image tool session at submit time)
    watchSelection((content) => input.setSelectionPreview(content));

    // Startup hygiene: chunk bookmarks from a run interrupted by a reload
    // (or a crash) are unreferenced — no in-memory apply/discard closure
    // survives — so reap them before any new run creates its own. Best
    // effort; never blocks startup.
    reapOrphanChunkBookmarks(addLog);

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
        // WordApi 1.3 covers table creation (Body/Range.insertTable, table
        // style/headerRowCount); table turns degrade gracefully without it.
        appState.supportsTables = Office.context.requirements.isSetSupported('WordApi', '1.3');
    }
    if (!appState.supportsComments) {
        addLog('Comment features unavailable (requires Word API 1.4)', 'info');
    }
    if (!appState.supportsTables) {
        addLog('Table creation unavailable (requires Word API 1.3)', 'info');
    }

    // Detect the host platform (PC/Mac/OfficeOnline/...) — table row-level
    // revisions are only recorded by Word desktop, so table edits branch on it.
    appState.platform = getHostPlatform();
    addLog(`Host platform: ${appState.platform}`, 'info');

    // Auto-test connection and load models
    testConnectionUI();

    // Log build fingerprint (dist content hash + UTC timestamp). Fire-and-forget:
    // the asset may be served from a stale cache or unreachable; a missing
    // build-info must never slow or surface anything to the user.
    fetch(new URL('build-info.json', location.href).href, { cache: 'no-store' })
        .then((res) => (res.ok ? res.json() : null))
        .then((info) => {
            if (!info || typeof info.hash !== 'string') return;
            const builtAt = typeof info.builtAt === 'string' ? info.builtAt : '';
            addLog(
                builtAt
                    ? `Claric build: ${info.hash} (${info.appVersion || 'unknown'}, ${builtAt})`
                    : `Claric build: ${info.hash} (${info.appVersion || 'unknown'})`,
                'info'
            );
        })
        .catch(() => { /* silent: build info missing or blocked is not an error */ });

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
        'Type "/" for slash commands (/polish, /simplify, /shorten, /expand, /translate, /check-clarity, ' +
        '/check-consistency, /action-items, /executive-summary, /key-points, ...), ' +
        'select text and type an instruction to edit it as tracked changes, or just ask a question about the document.'
    );
}
