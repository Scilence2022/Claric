/**
 * App State Module
 *
 * Central mutable state for the chat-driven taskpane: LLM configuration,
 * the shared PromptManager instance, processing flags, and the comment queue.
 * Extracted from the monolithic taskpane.js so every module reads/writes one
 * well-defined state object instead of module-level closures.
 *
 * Module-level code is side-effect free (safe to import under Jest/node).
 *
 * @module app-state
 */

import { PromptManager } from '../lib/prompt-manager.js';
import { CommentQueue } from '../lib/comment-queue.js';
import { KNOWN_PROVIDERS, defaultProviderConfig } from '../lib/providers.js';
import { TOOL_LOOP_LIMITS } from '../lib/tool-registry.js';

const KNOWN_RICHNESS = ['plain', 'headings', 'structured'];

/**
 * Builds the default configuration object.
 * Defaults come from provider presets, overridable via Settings/localStorage.
 *
 * @returns {object} A fresh default config
 */
export function defaultConfig() {
    return {
        backend: 'ollama',
        trackChangesEnabled: true,
        lineDiffEnabled: false,
        docExtraction: {
            richness: 'structured'
        },
        trackedChangesExtraction: false,
        commentGranularity: 0,
        includeCommentsInSelection: false,
        mcpServers: [],
        mcpStepBudget: TOOL_LOOP_LIMITS.MAX_STEPS_DEFAULT,
        providers: defaultProviderConfig()
    };
}

/**
 * The single shared application state. Mutated in place by UI handlers and
 * conversation orchestration; never re-assigned.
 */
export const appState = {
    config: defaultConfig(),
    promptManager: new PromptManager(),
    isProcessing: false,        // selection-scope op or chat Q&A in flight
    isProcessingDoc: false,     // document-scope chunked run in flight
    isProcessingSummary: false, // summary generation in flight
    processDocController: null, // AbortController for document-scope runs
    chatController: null,       // AbortController for chat Q&A streaming
    supportsComments: false,    // set during initialize() via WordApi 1.4 check
    supportsTables: false,      // set during initialize() via WordApi 1.3 check
    platform: 'unknown',        // set during initialize() via Office.context.platform
    // Logging indirection so the CommentQueue can be constructed before the
    // DOM-backed addLog exists; initialize() replaces appState.log.
    log: (message, type = 'info') => console.log(`[${type.toUpperCase()}] ${message}`),
    commentQueue: null,
};
appState.commentQueue = new CommentQueue((message, type) => appState.log(message, type));

/**
 * Returns the config object for the currently selected provider.
 *
 * @param {object} [state] - App state (defaults to the shared appState)
 * @returns {{ url: string, apiKey: string, model: string, apiPath: string }}
 */
export function getActiveBackendConfig(state = appState) {
    return state.config.providers[state.config.backend];
}

/**
 * Returns a debounced wrapper around fn. Repeated calls within waitMs
 * collapse into a single trailing invocation.
 *
 * @param {Function} fn
 * @param {number} waitMs
 * @returns {Function}
 */
export function debounce(fn, waitMs) {
    let timer = null;
    return function debounced(...args) {
        clearTimeout(timer);
        timer = setTimeout(() => fn.apply(this, args), waitMs);
    };
}

/**
 * Validates one per-provider entry against its preset defaults.
 *
 * @param {object} defaults - The built-in default entry for this provider
 * @param {object} saved - The persisted entry (possibly partial/corrupt)
 * @returns {{url: string, apiKey: string, model: string, apiPath: string}}
 */
function normalizeProviderEntry(defaults, saved) {
    return {
        url: typeof saved.url === 'string' && saved.url ? saved.url : defaults.url,
        apiKey: typeof saved.apiKey === 'string' ? saved.apiKey : '',
        model: typeof saved.model === 'string' && saved.model ? saved.model : defaults.model,
        apiPath: typeof saved.apiPath === 'string' && saved.apiPath ? saved.apiPath : defaults.apiPath,
    };
}

/**
 * Merges a persisted config object onto the defaults, field by field.
 *
 * A shallow spread ({ ...config, ...parsed }) lets a corrupt or hand-edited
 * localStorage entry drop whole sections (e.g. replace `providers` with a
 * partial object), which later crashes on config.providers[backend].url.
 * Every field is validated here so loadSettings can never produce a config
 * that getActiveBackendConfig() cannot serve.
 *
 * Supports the pre-0.4.0 `backends` shape (Ollama/vLLM only) by migrating
 * it into the `providers` map.
 *
 * @param {object} defaults - The built-in defaults
 * @param {object} parsed - Whatever JSON.parse returned from localStorage
 * @returns {object} A fully-populated config
 */
export function normalizeConfig(defaults, parsed) {
    const out = { ...defaults };

    if (typeof parsed.backend === 'string' && KNOWN_PROVIDERS.includes(parsed.backend)) {
        out.backend = parsed.backend;
    }

    // New shape: providers map, merged entry-by-entry over the presets so
    // unknown/partial provider entries never break getActiveBackendConfig().
    if (parsed.providers && typeof parsed.providers === 'object') {
        for (const name of KNOWN_PROVIDERS) {
            const savedProvider = parsed.providers[name];
            if (savedProvider && typeof savedProvider === 'object') {
                out.providers[name] = normalizeProviderEntry(defaults.providers[name], savedProvider);
            }
        }
    } else if (parsed.backends && typeof parsed.backends === 'object') {
        // Legacy shape (v0.3.x): backends map with ollama/vllm entries.
        for (const name of KNOWN_PROVIDERS) {
            const savedBackend = parsed.backends[name];
            if (savedBackend && typeof savedBackend === 'object') {
                out.providers[name] = normalizeProviderEntry(defaults.providers[name], savedBackend);
            }
        }
    }

    if (typeof parsed.trackChangesEnabled === 'boolean') {
        out.trackChangesEnabled = parsed.trackChangesEnabled;
    }
    if (typeof parsed.lineDiffEnabled === 'boolean') {
        out.lineDiffEnabled = parsed.lineDiffEnabled;
    }

    if (parsed.docExtraction && typeof parsed.docExtraction === 'object') {
        const richness = parsed.docExtraction.richness;
        out.docExtraction = {
            richness: typeof richness === 'string' && KNOWN_RICHNESS.includes(richness) ? richness : 'structured',
        };
    }

    if (typeof parsed.trackedChangesExtraction === 'boolean') {
        out.trackedChangesExtraction = parsed.trackedChangesExtraction;
    }
    if (Number.isFinite(parsed.commentGranularity)) {
        out.commentGranularity = parsed.commentGranularity;
    }
    if (typeof parsed.includeCommentsInSelection === 'boolean') {
        out.includeCommentsInSelection = parsed.includeCommentsInSelection;
    }

    // MCP tool servers (lib/mcp-client.js): sanitize field-by-field so a
    // corrupt entry can never break the add-in.
    if (Array.isArray(parsed.mcpServers)) {
        out.mcpServers = parsed.mcpServers
            .filter((s) => s && typeof s === 'object')
            .map((s) => ({
                name: typeof s.name === 'string' ? s.name : '',
                url: typeof s.url === 'string' ? s.url : '',
                token: typeof s.token === 'string' ? s.token : '',
                enabled: s.enabled !== false,
            }))
            .slice(0, 10);
    }
    if (Number.isFinite(parsed.mcpStepBudget) && parsed.mcpStepBudget > 0) {
        out.mcpStepBudget = Math.min(Math.round(parsed.mcpStepBudget), 48);
    }

    return out;
}

/**
 * Loads settings from localStorage ('wordAI.config') into state.config,
 * migrating legacy formats. Keeps defaults and logs a warning on corrupt data.
 *
 * @param {object} [state] - App state (defaults to the shared appState)
 * @param {function} [log] - Logging callback (message, type)
 */
export function loadSettings(state = appState, log = console.warn) {
    try {
        const saved = localStorage.getItem('wordAI.config');
        if (!saved) return;

        const parsed = JSON.parse(saved);

        if (parsed.ollamaUrl && !parsed.backends) {
            // Old flat format detected -- migrate to nested backends structure
            state.config = normalizeConfig(state.config, {
                backend: 'ollama',
                backends: {
                    ollama: {
                        url: parsed.ollamaUrl,
                        apiKey: parsed.apiKey || '',
                        model: parsed.selectedModel,
                    },
                },
                trackChangesEnabled: parsed.trackChangesEnabled,
                lineDiffEnabled: parsed.lineDiffEnabled,
            });
            // Save migrated config immediately so migration only runs once
            localStorage.setItem('wordAI.config', JSON.stringify(state.config));
        } else {
            // Nested format -- validate field by field against defaults
            state.config = normalizeConfig(state.config, parsed);
        }
    } catch (e) {
        // Corrupt JSON in localStorage -- keep defaults and tell the user
        // instead of silently discarding their saved settings.
        console.error("Failed to load settings:", e);
        log("Saved settings could not be read (corrupt data). Defaults restored.", "warning");
    }
}

/**
 * Persists state.config to localStorage ('wordAI.config').
 *
 * @param {object} [state] - App state (defaults to the shared appState)
 * @throws {Error} If localStorage write fails
 */
export function persistSettings(state = appState) {
    localStorage.setItem('wordAI.config', JSON.stringify(state.config));
}
