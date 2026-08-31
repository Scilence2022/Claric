/**
 * Tests for normalizeConfig (app-state.js settings validation).
 *
 * A corrupt or hand-edited `wordAI.config` in localStorage previously went
 * through a shallow spread merge, letting partial objects drop whole config
 * sections and crash later reads (config.providers[backend].url). These
 * tests pin the field-by-field validation behavior, including migration
 * from the pre-0.4.0 `backends` shape.
 */

const { normalizeConfig } = require('../src/taskpane/app-state.js');
const { PROVIDER_PRESETS, KNOWN_PROVIDERS, defaultProviderConfig } = require('../src/lib/providers.js');

/** Baseline defaults matching the module-level config literal. */
function makeDefaults() {
    return {
        backend: 'ollama',
        trackChangesEnabled: true,
        lineDiffEnabled: false,
        docExtraction: { richness: 'structured' },
        trackedChangesExtraction: false,
        commentGranularity: 0,
        includeCommentsInSelection: false,
        providers: defaultProviderConfig(),
    };
}

describe('normalizeConfig', () => {
    test('returns defaults unchanged for an empty parsed object', () => {
        const defaults = makeDefaults();
        expect(normalizeConfig(defaults, {})).toEqual(defaults);
    });

    test('accepts a fully valid saved config', () => {
        const parsed = {
            backend: 'vllm',
            trackChangesEnabled: false,
            lineDiffEnabled: true,
            docExtraction: { richness: 'plain' },
            trackedChangesExtraction: true,
            commentGranularity: 2,
            includeCommentsInSelection: true,
            backends: {
                ollama: { url: 'http://gpu-box:11434', apiKey: 'secret', model: 'llama3' },
                vllm: { url: 'http://gpu-box:8026', apiKey: '', model: 'qwen3.5' },
            },
        };
        const out = normalizeConfig(makeDefaults(), parsed);
        expect(out.backend).toBe('vllm');
        expect(out.providers.ollama.url).toBe('http://gpu-box:11434');
        expect(out.providers.vllm.model).toBe('qwen3.5');
        expect(out.docExtraction.richness).toBe('plain');
        expect(out.commentGranularity).toBe(2);
    });

    test('a partial providers object keeps the other providers intact', () => {
        const out = normalizeConfig(makeDefaults(), {
            providers: { vllm: { url: 'http://x:8026' } },
        });
        // vllm url applied, missing model/apiPath fall back to defaults
        expect(out.providers.vllm.url).toBe('http://x:8026');
        expect(out.providers.vllm.model).toBe(PROVIDER_PRESETS.vllm.model);
        expect(out.providers.vllm.apiPath).toBe('/v1');
        // ollama untouched
        expect(out.providers.ollama).toEqual(makeDefaults().providers.ollama);
    });

    test('legacy backends shape migrates into providers', () => {
        const out = normalizeConfig(makeDefaults(), {
            backend: 'vllm',
            backends: {
                ollama: { url: 'http://gpu-box:11434', apiKey: 'k', model: 'llama3' },
            },
        });
        expect(out.backend).toBe('vllm');
        expect(out.providers.ollama.url).toBe('http://gpu-box:11434');
        expect(out.providers.ollama.apiKey).toBe('k');
        expect(out.providers.ollama.model).toBe('llama3');
        // migrated entry gets the preset apiPath
        expect(out.providers.ollama.apiPath).toBe('/v1');
        // cloud providers keep preset defaults
        expect(out.providers.deepseek.url).toBe(PROVIDER_PRESETS.deepseek.url);
        expect(out.providers.glm.apiPath).toBe('/api/paas/v4');
    });

    test('cloud provider entries persist and validate', () => {
        const out = normalizeConfig(makeDefaults(), {
            backend: 'glm',
            providers: {
                deepseek: { apiKey: 'sk-1', model: 'deepseek-reasoner' },
                glm: { apiKey: 'g', model: 'glm-4.6' },
                kimi: { apiKey: 'm' },
            },
        });
        expect(out.backend).toBe('glm');
        expect(out.providers.deepseek.apiKey).toBe('sk-1');
        expect(out.providers.deepseek.model).toBe('deepseek-reasoner');
        expect(out.providers.deepseek.url).toBe(PROVIDER_PRESETS.deepseek.url);
        expect(out.providers.glm.model).toBe('glm-4.6');
        expect(out.providers.glm.apiPath).toBe('/api/paas/v4');
        expect(out.providers.kimi.model).toBe(PROVIDER_PRESETS.kimi.model);
    });

    test('an unknown backend name falls back to the default backend', () => {
        const out = normalizeConfig(makeDefaults(), { backend: 'gpt4all' });
        expect(out.backend).toBe('ollama');
        expect(Object.keys(out.providers).sort()).toEqual([...KNOWN_PROVIDERS].sort());
    });

    test('non-string url/apiKey/model/apiPath are replaced with safe values', () => {
        const out = normalizeConfig(makeDefaults(), {
            providers: { ollama: { url: 42, apiKey: null, model: { nested: true }, apiPath: 7 } },
        });
        expect(out.providers.ollama.url).toBe('/ollama');
        expect(out.providers.ollama.apiKey).toBe('');
        expect(out.providers.ollama.model).toBe(PROVIDER_PRESETS.ollama.model);
        expect(out.providers.ollama.apiPath).toBe('/v1');
    });

    test('providers of the wrong type is ignored entirely', () => {
        const out = normalizeConfig(makeDefaults(), { providers: 'https://evil' });
        expect(out.providers).toEqual(makeDefaults().providers);
    });

    test('unknown docExtraction.richness falls back to structured', () => {
        const out = normalizeConfig(makeDefaults(), { docExtraction: { richness: 'deluxe' } });
        expect(out.docExtraction).toEqual({ richness: 'structured' });
    });

    test('legacy docExtraction.maxLength is dropped', () => {
        const out = normalizeConfig(makeDefaults(), {
            docExtraction: { richness: 'headings', maxLength: 5000 },
        });
        expect(out.docExtraction).toEqual({ richness: 'headings' });
    });

    test('non-boolean flags and non-numeric granularity are ignored', () => {
        const out = normalizeConfig(makeDefaults(), {
            trackChangesEnabled: 'yes',
            lineDiffEnabled: 1,
            trackedChangesExtraction: 'true',
            includeCommentsInSelection: null,
            commentGranularity: 'lots',
        });
        expect(out.trackChangesEnabled).toBe(true);
        expect(out.lineDiffEnabled).toBe(false);
        expect(out.trackedChangesExtraction).toBe(false);
        expect(out.includeCommentsInSelection).toBe(false);
        expect(out.commentGranularity).toBe(0);
    });

    test('does not mutate the defaults object', () => {
        const defaults = makeDefaults();
        normalizeConfig(defaults, { backend: 'vllm', docExtraction: { richness: 'plain' } });
        expect(defaults.backend).toBe('ollama');
        expect(defaults.docExtraction.richness).toBe('structured');
        expect(defaults.providers.ollama.url).toBe('/ollama');
    });
});

describe('autoApplyChanges normalization', () => {
    test('persists the boolean and defaults to false', () => {
        const { normalizeConfig, defaultConfig } = require('../src/taskpane/app-state.js');
        expect(defaultConfig().autoApplyChanges).toBe(false);
        const out = normalizeConfig(defaultConfig(), { autoApplyChanges: true });
        expect(out.autoApplyChanges).toBe(true);
        const corrupt = normalizeConfig(defaultConfig(), { autoApplyChanges: 'yes' });
        expect(corrupt.autoApplyChanges).toBe(false);
    });
});
