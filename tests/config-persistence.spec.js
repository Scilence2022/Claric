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
const { KNOWN_IMAGE_PROVIDERS, defaultImageProviderConfig, DEFAULT_IMAGE_SIZE } = require('../src/lib/image-providers.js');

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

function makeImageDefaults() {
    return {
        ...makeDefaults(),
        imageGeneration: {
            enabled: false,
            provider: 'openai',
            providers: defaultImageProviderConfig(''),
        },
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
        // vllm url and missing generation settings fall back to defaults
        expect(out.providers.vllm.url).toBe('http://x:8026');
        expect(out.providers.vllm.model).toBe(PROVIDER_PRESETS.vllm.model);
        expect(out.providers.vllm.apiPath).toBe('/v1');
        expect(out.providers.vllm.thinkingLevel).toBe('default');
        expect(out.providers.vllm.temperature).toBe(1);
        // ollama untouched
        expect(out.providers.ollama).toEqual(makeDefaults().providers.ollama);
    });

    test('accepts valid thinking level and temperature values', () => {
        const out = normalizeConfig(makeDefaults(), {
            providers: { ollama: { thinkingLevel: 'high', temperature: 0.7 } },
        });
        expect(out.providers.ollama.thinkingLevel).toBe('high');
        expect(out.providers.ollama.temperature).toBe(0.7);
    });

    test('accepts the extended canonical thinking values persisted per provider', () => {
        // Model-specific profiles use values beyond low/medium/high (off, on,
        // adaptive, always, minimal, xhigh, max); a saved value must survive
        // normalizeConfig so the active provider keeps its own dial setting.
        const out = normalizeConfig(makeDefaults(), {
            providers: {
                deepseek: { thinkingLevel: 'max' },
                glm: { thinkingLevel: 'xhigh' },
                kimi: { thinkingLevel: 'on' },
                minimax: { thinkingLevel: 'adaptive' },
                'minimax-cn': { thinkingLevel: 'off' },
                zhongkeyu: { thinkingLevel: 'minimal' },
            },
        });
        expect(out.providers.deepseek.thinkingLevel).toBe('max');
        expect(out.providers.glm.thinkingLevel).toBe('xhigh');
        expect(out.providers.kimi.thinkingLevel).toBe('on');
        expect(out.providers.minimax.thinkingLevel).toBe('adaptive');
        expect(out.providers['minimax-cn'].thinkingLevel).toBe('off');
        expect(out.providers.zhongkeyu.thinkingLevel).toBe('minimal');
    });

    test('replaces invalid thinking level and temperature values with defaults', () => {
        const out = normalizeConfig(makeDefaults(), {
            providers: {
                ollama: { thinkingLevel: 'ultra', temperature: 3 },
                vllm: { thinkingLevel: 'low', temperature: Number.NaN },
            },
        });
        expect(out.providers.ollama.thinkingLevel).toBe('default');
        expect(out.providers.ollama.temperature).toBe(1);
        expect(out.providers.vllm.thinkingLevel).toBe('low');
        expect(out.providers.vllm.temperature).toBe(1);
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
        // local-served test context → origin-adaptive default is the proxy path
        expect(out.providers.deepseek.url).toBe('/deepseek');
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
        // local-served test context → origin-adaptive default is the proxy path
        expect(out.providers.deepseek.url).toBe('/deepseek');
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

    test('a saved provider entry does not leak into the defaults', () => {
        // A bare spread aliases `providers`, so the per-provider writes landed
        // in the caller's object: reusing one defaults object across two calls
        // gave the second call the first call's saved URL as its fallback.
        const defaults = makeDefaults();
        const first = normalizeConfig(defaults, {
            providers: { ollama: { url: 'http://gpu-box:11434', model: 'llama3' } },
        });
        expect(first.providers.ollama.url).toBe('http://gpu-box:11434');
        expect(defaults.providers.ollama.url).toBe('/ollama');

        // Second call over the same defaults sees pristine fallbacks.
        const second = normalizeConfig(defaults, { providers: { ollama: { model: 'qwen3' } } });
        expect(second.providers.ollama.url).toBe('/ollama');
        expect(second.providers.ollama.model).toBe('qwen3');
    });

    test('the legacy backends shape does not mutate the defaults either', () => {
        const defaults = makeDefaults();
        normalizeConfig(defaults, { backends: { vllm: { url: 'http://x:8026' } } });
        expect(defaults.providers.vllm.url).toBe(PROVIDER_PRESETS.vllm.url);
    });
});

describe('imageGeneration normalization', () => {
    test('preserves valid provider settings and fills missing entries', () => {
        const out = normalizeConfig(makeImageDefaults(), {
            imageGeneration: {
                enabled: true,
                provider: 'minimax',
                providers: {
                    minimax: {
                        url: '/custom-minimax',
                        apiKey: 'image-key',
                        model: 'image-01',
                        size: '1536x1024',
                    },
                },
            },
        });

        expect(out.imageGeneration.enabled).toBe(true);
        expect(out.imageGeneration.provider).toBe('minimax');
        expect(out.imageGeneration.providers.minimax).toEqual({
            url: '/custom-minimax',
            apiKey: 'image-key',
            model: 'image-01',
            apiPath: '/v1',
            size: '1536x1024',
        });
        expect(out.imageGeneration.providers.openai.model).toBe('gpt-image-1');
    });

    test('keeps explicitly cleared URL and model fields empty', () => {
        const out = normalizeConfig(makeImageDefaults(), {
            imageGeneration: {
                providers: {
                    custom: { url: '', model: '', apiKey: '' },
                },
            },
        });

        expect(out.imageGeneration.providers.custom.url).toBe('');
        expect(out.imageGeneration.providers.custom.model).toBe('');
    });

    test('falls back for invalid image provider, size, and field types', () => {
        const defaults = makeImageDefaults();
        const out = normalizeConfig(defaults, {
            imageGeneration: {
                provider: 'unknown',
                providers: {
                    openai: { url: 42, model: null, apiKey: 7, size: 'huge' },
                },
            },
        });

        expect(out.imageGeneration.provider).toBe('openai');
        expect(out.imageGeneration.providers.openai.url).toBe('/openai');
        expect(out.imageGeneration.providers.openai.model).toBe('gpt-image-1');
        expect(out.imageGeneration.providers.openai.apiKey).toBe('');
        expect(out.imageGeneration.providers.openai.size).toBe(DEFAULT_IMAGE_SIZE);
        expect(Object.keys(out.imageGeneration.providers)).toEqual(KNOWN_IMAGE_PROVIDERS);
    });

    test('does not invent an image section when defaults do not have one', () => {
        const out = normalizeConfig(makeDefaults(), {
            imageGeneration: { enabled: true },
        });
        expect(out.imageGeneration).toBeUndefined();
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
