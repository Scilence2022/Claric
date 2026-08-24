/**
 * Tests for normalizeConfig (taskpane.js settings validation).
 *
 * A corrupt or hand-edited `wordAI.config` in localStorage previously went
 * through a shallow spread merge, letting partial objects drop whole config
 * sections and crash later reads (config.backends[backend].url). These tests
 * pin the field-by-field validation behavior.
 */

const { normalizeConfig } = require('../src/taskpane/taskpane.js');

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
        backends: {
            ollama: { url: '/ollama', apiKey: '', model: 'gpt-oss:20b' },
            vllm: { url: '/vllm', apiKey: '', model: 'qwen3.5-35b-a3b' },
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
        expect(out.backends.ollama.url).toBe('http://gpu-box:11434');
        expect(out.backends.vllm.model).toBe('qwen3.5');
        expect(out.docExtraction.richness).toBe('plain');
        expect(out.commentGranularity).toBe(2);
    });

    test('a partial backends object keeps the other backend intact', () => {
        const out = normalizeConfig(makeDefaults(), {
            backends: { vllm: { url: 'http://x:8026' } },
        });
        // vllm url applied, missing model falls back to default
        expect(out.backends.vllm.url).toBe('http://x:8026');
        expect(out.backends.vllm.model).toBe('qwen3.5-35b-a3b');
        // ollama untouched
        expect(out.backends.ollama).toEqual(makeDefaults().backends.ollama);
    });

    test('an unknown backend name falls back to the default backend', () => {
        const out = normalizeConfig(makeDefaults(), { backend: 'openai' });
        expect(out.backend).toBe('ollama');
        expect(Object.keys(out.backends).sort()).toEqual(['ollama', 'vllm']);
    });

    test('non-string url/apiKey/model are replaced with safe values', () => {
        const out = normalizeConfig(makeDefaults(), {
            backends: { ollama: { url: 42, apiKey: null, model: { nested: true } } },
        });
        expect(out.backends.ollama.url).toBe('/ollama');
        expect(out.backends.ollama.apiKey).toBe('');
        expect(out.backends.ollama.model).toBe('gpt-oss:20b');
    });

    test('backends of the wrong type is ignored entirely', () => {
        const out = normalizeConfig(makeDefaults(), { backends: 'https://evil' });
        expect(out.backends).toEqual(makeDefaults().backends);
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
    });
});
