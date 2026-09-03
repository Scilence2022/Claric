/** @jest-environment jsdom */

/**
 * Image-provider settings keep endpoint/model/key state per provider. The
 * provider <select> changes before its change handler runs, so the visible
 * fields still belong to the old provider at that moment.
 */
import { syncImageFormToProvider } from '../src/taskpane/ui/settings-view.js';

function renderImageFields({ selected = 'glm', url = '', model = '', apiKey = '', size = '1024x1024' } = {}) {
    document.body.innerHTML = `
        <select id="imageProviderSelect">
            <option value="openai">OpenAI</option>
            <option value="glm">GLM</option>
        </select>
        <input id="imageEndpointUrl">
        <input id="imageModelInput">
        <input id="imageApiKey">
        <select id="imageSizeSelect">
            <option value="auto">auto</option>
            <option value="1024x1024">1024x1024</option>
            <option value="1536x1024">1536x1024</option>
        </select>`;
    document.getElementById('imageProviderSelect').value = selected;
    document.getElementById('imageEndpointUrl').value = url;
    document.getElementById('imageModelInput').value = model;
    document.getElementById('imageApiKey').value = apiKey;
    document.getElementById('imageSizeSelect').value = size;
}

function makeConfig() {
    return {
        imageGeneration: {
            enabled: true,
            provider: 'openai',
            providers: {
                openai: {
                    url: '/openai', apiKey: 'openai-old', model: 'gpt-image-1',
                    apiPath: '/v1', size: '1024x1024',
                },
                glm: {
                    url: '/glm', apiKey: 'glm-key', model: 'cogview-4',
                    apiPath: '/api/paas/v4', size: '1024x1024',
                },
            },
        },
    };
}

describe('syncImageFormToProvider', () => {
    afterEach(() => {
        document.body.innerHTML = '';
    });

    test('commits the old form after the provider select has already changed', () => {
        const config = makeConfig();
        renderImageFields({
            selected: 'glm',
            url: '  https://relay.example/openai  ',
            model: '  gpt-image-custom  ',
            apiKey: '  openai-new-key  ',
            size: '1536x1024',
        });

        expect(syncImageFormToProvider('openai', config)).toBe(true);
        expect(config.imageGeneration.providers.openai).toMatchObject({
            url: 'https://relay.example/openai',
            model: 'gpt-image-custom',
            apiKey: 'openai-new-key',
            size: '1536x1024',
        });
        expect(config.imageGeneration.providers.glm).toMatchObject({
            url: '/glm', model: 'cogview-4', apiKey: 'glm-key',
        });
    });

    test('persists deliberately cleared endpoint, model, and key fields', () => {
        const config = makeConfig();
        renderImageFields({ selected: 'glm', url: '  ', model: '', apiKey: ' ' });

        syncImageFormToProvider('glm', config);

        expect(config.imageGeneration.providers.glm).toMatchObject({
            url: '', model: '', apiKey: '', size: '1024x1024',
        });
    });

    test('leaves config untouched when the image settings section is absent', () => {
        renderImageFields({ selected: 'openai', url: '/new' });
        const config = {};

        expect(syncImageFormToProvider('openai', config)).toBe(false);
        expect(config).toEqual({});
    });
});
