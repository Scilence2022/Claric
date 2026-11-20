/**
 * Specs for the shared model catalog cache (lib/model-catalog.js): scope
 * isolation, connection-identity invalidation, per-profile entries, and
 * cache clearing.
 */

const {
    cachedModels,
    cacheModels,
    clearCachedModels,
    modelConnectionIdentity,
} = require('../src/lib/model-catalog.js');

const entry = (overrides = {}) => ({
    activeProfileId: 'default',
    url: 'https://api.example.com',
    apiPath: '/v1',
    apiKey: 'key-1',
    ...overrides,
});

describe('model catalog cache', () => {
    test('round-trips ids and keeps the chat and image scopes separate', () => {
        cacheModels('chat', 'ollama', entry(), ['a', 'b']);
        cacheModels('image', 'ollama', entry(), ['img-1']);
        expect(cachedModels('chat', 'ollama', entry())).toEqual(['a', 'b']);
        expect(cachedModels('image', 'ollama', entry())).toEqual(['img-1']);
    });

    test('a changed endpoint, path or key invalidates the cached list', () => {
        cacheModels('chat', 'ollama', entry(), ['a']);
        expect(cachedModels('chat', 'ollama', entry({ url: 'https://other.example.com' }))).toEqual([]);
        expect(cachedModels('chat', 'ollama', entry({ apiPath: '/api/v4' }))).toEqual([]);
        expect(cachedModels('chat', 'ollama', entry({ apiKey: 'key-2' }))).toEqual([]);
        // The original identity still resolves.
        expect(cachedModels('chat', 'ollama', entry())).toEqual(['a']);
    });

    test('profiles of one provider are cached independently', () => {
        cacheModels('chat', 'ollama', entry({ activeProfileId: 'p1' }), ['a']);
        cacheModels('chat', 'ollama', entry({ activeProfileId: 'p2' }), ['b']);
        expect(cachedModels('chat', 'ollama', entry({ activeProfileId: 'p1' }))).toEqual(['a']);
        expect(cachedModels('chat', 'ollama', entry({ activeProfileId: 'p2' }))).toEqual(['b']);
    });

    test('clearCachedModels drops only the matching entry', () => {
        cacheModels('chat', 'ollama', entry(), ['a']);
        cacheModels('chat', 'glm', entry(), ['g']);
        clearCachedModels('chat', 'ollama', entry());
        expect(cachedModels('chat', 'ollama', entry())).toEqual([]);
        expect(cachedModels('chat', 'glm', entry())).toEqual(['g']);
    });

    test('cacheModels stores a copy, and identity distinguishes connections', () => {
        const models = ['a'];
        cacheModels('chat', 'ollama', entry(), models);
        models.push('mutated');
        expect(cachedModels('chat', 'ollama', entry())).toEqual(['a']);
        expect(modelConnectionIdentity('ollama', entry()))
            .toEqual(modelConnectionIdentity('ollama', entry()));
        expect(modelConnectionIdentity('ollama', entry()))
            .not.toEqual(modelConnectionIdentity('ollama', entry({ apiKey: 'other' })));
    });

    test('unknown entries and providers read as empty without throwing', () => {
        expect(cachedModels('chat', 'never-configured', entry())).toEqual([]);
        expect(cachedModels('chat', 'ollama', entry({ activeProfileId: 'ghost' }))).toEqual([]);
    });
});
