const {
    addProviderProfile, selectProviderProfile, removeProviderProfile,
    syncActiveProviderProfile,
} = require('../src/lib/provider-profile-actions.js');
const {
    defaultConfig, normalizeConfig, getActiveBackendConfig, getActiveImageConfig,
} = require('../src/taskpane/app-state.js');

test('migrates single-key settings and round-trips distinct keys, models and parameters', () => {
    const config = normalizeConfig(defaultConfig(), {
        backend: 'claude', providers: { claude: { apiKey: 'personal', model: 'claude-a', temperature: 0.3 } },
    });
    const entry = config.providers.claude;
    addProviderProfile(entry);
    expect(entry.apiKey).toBe('');
    entry.apiKey = 'team';
    entry.model = 'claude-b';
    entry.url = '/team-relay';
    entry.thinkingLevel = 'high';
    syncActiveProviderProfile(entry);
    const restored = normalizeConfig(defaultConfig(), JSON.parse(JSON.stringify(config)));
    expect(getActiveBackendConfig({ config: restored })).toMatchObject({
        provider: 'claude', apiKey: 'team', model: 'claude-b', url: '/team-relay', thinkingLevel: 'high',
    });
    expect(getActiveBackendConfig({ config: restored })).not.toHaveProperty('profiles');
    selectProviderProfile(restored.providers.claude, 'default');
    expect(getActiveBackendConfig({ config: restored })).toMatchObject({
        apiKey: 'personal', model: 'claude-a', temperature: 0.3,
    });
    expect(entry.apiKey).toBe('team'); // restored profiles do not alias the input
});

test('image requests only receive the active profile and restore per-key sizes', () => {
    const config = defaultConfig();
    config.imageGeneration.enabled = true;
    const entry = config.imageGeneration.providers.openai;
    entry.apiKey = 'first';
    addProviderProfile(entry);
    entry.apiKey = 'second';
    entry.model = 'gpt-image-team';
    entry.size = '1536x1024';
    syncActiveProviderProfile(entry);
    const restored = normalizeConfig(defaultConfig(), JSON.parse(JSON.stringify(config)));
    expect(getActiveImageConfig({ config: restored })).toMatchObject({ apiKey: 'second', size: '1536x1024' });
    expect(JSON.stringify(getActiveImageConfig({ config: restored }))).not.toContain('first');
    selectProviderProfile(restored.imageGeneration.providers.openai, 'default');
    expect(getActiveImageConfig({ config: restored })).toMatchObject({ apiKey: 'first', size: '1024x1024' });
});

test('invalid selection is harmless and removal keeps a usable profile without deleted credentials', () => {
    const entry = defaultConfig().providers.deepseek;
    entry.apiKey = 'keep';
    addProviderProfile(entry);
    entry.apiKey = 'delete';
    syncActiveProviderProfile(entry);
    expect(selectProviderProfile(entry, 'missing')).toBe(false);
    expect(entry.apiKey).toBe('delete');
    expect(removeProviderProfile(entry)).toBe(true);
    expect(entry.apiKey).toBe('keep');
    expect(JSON.stringify(entry)).not.toContain('delete');
    expect(removeProviderProfile(entry)).toBe(false);
});

test('normalizes partial and corrupt profiles and falls back from a missing active id', () => {
    const profiles = [
        null, {}, { id: 'array', config: [] }, { id: '', config: {} },
        { id: 'a', name: ' Personal ', config: { apiKey: 'a', temperature: 99, thinkingLevel: 'bad' } },
        { id: 'a', config: { apiKey: 'duplicate' } },
        { id: 'b', name: 4, config: { apiKey: 17, model: 'b' } },
    ];
    const config = normalizeConfig(defaultConfig(), {
        providers: { deepseek: { apiKey: 'stale', activeProfileId: 'missing', profiles } },
        imageGeneration: { providers: { openai: {
            activeProfileId: 'image', profiles: [{ id: 'image', config: { url: '', model: '', size: 'bad' } }],
        } } },
    });
    const entry = config.providers.deepseek;
    expect(entry.activeProfileId).toBe('a');
    expect(entry.profiles).toHaveLength(2);
    expect(entry.apiKey).toBe('a');
    expect(entry.profiles[0]).toMatchObject({ name: 'Personal', config: { temperature: 1, thinkingLevel: 'default' } });
    expect(entry.profiles[1]).toMatchObject({ name: 'API key', config: { apiKey: '' } });
    expect(config.imageGeneration.providers.openai).toMatchObject({ url: '', model: '', size: '1024x1024' });
    expect(normalizeConfig(defaultConfig(), { providers: { deepseek: { apiKey: 'legacy', profiles: [null] } } })
        .providers.deepseek).toMatchObject({ apiKey: 'legacy' });
});
