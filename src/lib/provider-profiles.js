/** Named API-key configurations within a provider. The flat entry mirrors the
 * active profile for existing consumers; inactive secrets stay out of requests. */

/** @param {Object<string, any>} entry @returns {Object<string, any>} */
export function providerRequestConfig(entry = {}) {
    const { profiles: _profiles, activeProfileId: _activeProfileId, ...config } = entry;
    return config;
}

/** Validate profiles with the same field rules as the legacy provider entry.
 * @param {Object<string, any>} saved
 * @param {Object<string, any>} flat
 * @param {function(Object<string, any>): Object<string, any>} normalize */
export function normalizeProviderProfiles(saved, flat, normalize) {
    if (!Array.isArray(saved.profiles)) return flat;
    const ids = new Set();
    const profiles = [];
    for (const profile of saved.profiles) {
        if (!profile || typeof profile.id !== 'string' || !profile.id.trim() || ids.has(profile.id)
            || !profile.config || typeof profile.config !== 'object' || Array.isArray(profile.config)) continue;
        ids.add(profile.id);
        profiles.push({
            id: profile.id,
            name: typeof profile.name === 'string' && profile.name.trim() ? profile.name.trim() : 'API key',
            config: normalize(profile.config),
        });
    }
    if (!profiles.length) return flat;
    const active = profiles.find((profile) => profile.id === saved.activeProfileId) || profiles[0];
    return { ...active.config, profiles, activeProfileId: active.id };
}
