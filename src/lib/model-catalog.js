/**
 * Model catalog cache shared by the Settings combobox and the chat model
 * popover.
 *
 * Both surfaces show the ids a provider's `{url}{apiPath}/models` endpoint
 * returned. The cache lives here (not in the lazy settings chunk) so the
 * composer popover can read it without loading Settings. Entries are scoped
 * per surface ('chat' | 'image') and keyed by provider + active profile;
 * a saved list is only returned while the endpoint identity (url, apiPath,
 * key) is unchanged, so editing credentials immediately hides the old catalog.
 *
 * @module model-catalog
 */

/** Per-surface caches: scope -> (provider -> {identity, models}). */
const catalogs = new Map();

function scopeCache(scope) {
    if (!catalogs.has(scope)) catalogs.set(scope, new Map());
    return catalogs.get(scope);
}

function modelCacheKey(provider, entry) {
    return JSON.stringify([provider, entry.activeProfileId || 'default']);
}

/**
 * Identity of the connection a cached list was fetched with. Key or endpoint
 * edits change it and invalidate the saved list.
 *
 * @param {string} provider - Provider id (backend)
 * @param {object} entry - Provider entry from config (url, apiPath, apiKey, profiles)
 * @returns {string} Opaque identity string
 */
export function modelConnectionIdentity(provider, entry) {
    return JSON.stringify([modelCacheKey(provider, entry), entry.url, entry.apiPath, entry.apiKey]);
}

/**
 * Returns the cached model ids for a provider entry, or [] when nothing was
 * cached or the connection identity changed since.
 *
 * @param {string} scope - Cache scope ('chat' | 'image')
 * @param {string} provider - Provider id (backend)
 * @param {object} entry - Provider entry from config
 * @returns {string[]} Cached model ids
 */
export function cachedModels(scope, provider, entry) {
    const saved = scopeCache(scope).get(modelCacheKey(provider, entry));
    return saved?.identity === modelConnectionIdentity(provider, entry) ? saved.models : [];
}

/**
 * Stores model ids for a provider entry under the current connection identity.
 *
 * @param {string} scope - Cache scope ('chat' | 'image')
 * @param {string} provider - Provider id (backend)
 * @param {object} entry - Provider entry from config
 * @param {string[]} models - Model ids from the provider's models endpoint
 */
export function cacheModels(scope, provider, entry, models) {
    scopeCache(scope).set(modelCacheKey(provider, entry), {
        identity: modelConnectionIdentity(provider, entry),
        models: [...models],
    });
}

/**
 * Drops the cached list for a provider entry (e.g. its profile was removed).
 *
 * @param {string} scope - Cache scope ('chat' | 'image')
 * @param {string} provider - Provider id (backend)
 * @param {object} entry - Provider entry from config
 */
export function clearCachedModels(scope, provider, entry) {
    scopeCache(scope).delete(modelCacheKey(provider, entry));
}
