/** Retry downloads, never the document actions exported by a loaded module. */
export function isModuleLoadFailure(error) {
    return error?.name === 'ChunkLoadError'
        || /Loading chunk [\w-]+ failed|Failed to fetch dynamically imported module|Importing a module script failed|error loading dynamically imported module/i.test(error?.message || '');
}

export class LazyModuleLoadError extends Error {
    constructor(label, cause) {
        super(`Could not load ${label}. Check your connection and try again. If Claric was updated while this pane was open, close and reopen the pane after reviewing any available proposals.`, { cause });
        this.name = 'LazyModuleLoadError';
        this.code = 'lazy_module_unavailable';
        this.request = cause?.request;
    }
}

function awaitForCaller(pending, signal) {
    if (!signal) return pending;
    if (signal.aborted) return Promise.reject(new DOMException('Module loading cancelled.', 'AbortError'));
    return new Promise((resolve, reject) => {
        const abort = () => reject(new DOMException('Module loading cancelled.', 'AbortError'));
        signal.addEventListener('abort', abort, { once: true });
        pending.then(resolve, reject).finally(() => signal.removeEventListener('abort', abort));
    });
}

/**
 * Each caller may cancel its wait without aborting another caller's import.
 * Failed downloads are evicted so a later turn can make a fresh attempt.
 * @param {{retryDelays?: number[], delay?: (ms: number) => Promise<any>}} [options]
 */
export function createLazyModuleLoader({
    retryDelays = [250, 1000],
    delay = (ms) => new Promise((resolve) => setTimeout(resolve, ms)),
} = {}) {
    if (!Array.isArray(retryDelays) || retryDelays.length > 5
        || retryDelays.some((ms) => !Number.isFinite(ms) || ms < 0 || ms > 60000)) {
        throw new Error('Module retry delays must contain at most five bounded delays.');
    }
    const delays = [...retryDelays];
    const modules = new Map();
    /**
     * @param {string} id Stable module key (the import path or webpack chunk name).
     * @param {() => Promise<any>} load An import callback, without document actions.
     * @param {{signal?: AbortSignal, label?: string}} [options]
     * @returns {Promise<any>}
     */
    return function loadLazyModule(id, load, { signal, label = id } = {}) {
        if (signal?.aborted) return Promise.reject(new DOMException('Module loading cancelled.', 'AbortError'));
        if (!modules.has(id)) {
            const pending = (async () => {
                for (let attempt = 0; ; attempt++) {
                    try { return await load(); }
                    catch (error) {
                        if (!isModuleLoadFailure(error)) throw error;
                        if (attempt === delays.length) throw new LazyModuleLoadError(label, error);
                        await delay(delays[attempt]);
                    }
                }
            })().catch((error) => {
                modules.delete(id);
                throw error;
            });
            modules.set(id, pending);
            // All callers may cancel before a download settles. Keep the
            // shared failure observed even when no caller is still waiting.
            void pending.catch(() => {});
        }
        return awaitForCaller(modules.get(id), signal);
    };
}

export const loadLazyModule = createLazyModuleLoader();
