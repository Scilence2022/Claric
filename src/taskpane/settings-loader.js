import { loadLazyModule } from '../lib/lazy-module-loader.js';

export function createSettingsLoader({
    onConfigChanged,
    documentWriting,
    log,
    load = () => loadLazyModule('settings', () => import(/* webpackChunkName: "settings" */ './ui/settings-view.js')),
}) {
    let pending = null;
    const ready = () => {
        if (!pending) {
            pending = Promise.resolve().then(load).then((view) => {
                view.initSettings({ onConfigChanged, bindOpenButton: false, ...(documentWriting ? { documentWriting } : {}) });
                return view;
            }).catch((error) => {
                pending = null;
                throw error;
            });
        }
        return pending;
    };
    let startupProbe = null;
    const testConnection = () => {
        if (!startupProbe) {
            startupProbe = ready().then((view) => view.testConnectionUI()).catch((error) => {
                startupProbe = null;
                log(`Settings connection initialization failed: ${error.message}`, 'error');
            });
        }
        return startupProbe;
    };
    return {
        testConnection,
        async open(section) {
            try {
                const view = await ready();
                if (typeof section === 'string') view.openSettings(section);
                else view.openSettings();
                void testConnection();
            } catch (error) {
                log(`Open settings failed: ${error.message}`, 'error');
            }
        },
        /**
         * Commits a model and/or thinking level picked in the composer's
         * model popover ({model, thinkingLevel}). Resolves the settings
         * chunk first — it is loaded at startup anyway, and the commit
         * shares the saveSettings path so the form stays consistent.
         */
        async applyChatSelection(changes) {
            try {
                const view = await ready();
                view.applyChatSelection(changes);
            } catch (error) {
                log(`Could not apply model selection: ${error.message}`, 'error');
            }
        },
    };
}
