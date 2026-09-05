export function createSettingsLoader({
    onConfigChanged,
    log,
    load = () => import(/* webpackChunkName: "settings" */ './ui/settings-view.js'),
}) {
    let pending = null;
    const ready = () => {
        if (!pending) {
            pending = Promise.resolve().then(load).then((view) => {
                view.initSettings({ onConfigChanged, bindOpenButton: false });
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
        async open() {
            try {
                const view = await ready();
                view.openSettings();
                void testConnection();
            } catch (error) {
                log(`Open settings failed: ${error.message}`, 'error');
            }
        },
    };
}
