const { createSettingsLoader } = require('../src/taskpane/settings-loader.js');

function setup(loadOverride) {
    const view = {
        initSettings: jest.fn(), openSettings: jest.fn(),
        testConnectionUI: jest.fn().mockResolvedValue(undefined),
    };
    const load = loadOverride || jest.fn().mockResolvedValue(view);
    const onConfigChanged = jest.fn();
    const log = jest.fn();
    const settings = createSettingsLoader({ load, onConfigChanged, log });
    return { view, load, onConfigChanged, log, settings };
}

test('does not load until needed and initializes before the startup connection probe', async () => {
    const { settings, view, load, onConfigChanged } = setup();
    expect(load).not.toHaveBeenCalled();
    await settings.testConnection();
    expect(load).toHaveBeenCalledTimes(1);
    expect(view.initSettings).toHaveBeenCalledWith({ onConfigChanged, bindOpenButton: false });
    expect(view.initSettings.mock.invocationCallOrder[0]).toBeLessThan(view.testConnectionUI.mock.invocationCallOrder[0]);
    expect(view.openSettings).not.toHaveBeenCalled();
});

test('opening during startup shares the import and binds settings only once', async () => {
    let resolve;
    const load = jest.fn(() => new Promise((done) => { resolve = done; }));
    const { settings, view } = setup(load);
    const probe = settings.testConnection();
    const open = settings.open();
    await Promise.resolve();
    expect(view.openSettings).not.toHaveBeenCalled();
    resolve(view);
    await Promise.all([probe, open]);
    await settings.open();
    await settings.testConnection();
    expect(load).toHaveBeenCalledTimes(1);
    expect(view.initSettings).toHaveBeenCalledTimes(1);
    expect(view.openSettings).toHaveBeenCalledTimes(2);
    expect(view.testConnectionUI).toHaveBeenCalledTimes(1);
});

test('opening does not wait for the connection request to finish', async () => {
    const { settings, view } = setup();
    let finish;
    view.testConnectionUI.mockImplementation(() => new Promise((resolve) => { finish = resolve; }));
    const probe = settings.testConnection();
    await settings.open();
    expect(view.openSettings).toHaveBeenCalledTimes(1);
    finish();
    await probe;
});

test('startup chunk failure is handled and a later open retries initialization and probe', async () => {
    const { settings, view, load, log } = setup();
    load.mockRejectedValueOnce(new Error('chunk unavailable'));
    await expect(settings.testConnection()).resolves.toBeUndefined();
    expect(log).toHaveBeenCalledWith(expect.stringContaining('chunk unavailable'), 'error');
    await settings.open();
    await settings.testConnection();
    expect(load).toHaveBeenCalledTimes(2);
    expect(view.initSettings).toHaveBeenCalledTimes(1);
    expect(view.testConnectionUI).toHaveBeenCalledTimes(1);
    expect(view.openSettings).toHaveBeenCalledTimes(1);
});

test('click load failure is handled and retry succeeds', async () => {
    const { settings, view, load, log } = setup();
    load.mockRejectedValueOnce(new Error('offline'));
    await expect(settings.open()).resolves.toBeUndefined();
    expect(log).toHaveBeenCalledWith('Open settings failed: offline', 'error');
    await settings.open();
    await settings.testConnection();
    expect(view.initSettings).toHaveBeenCalledTimes(1);
    expect(view.openSettings).toHaveBeenCalledTimes(1);
});

test('a rejected connection probe does not block opening settings', async () => {
    const { settings, view, log } = setup();
    view.testConnectionUI.mockRejectedValueOnce(new Error('connection failed'));
    await settings.testConnection();
    await settings.open();
    await settings.testConnection();
    expect(log).toHaveBeenCalledWith(expect.stringContaining('connection failed'), 'error');
    expect(view.initSettings).toHaveBeenCalledTimes(1);
    expect(view.openSettings).toHaveBeenCalledTimes(1);
});
