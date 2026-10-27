/** @jest-environment jsdom */
import { initAboutView } from '../src/taskpane/ui/about-view.js';
import appMetadata from '../package.json';

const runtime = 'a'.repeat(20);
const deployed = { hash: 'b'.repeat(12), runtimeHash: runtime, appVersion: appMetadata.version,
    builtAt: '2026-10-27T14:00:00.000Z' };
const settle = async () => { for (let i = 0; i < 15; i++) await Promise.resolve(); };

beforeEach(() => {
    document.body.innerHTML = '<button id="infoBtn"></button><div id="aboutOverlay"><div id="aboutPanel">'
        + '<button id="aboutCloseBtn"></button><span id="aboutVersion"></span><span id="aboutBuild">Loading…</span>'
        + '<span id="aboutBuiltAt">Loading…</span><span id="aboutPlatform"></span><span id="aboutWordApi"></span></div></div>';
    global.__webpack_hash__ = runtime;
    global.fetch = jest.fn(async () => ({ ok: true, json: async () => deployed }));
});
afterEach(() => { delete global.__webpack_hash__; delete global.fetch; });

test('reports the loaded runtime before metadata arrives and confirms matching deployment metadata', async () => {
    const log = jest.fn();
    initAboutView({ platform: 'Mac', wordApiVersion: '1.8', log });
    expect(log).toHaveBeenCalledWith(expect.stringContaining(`loaded runtime: ${runtime}`), 'info');
    await settle();
    expect(log).toHaveBeenCalledWith(expect.stringContaining(`Claric build: ${deployed.hash}`), 'info');
    expect(document.getElementById('aboutBuild').textContent).toBe(deployed.hash);
});

test('a cached runtime is not mislabeled with the new deployment build or version', async () => {
    global.fetch.mockResolvedValue({ ok: true, json: async () => ({ ...deployed, runtimeHash: 'c'.repeat(20), appVersion: '9.0.0' }) });
    const log = jest.fn();
    initAboutView({ platform: 'Mac', wordApiVersion: '1.8', log });
    await settle();
    expect(log).toHaveBeenCalledWith(expect.stringMatching(/different runtime.*Reload/), 'warning');
    expect(log).not.toHaveBeenCalledWith(expect.stringContaining('Claric build:'), expect.anything());
    expect(document.getElementById('aboutBuild').textContent).toBe(runtime);
    expect(document.getElementById('aboutVersion').textContent).toBe(appMetadata.version);
    expect(document.getElementById('aboutBuiltAt').textContent).toBe('Unavailable');
});

test.each(['offline', 'legacy metadata', 'development'])('keeps actual runtime evidence when %s prevents confirmation', async (mode) => {
    if (mode === 'offline') global.fetch.mockRejectedValue(new Error('Offline'));
    if (mode === 'legacy metadata') global.fetch.mockResolvedValue({ ok: true, json: async () => ({ ...deployed, runtimeHash: undefined }) });
    if (mode === 'development') delete global.__webpack_hash__;
    const log = jest.fn();
    initAboutView({ platform: 'Mac', wordApiVersion: '1.8', log });
    await settle();
    expect(log).not.toHaveBeenCalledWith(expect.stringContaining('Claric build:'), expect.anything());
    expect(document.getElementById('aboutBuild').textContent).toBe(mode === 'development' ? 'Unavailable' : runtime);
    expect(document.getElementById('aboutBuiltAt').textContent).toBe('Unavailable');
});
