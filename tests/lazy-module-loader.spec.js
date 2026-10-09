const { createLazyModuleLoader, isModuleLoadFailure, LazyModuleLoadError } = require('../src/lib/lazy-module-loader.js');

function chunkError() {
    return Object.assign(new Error('Loading chunk 904 failed.'), {
        name: 'ChunkLoadError', request: 'https://example.test/agent-actions.12345678.js',
    });
}

test('coalesces callers and reuses the module without rerunning its import', async () => {
    const loader = createLazyModuleLoader();
    let finish;
    const load = jest.fn(() => new Promise((resolve) => { finish = resolve; }));
    const first = loader('tools', load);
    const second = loader('tools', load);
    expect(load).toHaveBeenCalledTimes(1);
    const module = { apply: jest.fn() };
    finish(module);
    expect(await first).toBe(module);
    expect(await second).toBe(module);
    expect(await loader('tools', load)).toBe(module);
    expect(load).toHaveBeenCalledTimes(1);
    expect(module.apply).not.toHaveBeenCalled();
});

test('retries transient chunk fetches only and never executes exported document actions', async () => {
    const delay = jest.fn().mockResolvedValue(undefined);
    const loader = createLazyModuleLoader({ retryDelays: [1, 2], delay });
    const module = { write: jest.fn() };
    const load = jest.fn().mockRejectedValueOnce(chunkError()).mockRejectedValueOnce(chunkError()).mockResolvedValue(module);
    expect(await loader('tools', load)).toBe(module);
    expect(delay.mock.calls).toEqual([[1], [2]]);
    expect(load).toHaveBeenCalledTimes(3);
    expect(module.write).not.toHaveBeenCalled();
});

test('stops after bounded retries, retains diagnostics, and permits a later fresh attempt', async () => {
    const loader = createLazyModuleLoader({ retryDelays: [0, 0], delay: async () => {} });
    const cause = chunkError();
    const load = jest.fn().mockRejectedValue(cause);
    const failure = await loader('tools', load, { label: 'table tools' }).catch((error) => error);
    expect(failure).toBeInstanceOf(LazyModuleLoadError);
    expect(failure).toMatchObject({ cause, request: cause.request, code: 'lazy_module_unavailable' });
    expect(failure.message).toContain('table tools');
    expect(failure.message).toContain('reviewing any available proposals');
    expect(load).toHaveBeenCalledTimes(3);
    const module = {};
    load.mockResolvedValue(module);
    expect(await loader('tools', load)).toBe(module);
    expect(load).toHaveBeenCalledTimes(4);
});

test.each([
    new Error('GeneralException'), new TypeError('Cannot read properties of undefined'),
    new DOMException('Cancelled', 'AbortError'), new SyntaxError('Unexpected token'),
])('does not retry an evaluation, action, or cancellation error: %s', async (failure) => {
    const delay = jest.fn();
    const loader = createLazyModuleLoader({ delay });
    const load = jest.fn().mockRejectedValue(failure);
    await expect(loader('tools', load)).rejects.toBe(failure);
    expect(load).toHaveBeenCalledTimes(1);
    expect(delay).not.toHaveBeenCalled();
});

test.each([
    'Failed to fetch dynamically imported module: https://example.test/module.js',
    'Importing a module script failed.',
    'error loading dynamically imported module',
    'Loading chunk settings failed.',
])('recognizes browser module download failures: %s', (message) => {
    expect(isModuleLoadFailure(new TypeError(message))).toBe(true);
});

test('handles absent errors without treating them as download failures', () => {
    expect(isModuleLoadFailure(undefined)).toBe(false);
});

test('cancelling one caller preserves the shared module download for another', async () => {
    const loader = createLazyModuleLoader();
    const controller = new AbortController();
    let finish;
    const load = jest.fn(() => new Promise((resolve) => { finish = resolve; }));
    const cancelled = loader('tools', load, { signal: controller.signal });
    const other = loader('tools', load);
    controller.abort();
    await expect(cancelled).rejects.toMatchObject({ name: 'AbortError' });
    const module = {};
    finish(module);
    expect(await other).toBe(module);
    expect(await loader('tools', load)).toBe(module);
    expect(load).toHaveBeenCalledTimes(1);
});

test('an already cancelled caller does not start a download', async () => {
    const controller = new AbortController();
    controller.abort();
    const load = jest.fn();
    await expect(createLazyModuleLoader()('tools', load, { signal: controller.signal })).rejects.toMatchObject({ name: 'AbortError' });
    expect(load).not.toHaveBeenCalled();
});

test('removes the caller abort listener after a download fails', async () => {
    const controller = new AbortController();
    const remove = jest.spyOn(controller.signal, 'removeEventListener');
    const load = jest.fn().mockRejectedValue(chunkError());
    await expect(createLazyModuleLoader({ retryDelays: [] })('tools', load, { signal: controller.signal })).rejects.toBeInstanceOf(LazyModuleLoadError);
    await Promise.resolve();
    expect(remove).toHaveBeenCalledWith('abort', expect.any(Function));
});

test('an abandoned failed download is observed and a later caller can recover', async () => {
    const loader = createLazyModuleLoader({ retryDelays: [] });
    const controller = new AbortController();
    let fail;
    const cancelled = loader('tools', () => new Promise((_resolve, reject) => { fail = reject; }), { signal: controller.signal });
    controller.abort();
    await expect(cancelled).rejects.toMatchObject({ name: 'AbortError' });
    fail(chunkError());
    await Promise.resolve();
    await Promise.resolve();
    const module = {};
    expect(await loader('tools', async () => module)).toBe(module);
});

test('uses bounded timers by default', async () => {
    jest.useFakeTimers();
    try {
        const load = jest.fn().mockRejectedValueOnce(chunkError()).mockResolvedValue({});
        const pending = createLazyModuleLoader()('tools', load);
        await jest.advanceTimersByTimeAsync(249);
        expect(load).toHaveBeenCalledTimes(1);
        await jest.advanceTimersByTimeAsync(1);
        await pending;
        expect(load).toHaveBeenCalledTimes(2);
    } finally { jest.useRealTimers(); }
});

test.each([[-1], [Infinity], [60001], [0, 0, 0, 0, 0, 0], null])('rejects unbounded retry configuration: %j', (retryDelays) => {
    expect(() => createLazyModuleLoader({ retryDelays })).toThrow('bounded delays');
});
