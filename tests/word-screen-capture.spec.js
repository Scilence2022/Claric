/** @jest-environment jsdom */
const { screenCaptureState, startWordScreenSharing, stopWordScreenSharing, readWordScreen, looksLikeWordViewRequest, captureWordViewForQuestion } = require('../src/taskpane/word-screen-capture.js');
const { createWordVisualTools } = require('../src/taskpane/word-render-tools.js');
const { initScreenSharingView } = require('../src/taskpane/ui/screen-sharing-view.js');

let originalCreate;
function setup(surface = 'window') {
    const track = { readyState: 'live', getSettings: () => ({ displaySurface: surface }),
        stop: jest.fn(() => { track.readyState = 'ended'; }), addEventListener: jest.fn() };
    const stream = { getVideoTracks: () => [track], getTracks: () => [track] };
    const getDisplayMedia = jest.fn(async () => stream);
    Object.defineProperty(navigator, 'mediaDevices', { configurable: true, value: { getDisplayMedia } });
    const player = originalCreate.call(document, 'video');
    Object.defineProperties(player, { videoWidth: { value: 2400, configurable: true }, videoHeight: { value: 1400, configurable: true }, readyState: { value: 2 } });
    player.play = jest.fn(async () => {}); player.pause = jest.fn();
    player.requestVideoFrameCallback = jest.fn((callback) => { Promise.resolve().then(callback); return 1; });
    player.cancelVideoFrameCallback = jest.fn();
    let frame = 0;
    const canvas = { width: 0, height: 0, getContext: jest.fn(() => ({ drawImage: jest.fn() })),
        toDataURL: jest.fn(() => `data:image/jpeg;base64,fresh${++frame}`) };
    jest.spyOn(document, 'createElement').mockImplementation((tag, ...args) => tag === 'video' ? player
        : tag === 'canvas' ? canvas : originalCreate.call(document, tag, ...args));
    return { track, stream, player, canvas, getDisplayMedia };
}
beforeEach(() => { originalCreate = document.createElement; });
afterEach(() => { stopWordScreenSharing(); jest.restoreAllMocks(); delete navigator.mediaDevices; delete global.Office; });

test('model screenshots cannot trigger permission dialogs without user sharing', async () => {
    const w = setup();
    expect((await readWordScreen()).ok).toBe(false);
    expect(w.getDisplayMedia).not.toHaveBeenCalled();
});

test('direct shared-window frames are fresh, bounded and preferred without PDF conversion', async () => {
    const w = setup();
    await startWordScreenSharing();
    expect(w.getDisplayMedia).toHaveBeenCalledWith(expect.objectContaining({ audio: false, video: { displaySurface: 'window' } }));
    const visual = createWordVisualTools();
    expect(visual.status().preferredSource).toBe('word_shared_window');
    expect(visual.tools.map((tool) => tool.name)).toEqual(['read_word_screen']);
    const first = await visual.execute('read_word_screen', {});
    const second = await visual.execute('read_word_screen', {});
    expect(first.result).toMatchObject({ source: 'word_shared_window', width: 2000, height: 1167, visualInputAvailable: true });
    expect(first.attachments).not.toEqual(second.attachments);
    expect(w.player.requestVideoFrameCallback).toHaveBeenCalledTimes(2);
    expect(visual.status().inspectedScreen).toBe(true);
    expect(w.canvas.width).toBe(0);
    await visual.dispose();
    expect(screenCaptureState().active).toBe(true); // only the user ends window sharing
    stopWordScreenSharing();
    expect(screenCaptureState().active).toBe(false);
    expect((await visual.execute('read_word_screen', {})).ok).toBe(false);
});

test.each(['monitor', 'browser', undefined])('rejects an unapproved capture surface: %s', async (surface) => {
    const w = setup(surface === undefined ? null : surface);
    await expect(startWordScreenSharing()).rejects.toThrow(/Choose the Word application window/);
    expect(w.track.stop).toHaveBeenCalled();
    expect(screenCaptureState().active).toBe(false);
});

test('permission refusal and unsupported hosts leave capture inactive', async () => {
    const w = setup();
    w.getDisplayMedia.mockRejectedValue(new DOMException('User declined', 'NotAllowedError'));
    await expect(startWordScreenSharing()).rejects.toMatchObject({ name: 'NotAllowedError' });
    expect(screenCaptureState()).toMatchObject({ active: false, starting: false });
    delete navigator.mediaDevices;
    await expect(startWordScreenSharing()).rejects.toThrow(/does not support/);
});

test('stopping during the chooser closes a late stream', async () => {
    const w = setup(); let resolve;
    w.getDisplayMedia.mockImplementation(() => new Promise((done) => { resolve = done; }));
    const opening = startWordScreenSharing();
    stopWordScreenSharing(); resolve(w.stream); await opening;
    expect(w.track.stop).toHaveBeenCalledTimes(1);
    expect(screenCaptureState().active).toBe(false);
});

test('stopping sharing during frame acquisition never returns the stale image', async () => {
    const w = setup(); await startWordScreenSharing(); let frame;
    w.player.requestVideoFrameCallback.mockImplementation((callback) => { frame = callback; return 7; });
    const reading = readWordScreen();
    stopWordScreenSharing(); frame();
    expect(await reading).toMatchObject({ ok: false, visualInputAvailable: false });
    expect(w.canvas.toDataURL).not.toHaveBeenCalled();
});

test('capture cancellation clears pending frame callbacks', async () => {
    const w = setup(); await startWordScreenSharing();
    w.player.requestVideoFrameCallback.mockReturnValue(9);
    const controller = new AbortController();
    const reading = readWordScreen({ signal: controller.signal });
    const rejected = expect(reading).rejects.toMatchObject({ name: 'AbortError' });
    controller.abort(); await rejected;
    expect(w.player.cancelVideoFrameCallback).toHaveBeenCalledWith(9);
    expect(w.canvas.toDataURL).not.toHaveBeenCalled();
});

test.each(['zero dimensions', 'muted'])('an unavailable window frame is never reported as inspected: %s', async (kind) => {
    const w = setup(); await startWordScreenSharing();
    if (kind === 'zero dimensions') Object.defineProperty(w.player, 'videoWidth', { value: 0 });
    else w.track.muted = true;
    const reader = createWordVisualTools();
    const result = await reader.execute('read_word_screen', {});
    expect(result).toMatchObject({ ok: false, visualInputAvailable: false, error: expect.stringContaining('no visible live frame') });
    expect(reader.status().inspectedScreen).toBe(false);
    expect(w.canvas.toDataURL).not.toHaveBeenCalled();
    await reader.dispose();
});

test('native sharing termination updates the button and prevents later screenshots', async () => {
    const w = setup();
    document.body.innerHTML = '<button id="shareWordViewBtn"></button><div id="wordViewStatus" hidden></div>';
    const log = jest.fn(); initScreenSharingView({ log });
    const button = document.getElementById('shareWordViewBtn');
    button.click(); await new Promise((resolve) => setTimeout(resolve, 0));
    expect(button.getAttribute('aria-pressed')).toBe('true');
    expect(document.getElementById('wordViewStatus').hidden).toBe(false);
    const ended = w.track.addEventListener.mock.calls.find((call) => call[0] === 'ended')[1];
    ended();
    expect(button.getAttribute('aria-pressed')).toBe('false');
    expect((await readWordScreen()).ok).toBe(false);
    window.dispatchEvent(new Event('pagehide'));
});

test('sharing failure is explained in the UI without claiming visual access', async () => {
    setup(); delete navigator.mediaDevices;
    document.body.innerHTML = '<button id="shareWordViewBtn"></button><div id="wordViewStatus" hidden></div>';
    initScreenSharingView(); document.getElementById('shareWordViewBtn').click();
    await new Promise((resolve) => setTimeout(resolve, 0));
    expect(document.getElementById('wordViewStatus').textContent).toContain('does not support window sharing');
    window.dispatchEvent(new Event('pagehide'));
});

function mountToolbar() {
    document.body.innerHTML = '<button id="shareWordViewBtn"></button><div id="wordViewStatus" hidden></div>';
    const log = jest.fn(); initScreenSharingView({ log });
    return { log, button: document.getElementById('shareWordViewBtn') };
}
const tick = () => new Promise((resolve) => setTimeout(resolve, 0));
const shareDialog = () => document.querySelector('[aria-labelledby="wordShareRequestTitle"]');
const dialogButton = (label) => [...shareDialog().querySelectorAll('button')].find((el) => el.textContent === label);

test('models ask through a prompt whose button takes the toolbar sharing path', async () => {
    const w = setup(); const { log, button } = mountToolbar();
    const visual = createWordVisualTools();
    expect(visual.tools.map((tool) => tool.name)).toEqual(['read_word_screen', 'request_word_screen_share']);
    const asking = visual.execute('request_word_screen_share', { reason: '  Check the   selected heading.  ' });
    await tick();
    expect(shareDialog().textContent).toContain('Check the selected heading.');
    expect(w.getDisplayMedia).not.toHaveBeenCalled();
    dialogButton('Share Word window').click();
    expect(await asking).toMatchObject({ ok: true, result: { shared: true } });
    expect(w.getDisplayMedia).toHaveBeenCalledTimes(1);
    expect(shareDialog()).toBeNull();
    expect(button.getAttribute('aria-pressed')).toBe('true');
    expect(log).toHaveBeenCalledWith('Word window shared for model screenshots.', 'info');
    expect((await visual.execute('read_word_screen', {})).ok).toBe(true);
    expect(await visual.execute('request_word_screen_share', {})).toMatchObject({ ok: false, error: expect.stringContaining('already requested') });
    window.dispatchEvent(new Event('pagehide'));
});

test('declining, cancelling the chooser or a wrong surface never starts sharing', async () => {
    const w = setup(); mountToolbar();
    const declined = createWordVisualTools().execute('request_word_screen_share', {});
    await tick();
    expect(shareDialog().querySelector('.confirm-args').hidden).toBe(true);
    dialogButton('Not now').click();
    expect(await declined).toMatchObject({ ok: true, result: { shared: false } });
    expect(w.getDisplayMedia).not.toHaveBeenCalled();

    w.getDisplayMedia.mockRejectedValueOnce(new DOMException('User declined', 'NotAllowedError'));
    const cancelled = createWordVisualTools().execute('request_word_screen_share', {});
    await tick(); dialogButton('Share Word window').click();
    expect(await cancelled).toMatchObject({ ok: true, result: { shared: false } });
    expect(shareDialog()).toBeNull();

    w.track.getSettings = () => ({ displaySurface: 'monitor' });
    const wrongSurface = createWordVisualTools().execute('request_word_screen_share', {});
    await tick(); dialogButton('Share Word window').click(); await tick();
    expect(shareDialog().querySelector('[role="alert"]').textContent).toMatch(/Choose the Word application window/);
    shareDialog().dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    expect(await wrongSurface).toMatchObject({ ok: true, result: { shared: false } });
    expect(screenCaptureState().active).toBe(false);
    window.dispatchEvent(new Event('pagehide'));
});

test('cancelling the turn closes the share prompt', async () => {
    setup(); mountToolbar();
    const controller = new AbortController();
    const asking = createWordVisualTools({ signal: controller.signal }).execute('request_word_screen_share', {});
    await tick();
    controller.abort();
    await expect(asking).rejects.toMatchObject({ name: 'AbortError' });
    expect(shareDialog()).toBeNull();
    window.dispatchEvent(new Event('pagehide'));
});

test('the request tool is offered only when a prompt can start sharing', async () => {
    setup();
    expect(createWordVisualTools().tools.map((tool) => tool.name)).toEqual(['read_word_screen']);
    mountToolbar();
    await startWordScreenSharing();
    expect(createWordVisualTools().tools.map((tool) => tool.name)).toEqual(['read_word_screen']);
    window.dispatchEvent(new Event('pagehide'));
    expect(screenCaptureState().active).toBe(false);
    expect(createWordVisualTools().tools.map((tool) => tool.name)).toEqual(['read_word_screen']);
});

test.each(['拉起窗口共享请求', '发起屏幕共享', '请共享 Word 窗口', '看一下我的屏幕', '窗口里这段标题是什么样', 'Share my Word window',
    'request window sharing', 'What is on my screen?', 'Can you see the Word window?', 'check the visible view'])('%s asks about the Word window', (text) => {
    expect(looksLikeWordViewRequest(text)).toBe(true);
});

test.each(['停止窗口共享', '不要共享窗口', 'stop sharing the window', 'Summarize the discussion section', '这篇文章讲了什么', ''])('%s does not ask for the Word window', (text) => {
    expect(looksLikeWordViewRequest(text)).toBe(false);
});

describe('Q&A window capture', () => {
    test('stays out of the way when nothing is shared or requested', async () => {
        const w = setup(); mountToolbar();
        expect(await captureWordViewForQuestion()).toEqual({ image: null, note: '' });
        expect(shareDialog()).toBeNull();
        expect(w.getDisplayMedia).not.toHaveBeenCalled();
        window.dispatchEvent(new Event('pagehide'));
    });

    test('active sharing attaches a fresh frame without prompting', async () => {
        setup(); mountToolbar();
        await startWordScreenSharing();
        const view = await captureWordViewForQuestion({ request: true });
        expect(view.image.dataUrl).toMatch(/^data:image\/jpeg;base64,fresh/);
        expect(view.note).toContain('untrusted reference data');
        expect(shareDialog()).toBeNull();
        window.dispatchEvent(new Event('pagehide'));
    });

    test('a request prompts first, and only the click starts sharing', async () => {
        const w = setup(); mountToolbar();
        const onStatus = jest.fn();
        const capture = captureWordViewForQuestion({ request: true, onStatus });
        await tick();
        expect(onStatus).toHaveBeenCalledWith('Waiting for you to share the Word window...');
        expect(w.getDisplayMedia).not.toHaveBeenCalled();
        dialogButton('Share Word window').click();
        expect((await capture).image.dataUrl).toMatch(/^data:image\/jpeg/);
        expect(w.getDisplayMedia).toHaveBeenCalledTimes(1);
        window.dispatchEvent(new Event('pagehide'));
    });

    test('declining tells the model it has no view', async () => {
        const w = setup(); mountToolbar();
        const capture = captureWordViewForQuestion({ request: true });
        await tick();
        dialogButton('Not now').click();
        const view = await capture;
        expect(view.image).toBeNull();
        expect(view.note).toContain('the user did not share the Word window.');
        expect(view.note).toContain('Do not claim to see');
        expect(w.getDisplayMedia).not.toHaveBeenCalled();
        window.dispatchEvent(new Event('pagehide'));
    });

    test('a host without a share prompt reports the limitation', async () => {
        const view = await captureWordViewForQuestion({ request: true });
        expect(view).toMatchObject({ image: null, note: expect.stringContaining('unavailable on this Word host') });
    });

    test('a failed frame leaves the question answerable; cancelling does not', async () => {
        const w = setup(); mountToolbar();
        await startWordScreenSharing();
        w.canvas.toDataURL.mockReturnValue('not-an-image');
        const log = jest.fn();
        const view = await captureWordViewForQuestion({ log });
        expect(view).toMatchObject({ image: null, note: expect.stringContaining('exceeds the visual-input limit') });
        expect(log).toHaveBeenCalledWith(expect.stringContaining('screenshot unavailable'), 'warning');
        const controller = new AbortController(); controller.abort();
        await expect(captureWordViewForQuestion({ signal: controller.signal })).rejects.toMatchObject({ name: 'AbortError' });
        window.dispatchEvent(new Event('pagehide'));
    });

    test('cancelling closes a pending share prompt', async () => {
        setup(); mountToolbar();
        const controller = new AbortController();
        const capture = captureWordViewForQuestion({ request: true, signal: controller.signal });
        await tick();
        controller.abort();
        await expect(capture).rejects.toMatchObject({ name: 'AbortError' });
        expect(shareDialog()).toBeNull();
        window.dispatchEvent(new Event('pagehide'));
    });
});
