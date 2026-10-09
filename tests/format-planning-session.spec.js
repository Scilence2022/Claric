const { planFormatWithRendering } = require('../src/taskpane/format-planning-session.js');
const { WORD_RENDER_TOOL_SPECS } = require('../src/taskpane/word-render-tools.js');
const { WORD_SCREEN_TOOL_SPEC } = require('../src/taskpane/word-screen-capture.js');
const call = (tool, args = {}) => JSON.stringify({ tool, args });
function renderer() {
    const status = { inspectedPages: [], unavailable: null };
    return { tools: WORD_RENDER_TOOL_SPECS, status: () => ({ ...status }),
        execute: jest.fn(async (name) => {
            if (name === 'list_rendered_pages') return { ok: true, result: { pages: [{ pageNumber: 2 }], visualInputAvailable: false } };
            status.inspectedPages.push(2);
            return { ok: true, result: { source: 'word_native_pdf', visualInputAvailable: true }, attachments: [{ dataUrl: 'data:image/jpeg;base64,realpage' }] };
        }),
    };
}

test('an authorized screenshot fulfills visual inspection without exporting PDF pages', async () => {
    let inspectedScreen = false;
    const renderer = { tools: [WORD_SCREEN_TOOL_SPEC], status: () => ({ preferredSource: 'word_shared_window', inspectedScreen, inspectedPages: [], unavailable: null }),
        execute: jest.fn(async () => { inspectedScreen = true; return { ok: true, result: { source: 'word_shared_window', visualInputAvailable: true }, attachments: [{ dataUrl: 'data:image/jpeg;base64,screen' }] }; }) };
    const replies = [call('read_word_screen'), call('propose_format_ops', { ops: [{ font: { bold: false } }] }), call('finish', { summary: 'Ready for review' })];
    const send = jest.fn(async () => replies.shift());
    const result = await planFormatWithRendering({ prompt: 'format', scopeText: 'Example text', renderer, send });
    expect(result.rendering).toMatchObject({ inspectedScreen: true, modelVisualInputAccepted: true });
    expect(renderer.execute).toHaveBeenCalledTimes(1);
    expect(renderer.execute).toHaveBeenCalledWith('read_word_screen', {});
    expect(JSON.stringify(send.mock.calls[0][0])).toContain('word_shared_window');
});

test('model calls visual tools before staging, and receives page images as multimodal messages', async () => {
    const render = renderer();
    const replies = [call('propose_format_ops', { ops: [{ font: { bold: false } }] }),
        call('list_rendered_pages', { search: 'Example heading' }), call('read_rendered_pages', { pages: [2] }),
        call('propose_format_ops', { ops: [{ font: { bold: false } }] }), call('finish', { summary: 'Formatting ready for review' })];
    const send = jest.fn(async () => replies.shift());
    const result = await planFormatWithRendering({ prompt: 'format', scopeText: 'Example heading', renderer: render, send });
    expect(result.ops).toEqual([{ font: { bold: false } }]);
    expect(result.rendering.modelVisualInputAccepted).toBe(true);
    expect(JSON.stringify(send.mock.calls[1][0])).toContain('Inspect the relevant Word-rendered pages');
    expect(send.mock.calls[3][0].some((message) => Array.isArray(message.content)
        && message.content.some((part) => part.image_url?.url === 'data:image/jpeg;base64,realpage'))).toBe(true);
});

test('text-only models receive an explicit fallback, and unsupported images are not replayed each step', async () => {
    const render = renderer(); let rejected = false;
    const replies = [call('list_rendered_pages'), call('read_rendered_pages', { pages: [2] }),
        call('propose_format_ops', { ops: [{ font: { bold: false } }] }), call('finish', { summary: 'Text-only formatting draft' })];
    const send = jest.fn(async (messages) => {
        if (!rejected && JSON.stringify(messages).includes('image_url')) { rejected = true; throw new Error('HTTP 422: no vision'); }
        return replies.shift();
    });
    const result = await planFormatWithRendering({ prompt: 'format', scopeText: 'heading', renderer: render, send });
    expect(result.rendering.modelVisualInputAccepted).toBe(false);
    expect(result.rendering.visualUnavailable).toContain('No visual assessment');
    expect(send.mock.calls.slice(3).every(([messages]) => !JSON.stringify(messages).includes('image_url'))).toBe(true);
});

test('native export failure permits an explicitly limited text-based draft', async () => {
    const render = renderer();
    render.execute.mockImplementation(async () => {
        render.status = () => ({ inspectedPages: [], unavailable: 'Native export failed' });
        return { ok: false, error: 'Native export failed' };
    });
    const replies = [call('list_rendered_pages'), call('propose_format_ops', { ops: [{ font: { bold: false } }] }), call('finish', { summary: 'Text evidence only' })];
    const result = await planFormatWithRendering({ prompt: 'format', scopeText: '', renderer: render, send: async () => replies.shift() });
    expect(result.rendering).toMatchObject({ unavailable: 'Native export failed', modelVisualInputAccepted: false });
});

test('a finish without any draft never approves a formatting proposal', async () => {
    await expect(planFormatWithRendering({ prompt: 'format', scopeText: '', renderer: renderer(),
        send: async () => call('finish', { summary: 'Done' }) })).rejects.toThrow(/did not complete/);
});

test('unmatched native targets are rejected before staging and the model receives verified counts after correction', async () => {
    const render = renderer();
    const wrong = [{ paragraphIds: ['p999'], paragraph: { alignment: 'justified' } }];
    const right = [{ paragraphRole: 'body', paragraph: { alignment: 'justified' } }];
    const replies = [call('read_rendered_pages', { pages: [2] }), call('propose_format_ops', { ops: wrong }),
        call('propose_format_ops', { ops: right }), call('finish', { summary: 'Verified body alignment proposal' })];
    const send = jest.fn(async () => replies.shift());
    const validateOps = jest.fn(async (ops) => {
        if (ops[0].paragraphIds) throw new Error('No verified formatting paragraphs matched.');
        return { verifiedParagraphs: 12, excludedParagraphs: 3 };
    });
    const result = await planFormatWithRendering({ prompt: 'format', scopeText: 'Body', renderer: render, send, validateOps });
    expect(result.ops).toEqual(right);
    expect(JSON.stringify(send.mock.calls[2][0])).toContain('No verified formatting paragraphs matched');
    const observations = send.mock.calls[3][0].filter((message) => message.role === 'user' && typeof message.content === 'string')
        .map((message) => { try { return JSON.parse(message.content); } catch { return {}; } });
    expect(observations.find((observation) => observation.result?.targets)?.result.targets)
        .toEqual({ verifiedParagraphs: 12, excludedParagraphs: 3 });
});
