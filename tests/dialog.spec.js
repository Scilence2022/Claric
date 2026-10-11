/** @jest-environment jsdom */
const { confirmAutoApply, containFocus } = require('../src/taskpane/ui/dialog.js');

test.each(['outside', 'hidden', 'disabled'])('focus entry rejects %s initial focus and skips hidden ancestors', (target) => {
    document.body.innerHTML = '<button id="outside">Outside</button><section id="root"><div style="display:none"><button id="hidden">Hidden</button></div><button id="disabled" disabled>Disabled</button><button id="first">First</button><button id="last">Last</button></section>';
    const root = document.getElementById('root');
    const release = containFocus(root, () => {}, document.getElementById(target));
    expect(document.activeElement.id).toBe('first');
    document.activeElement.dispatchEvent(new KeyboardEvent('keydown', { key: 'Tab', shiftKey: true, bubbles: true }));
    expect(document.activeElement.id).toBe('last');
    release();
});

test('confirmation starts at Cancel, traps focus, escapes and restores focus', async () => {
    document.body.innerHTML = '<button id="opener">Auto-apply</button>';
    const opener = document.getElementById('opener');
    opener.focus();
    const answer = confirmAutoApply();
    const cancel = document.querySelector('[data-confirm="cancel"]');
    const enable = document.querySelector('[data-confirm="enable"]');
    expect(document.activeElement).toBe(cancel);
    cancel.dispatchEvent(new KeyboardEvent('keydown', { key: 'Tab', shiftKey: true, bubbles: true }));
    expect(document.activeElement).toBe(enable);
    enable.dispatchEvent(new KeyboardEvent('keydown', { key: 'Tab', bubbles: true }));
    expect(document.activeElement).toBe(cancel);
    opener.focus();
    expect(document.activeElement).toBe(cancel);
    cancel.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    expect(await answer).toBe(false);
    expect(document.activeElement).toBe(opener);
    expect(document.querySelector('[role="alertdialog"]')).toBeNull();
});

describe('confirmMcpToolCall', () => {
    const { confirmMcpToolCall, MCP_ARGS_PREVIEW_CHARS } = require('../src/taskpane/ui/dialog.js');

    test('renders server, tool and arguments as text, never markup', async () => {
        document.body.innerHTML = '';
        const answer = confirmMcpToolCall({ serverName: '<b>srv</b>', tool: '<img src=x onerror=alert(1)>',
            args: { body: '<script>alert(1)</script>' } });
        const dialog = document.querySelector('[role="alertdialog"]');
        expect(dialog.querySelector('img, script, b')).toBeNull();
        expect(dialog.textContent).toContain('<img src=x onerror=alert(1)>');
        expect(dialog.querySelector('pre').textContent).toContain('<script>alert(1)</script>');
        document.querySelector('[data-confirm="deny"]').click();
        expect(await answer).toBe('deny');
    });

    test.each(['once', 'run', 'deny'])('the %s button resolves to its decision and closes the dialog', async (value) => {
        document.body.innerHTML = '';
        const answer = confirmMcpToolCall({ serverName: 's', tool: 't', args: {} });
        document.querySelector(`[data-confirm="${value}"]`).click();
        expect(await answer).toBe(value);
        expect(document.querySelector('[role="alertdialog"]')).toBeNull();
    });

    test('focus starts on Deny and Escape denies', async () => {
        document.body.innerHTML = '';
        const answer = confirmMcpToolCall({ serverName: 's', tool: 't', args: {} });
        const deny = document.querySelector('[data-confirm="deny"]');
        expect(document.activeElement).toBe(deny);
        deny.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
        expect(await answer).toBe('deny');
    });

    test('long arguments are truncated in the preview', async () => {
        document.body.innerHTML = '';
        const answer = confirmMcpToolCall({ serverName: 's', tool: 't', args: { text: 'x'.repeat(MCP_ARGS_PREVIEW_CHARS * 2) } });
        expect(document.querySelector('pre').textContent).toMatch(/more characters\)$/);
        document.querySelector('[data-confirm="deny"]').click();
        await answer;
    });

    test('cancelling the turn closes the dialog with an AbortError', async () => {
        document.body.innerHTML = '';
        const controller = new AbortController();
        const answer = confirmMcpToolCall({ serverName: 's', tool: 't', args: {} }, { signal: controller.signal });
        controller.abort();
        await expect(answer).rejects.toMatchObject({ name: 'AbortError' });
        expect(document.querySelector('[role="alertdialog"]')).toBeNull();
        await expect(confirmMcpToolCall({ serverName: 's', tool: 't', args: {} }, { signal: controller.signal }))
            .rejects.toMatchObject({ name: 'AbortError' });
        expect(document.querySelector('[role="alertdialog"]')).toBeNull();
    });
});
