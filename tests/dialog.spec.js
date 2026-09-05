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
