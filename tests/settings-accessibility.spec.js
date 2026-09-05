/** @jest-environment jsdom */
const fs = require('fs');
const path = require('path');
jest.mock('../src/lib/llm-client.js', () => ({ testConnection: jest.fn() }));
const { testConnection } = require('../src/lib/llm-client.js');
const { initSettings, openSettings, closeSettings, testConnectionUI } = require('../src/taskpane/ui/settings-view.js');
const { appState, defaultConfig } = require('../src/taskpane/app-state.js');
const { setConnectionStatus } = require('../src/taskpane/ui/status-bar.js');

beforeEach(() => {
    const html = fs.readFileSync(path.join(__dirname, '../src/taskpane/taskpane.html'), 'utf8');
    document.body.innerHTML = html.slice(html.indexOf('<body>') + 6, html.indexOf('</body>'));
    appState.config = defaultConfig();
    global.fetch = jest.fn(async () => ({ ok: false }));
    testConnection.mockResolvedValue({ models: [{ id: 'alpha' }, { id: 'beta' }] });
    initSettings();
});
afterEach(() => { closeSettings(); });

test('real taskpane markup integrates IME send, DOM confirmation and Settings entry', async () => {
    const { initInputBar } = require('../src/taskpane/ui/input-bar.js');
    const onSubmit = jest.fn();
    initInputBar({
        onSubmit, onCancel: jest.fn(), getSkills: () => [], onOpenSettings: openSettings,
        getAutoApply: () => appState.config.autoApplyChanges,
        getTrackChanges: () => appState.config.trackChangesEnabled,
        setAutoApply: (enabled) => { appState.config.autoApplyChanges = enabled; },
    });
    const input = document.getElementById('chatInput');
    input.focus();
    input.value = 'Review this paragraph';
    input.dispatchEvent(new CompositionEvent('compositionstart', { bubbles: true }));
    input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', isComposing: true, bubbles: true }));
    expect(onSubmit).not.toHaveBeenCalled();
    input.dispatchEvent(new CompositionEvent('compositionend', { bubbles: true }));
    input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
    expect(onSubmit).toHaveBeenCalledWith('Review this paragraph', []);
    const toggle = document.getElementById('autoApplyToggle');
    toggle.focus();
    toggle.click();
    expect(document.activeElement.dataset.confirm).toBe('cancel');
    expect(appState.config.autoApplyChanges).toBe(false);
    document.querySelector('[data-confirm="enable"]').click();
    await Promise.resolve();
    expect(appState.config.autoApplyChanges).toBe(true);
    expect(document.activeElement).toBe(toggle);
    const model = document.getElementById('modelPill');
    model.focus();
    model.click();
    expect(document.activeElement.id).toBe('settingsCloseBtn');
    document.activeElement.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    expect(document.getElementById('settingsOverlay').hidden).toBe(true);
    expect(document.activeElement).toBe(model);
});

test('settings enters, traps and restores focus; model keyboard selection updates ARIA', async () => {
    const opener = document.getElementById('settingsBtn');
    opener.focus();
    openSettings();
    expect(document.activeElement.id).toBe('settingsCloseBtn');
    opener.focus();
    expect(document.activeElement.id).toBe('settingsCloseBtn');
    await testConnectionUI();
    const input = document.getElementById('modelSelect');
    input.focus();
    input.value = 'bet';
    input.dispatchEvent(new Event('input'));
    input.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowDown', bubbles: true }));
    const option = document.getElementById(input.getAttribute('aria-activedescendant'));
    expect(option.textContent).toBe('beta');
    expect(option.getAttribute('aria-selected')).toBe('true');
    input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
    expect(input.value).toBe('beta');
    expect(input.getAttribute('aria-expanded')).toBe('false');
    input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    expect(document.getElementById('settingsOverlay').hidden).toBe(true);
    expect(document.activeElement).toBe(opener);
});

test('Settings tabs and nested prompt modal remain keyboard accessible', () => {
    openSettings();
    const general = document.getElementById('settingsTabGeneral');
    general.focus();
    general.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowRight', bubbles: true }));
    expect(document.activeElement.id).toBe('settingsTabPrompts');
    expect(general.tabIndex).toBe(-1);
    const select = document.getElementById('promptSelect-amendment');
    select.value = '';
    document.getElementById('promptTextarea-amendment').value = 'Rewrite';
    const save = document.getElementById('savePromptBtn-amendment');
    save.focus();
    save.click();
    expect(document.activeElement.id).toBe('promptName');
    document.getElementById('promptName').dispatchEvent(new KeyboardEvent('keydown', { key: 'Tab', shiftKey: true, bubbles: true }));
    expect(document.activeElement.id).toBe('savePromptCancelBtn');
    document.activeElement.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    expect(document.activeElement).toBe(save);
    expect(document.getElementById('settingsOverlay').hidden).toBe(false);
});

test('connection errors are visible in the composer and settings, then clear on success', () => {
    setConnectionStatus('error', 'API key required');
    expect(document.getElementById('connectionError').hidden).toBe(false);
    expect(document.getElementById('settingsConnectionStatus').textContent).toBe('API key required');
    setConnectionStatus('connected', 'Connected');
    expect(document.getElementById('connectionError').hidden).toBe(true);
    expect(document.getElementById('settingsConnectionStatus').textContent).toBe('Connected');
});

test('turning Track Changes off disables automatic writes and explains why', () => {
    appState.config.autoApplyChanges = true;
    const tracking = document.getElementById('trackChangesCheckbox');
    tracking.checked = false;
    tracking.dispatchEvent(new Event('change'));
    expect(appState.config.autoApplyChanges).toBe(false);
    expect(document.getElementById('inputError').textContent).toContain('Track Changes is off');
});
