/** @jest-environment jsdom */

/**
 * Status-bar drawer wiring tests.
 *
 * Regression: initStatusBar() used to bind #logBtn AND taskpane.js bound it
 * again at bootstrap, so one click fired toggleLogDrawer twice and the
 * drawer could never open. initStatusBar must leave the header button to
 * the bootstrap caller (its own docstring says exactly that); only
 * close/clear controls are bound here.
 */

const { initStatusBar, toggleLogDrawer, isLogDrawerOpen } = require('../src/taskpane/ui/status-bar.js');

function setupDom() {
    document.body.innerHTML = `
        <button id="logBtn" aria-expanded="false"></button>
        <div id="logDrawer" hidden></div>
        <button id="logDrawerCloseBtn"></button>
        <button id="clearLogsBtn"></button>
        <div id="logs"></div>`;
}

describe('log drawer wiring', () => {
    beforeEach(() => {
        setupDom();
    });

    it('initStatusBar does not bind #logBtn (bootstrap owns that handler)', () => {
        initStatusBar();
        const btn = document.getElementById('logBtn');
        const before = btn.getAttribute('aria-expanded');
        btn.click();
        // No listener attached: the drawer must still be closed and the
        // aria state untouched.
        expect(isLogDrawerOpen()).toBe(false);
        expect(btn.getAttribute('aria-expanded')).toBe(before);
    });

    it('a single toggle opens the drawer and reports open', () => {
        expect(toggleLogDrawer()).toBe(true);
        expect(isLogDrawerOpen()).toBe(true);
        expect(document.getElementById('logBtn').getAttribute('aria-expanded')).toBe('true');
        // And toggling again closes it.
        expect(toggleLogDrawer()).toBe(false);
        expect(isLogDrawerOpen()).toBe(false);
    });
});
