const fs = require('fs');
const path = require('path');
const { JSDOM } = require('jsdom');

const root = path.resolve(__dirname, '..');
const html = fs.readFileSync(path.join(root, 'docs/index.html'), 'utf8');
const script = fs.readFileSync(path.join(root, 'docs/assets/site.js'), 'utf8');
let dom;

function openSite({ query = '', saved, language = 'en-US', blocked = false, execute = true } = {}) {
    dom = new JSDOM(html, { url: `https://scilence2022.github.io/Claric/${query}`, runScripts: 'outside-only' });
    const { window } = dom;
    Object.defineProperty(window.navigator, 'language', { value: language });
    if (saved !== undefined) window.localStorage.setItem('claric.site.lang', saved);
    if (blocked) Object.defineProperty(window, 'localStorage', { get() { throw new Error('Storage disabled'); } });
    if (execute) window.eval(script);
    return window;
}

afterEach(() => dom?.window.close());

describe('public website delivery contract', () => {
    test('preserves navigation anchors and unique identifiers', () => {
        const { document } = openSite();
        const ids = [...document.querySelectorAll('[id]')].map(el => el.id);
        expect(new Set(ids).size).toBe(ids.length);
        for (const id of ['main', 'how', 'gallery', 'features', 'scenarios', 'models', 'start']) {
            expect(document.getElementById(id)).not.toBeNull();
        }
        expect(document.querySelectorAll('h1')).toHaveLength(1);
        expect(document.querySelector('h1').textContent).toBe('Claric');
        for (const link of document.querySelectorAll('a[href^="#"]')) {
            expect(document.getElementById(link.hash.slice(1))).not.toBeNull();
        }
    });

    test('ships every referenced local asset and avoids external runtime dependencies', () => {
        const { document } = openSite();
        for (const el of document.querySelectorAll('[src],link[href],a[href]')) {
            const value = el.getAttribute('src') || el.getAttribute('href');
            if (/^(https?:|#)/.test(value)) continue;
            expect(fs.existsSync(path.join(root, 'docs', value))).toBe(true);
        }
        expect([...document.scripts].every(el => el.src.startsWith('https://scilence2022.github.io/Claric/'))).toBe(true);
        expect(html).not.toMatch(/fonts\.googleapis|hero-demo\.webp|IntersectionObserver/);
        const images = [...document.querySelectorAll('img')];
        expect(images.every(img => Number(img.width) > 0 && Number(img.height) > 0)).toBe(true);
        expect(document.getElementById('hero-image').getAttribute('fetchpriority')).toBe('high');
        expect(document.querySelector('link[rel="canonical"]').href).toBe('https://scilence2022.github.io/Claric/');
    });

    test('retains visible content and all installation routes without JavaScript', () => {
        const { document } = openSite({ execute: false });
        expect(document.documentElement.dataset.lang).toBe('en');
        expect(document.getElementById('langToggle').hidden).toBe(true);
        expect(document.getElementById('install-tabs').hidden).toBe(true);
        expect([...document.querySelectorAll('[data-install-panel]')].every(el => !el.hidden)).toBe(true);
        expect(document.querySelectorAll('#faq details summary').length).toBeGreaterThan(3);
    });

    test.each([
        [{ query: '?lang=zh', saved: 'en' }, 'zh'],
        [{ query: '?lang=en', saved: 'zh', language: 'zh-CN' }, 'en'],
        [{ query: '?lang=invalid', saved: 'zh' }, 'zh'],
        [{ saved: 'invalid', language: 'zh-TW' }, 'zh'],
        [{ saved: 'invalid', language: 'fr-FR' }, 'en'],
        [{ blocked: true, language: 'zh-CN' }, 'zh'],
    ])('resolves language preferences %j', (options, expected) => {
        expect(openSite(options).document.documentElement.dataset.lang).toBe(expected);
    });

    test('switches visible copy, accessible labels, image alternatives and shareable URL', () => {
        const window = openSite();
        const { document } = window;
        document.getElementById('langToggle').click();
        expect(document.documentElement.lang).toBe('zh-CN');
        expect(document.title).toContain('Word');
        expect(document.getElementById('hero-image').alt).toContain('修订');
        expect(document.getElementById('langToggle').getAttribute('aria-label')).toBe('Switch to English');
        expect(window.localStorage.getItem('claric.site.lang')).toBe('zh');
        expect(window.location.search).toBe('?lang=zh');
        expect([...document.querySelectorAll('span[data-lang="en"]')].every(el => el.hidden)).toBe(true);
        document.getElementById('langToggle').click();
        expect(document.documentElement.lang).toBe('en');
    });

    test('switches with unavailable storage and history', () => {
        const window = openSite({ blocked: true });
        window.history.replaceState = () => { throw new Error('History disabled'); };
        window.document.getElementById('langToggle').click();
        expect(window.document.documentElement.dataset.lang).toBe('zh');
    });

    test('supports installation clicks, keyboard traversal and deep links', () => {
        const window = openSite({ query: '#install-windows' });
        const { document } = window;
        const windows = document.getElementById('tab-windows');
        expect(windows.getAttribute('aria-selected')).toBe('true');
        windows.focus();
        windows.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'ArrowRight', bubbles: true }));
        expect(document.activeElement.id).toBe('tab-selfhost');
        expect(document.getElementById('install-selfhost').hidden).toBe(false);
        document.activeElement.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'Home', bubbles: true }));
        expect(document.activeElement.id).toBe('tab-macos');
        document.getElementById('tab-windows').click();
        document.getElementById('langToggle').click();
        expect(document.getElementById('install-windows').hidden).toBe(false);
        expect(document.getElementById('install-macos').hidden).toBe(true);
        expect(document.querySelectorAll('[role="tab"][tabindex="0"]')).toHaveLength(1);
        window.location.hash = '#install-selfhost';
        window.dispatchEvent(new window.HashChangeEvent('hashchange'));
        expect(document.getElementById('install-selfhost').hidden).toBe(false);
    });
});
