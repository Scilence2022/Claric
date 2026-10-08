import appMetadata from '../../../package.json';
import { containFocus } from './dialog.js';

/** Show app and host details in About, using the same metadata as the activity log. */
export function initAboutView({ platform, wordApiVersion, log }) {
    const button = document.getElementById('infoBtn');
    const overlay = document.getElementById('aboutOverlay');
    const panel = document.getElementById('aboutPanel');
    const closeButton = document.getElementById('aboutCloseBtn');
    const appVersion = document.getElementById('aboutVersion');
    const build = document.getElementById('aboutBuild');
    const builtAt = document.getElementById('aboutBuiltAt');
    appVersion.textContent = appMetadata.version;
    document.getElementById('aboutPlatform').textContent = platform || 'Unknown';
    document.getElementById('aboutWordApi').textContent = wordApiVersion || 'Unknown';

    let releaseFocus;
    const close = () => {
        overlay.classList.remove('active');
        releaseFocus?.();
        releaseFocus = null;
    };
    button.addEventListener('click', () => {
        if (overlay.classList.contains('active')) return;
        button.focus();
        overlay.classList.add('active');
        releaseFocus = containFocus(panel, close, closeButton);
    });
    closeButton.addEventListener('click', close);
    overlay.addEventListener('click', (event) => { if (event.target === overlay) close(); });

    // Load without delaying startup. Keep About usable when this optional asset
    // is missing (for example in development), blocked, or slow to respond.
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), 8000);
    fetch(new URL('build-info.json', location.href).href, { cache: 'no-store', signal: controller.signal })
        .then((response) => (response.ok ? response.json() : null))
        .then((info) => {
            if (!info || typeof info.hash !== 'string' || !/^[a-f0-9]{12}$/.test(info.hash)) return;
            if (typeof info.appVersion === 'string' && info.appVersion.trim()) appVersion.textContent = info.appVersion;
            build.textContent = info.hash;
            const date = typeof info.builtAt === 'string' ? new Date(info.builtAt) : null;
            builtAt.textContent = date && Number.isFinite(date.getTime())
                ? date.toISOString().replace('T', ' ').replace(/\.\d{3}Z$/, ' UTC') : 'Unavailable';
            log(`Claric build: ${info.hash} (${appVersion.textContent}${date && Number.isFinite(date.getTime()) ? `, ${date.toISOString()}` : ''})`, 'info');
        })
        .catch(() => { /* Optional build metadata is not an application error. */ })
        .finally(() => {
            clearTimeout(timeout);
            if (build.textContent === 'Loading…') build.textContent = 'Unavailable';
            if (builtAt.textContent === 'Loading…') builtAt.textContent = 'Unavailable';
        });
}
