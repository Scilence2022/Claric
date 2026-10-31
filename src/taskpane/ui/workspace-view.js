import { containFocus } from './dialog.js';
import { allowsDocumentAccess } from '../../lib/coordination/access.cjs';

/** Document-first workspace. It describes incoming permissions, never grants remote writes. */
export function createWorkspaceView({ identity, access, onAccessChange, onOpenSettings = () => {}, clock = Date.now } = {}) {
    const el = (id) => document.getElementById(id);
    const overlay = el('workspaceOverlay');
    const button = el('workspaceBtn');
    const list = el('workspaceDocumentList');
    const controller = new AbortController();
    const options = { signal: controller.signal };
    let documents = new Map();
    let connected = false;
    let saving = false;
    let releaseFocus;
    let signature = '';
    function selectTab(evidence = false) {
        for (const [name, selected] of [['Documents', !evidence], ['Evidence', evidence]]) {
            const tab = el(`workspaceTab${name}`);
            tab?.setAttribute('aria-selected', String(selected));
            if (tab) tab.tabIndex = selected ? 0 : -1;
            el(`workspace${name}`)?.toggleAttribute('hidden', !selected);
        }
    }
    function close() {
        overlay?.setAttribute('hidden', ''); button?.setAttribute('aria-expanded', 'false');
        releaseFocus?.(); releaseFocus = null;
    }
    function open() {
        if (!overlay) return;
        selectTab(); render(); overlay.hidden = false; button?.setAttribute('aria-expanded', 'true');
        releaseFocus ||= containFocus(el('workspacePanel'), close, el('workspaceCloseBtn'));
    }
    function node(tag, className, text) {
        const element = document.createElement(tag);
        if (className) element.className = className;
        if (text !== undefined) element.textContent = text;
        return element;
    }
    function permission(card, documentId, operation, title, checked, disabled) {
        const label = node('label', 'workspace-permission');
        const input = node('input'); input.type = 'checkbox'; input.checked = checked; input.disabled = disabled || saving;
        input.dataset.focusKey = `${documentId || 'defaults'}:${operation}`;
        input.setAttribute('aria-label', `${title} — ${card.dataset.name}`);
        input.addEventListener('change', async () => {
            const allowed = input.checked;
            saving = true; render(true);
            try {
                await onAccessChange(documentId, operation, allowed);
                el('workspaceAccessStatus').textContent = 'Permissions saved. Previously read copies are not removed.';
            } catch (error) { el('workspaceAccessStatus').textContent = `Permission update: ${error.message}`; }
            finally {
                saving = false; render(true);
                [...overlay.querySelectorAll('[data-focus-key]')].find((item) => item.dataset.focusKey === input.dataset.focusKey)?.focus();
            }
        }, options);
        label.append(input, node('span', '', title)); card.append(label);
    }
    function cardFor(entry, own, policy) {
        const card = node('article', 'workspace-document');
        const name = entry.title || entry.displayName || (own ? 'Current Word document' : 'Word document');
        card.dataset.name = name;
        const heading = node('div', 'workspace-document-heading');
        heading.append(node('strong', 'workspace-document-name', name), node('span', `workspace-badge${connected ? ' is-connected' : ''}`, own ? 'This document' : 'Connected'));
        card.append(heading);
        if (own) card.append(node('p', 'help-text', connected ? 'Default access for linked documents. Individual restrictions appear below.' : 'Connect in settings to discover other documents.'));
        else {
            card.append(node('p', 'help-text workspace-instance', `Window ${entry.instanceId.slice(-6)} · access to this document`));
            const actions = node('div', 'workspace-document-actions');
            for (const [label, id, operation] of [['Read context', 'crossDocumentContextBtn', 'readContext'], ['Use as target', null, 'requestEdits']]) {
                const action = node('button', 'btn btn-compact', label); action.type = 'button';
                action.dataset.focusKey = `${entry.instanceId}:${operation}`;
                action.disabled = !allowsDocumentAccess(entry.access, identity.documentId, operation);
                if (action.disabled) action.title = 'Access is disabled by the target document';
                action.addEventListener('click', () => {
                    const target = el('crossDocumentTarget'); target.value = entry.instanceId; target.dispatchEvent(new Event('change'));
                    if (id) el(id).click();
                    else { el('crossDocumentTaskType').focus(); el('crossDocumentBar').scrollIntoView?.({ block: 'nearest' }); }
                }, options);
                actions.append(action);
            }
            card.append(actions);
        }
        const peer = policy.peers.find((item) => item.documentId === entry.documentId);
        permission(card, own ? null : entry.documentId, 'readContext', own ? 'Allow context reads' : 'Can read this document', own ? policy.readContext : policy.readContext && peer?.readContext !== false, !own && !policy.readContext);
        permission(card, own ? null : entry.documentId, 'requestEdits', own ? 'Allow edit requests' : 'Can request edits here', own ? policy.requestEdits : policy.requestEdits && peer?.requestEdits !== false, !own && !policy.requestEdits);
        return card;
    }
    function render(force = false) {
        if (!list || !access) return;
        const peers = connected ? [...documents.values()].filter((item) => item.expiresAt > clock() && item.instanceId !== identity.instanceId) : [];
        const policy = access.get();
        const key = JSON.stringify([peers, policy, connected, saving]);
        el('workspaceCount').textContent = String(peers.length); el('workspaceCount').hidden = peers.length === 0;
        button?.setAttribute('aria-label', `Cross-document workspace${peers.length ? `, ${peers.length} connected documents` : ''}`);
        el('workspaceListCount').textContent = String(peers.length);
        el('workspaceEmpty').hidden = peers.length > 0;
        if (!force && signature === key) return;
        signature = key;
        const focused = document.activeElement?.dataset.focusKey;
        el('workspaceOwnDocument').replaceChildren(cardFor(identity, true, policy));
        list.replaceChildren(...peers.map((entry) => cardFor(entry, false, policy)));
        if (focused) [...overlay.querySelectorAll('[data-focus-key]')].find((item) => item.dataset.focusKey === focused)?.focus();
    }
    button?.addEventListener('click', open, options);
    el('workspaceCloseBtn')?.addEventListener('click', close, options);
    overlay?.addEventListener('click', (event) => { if (event.target === overlay) close(); }, options);
    el('workspaceSettingsBtn')?.addEventListener('click', () => { close(); void onOpenSettings(); }, options);
    const tabs = ['Documents', 'Evidence'].map((name) => el(`workspaceTab${name}`));
    tabs.forEach((tab, index) => {
        tab?.addEventListener('click', () => selectTab(index === 1), options);
        tab?.addEventListener('keydown', (event) => {
            if (!['ArrowLeft', 'ArrowRight', 'Home', 'End'].includes(event.key)) return;
            event.preventDefault(); const next = event.key === 'Home' ? 0 : event.key === 'End' ? 1 : 1 - index;
            selectTab(next === 1); tabs[next].focus();
        }, options);
    });
    render();
    return {
        open, close, refresh: () => render(true),
        update(next, online = connected) { documents = next; connected = online; render(); },
        connection(text, online) { connected = online; if (el('workspaceConnectionStatus')) el('workspaceConnectionStatus').textContent = text; render(); },
        dispose() { close(); controller.abort(); },
    };
}
