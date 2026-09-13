/**
 * Document identity: distinguishes the open Word file from every other file
 * without leaking its path, and (when possible) survives reloads and detects
 * Save As / Copy forks.
 *
 * Layers, from most to least durable:
 *   custom-xml — an identity part stored in the document (shared-API
 *     customXmlParts, Word 2013+) whose recorded URL hash matches the current
 *     URL. Only the URL *hash* is stored, never the path.
 *   fork — a stored identity whose URL hash no longer matches: the file was
 *     duplicated or moved. The identity is derived deterministically from
 *     (stored id + new URL) so the original keeps its id and the fork is
 *     stable across its own reloads without rewriting the file at startup.
 *   url — no stored identity; the documentId is the URL hash. Opening the
 *     file never marks it modified; the identity part is written lazily by
 *     persistDocumentIdentity the first time the user actually links this
 *     document into a cross-document workspace.
 *   ephemeral — the host exposes no URL and no stored identity (unsaved or
 *     unidentified document). A random id per load; history does not carry
 *     over, matching the previous behavior.
 *
 * @module document-identity
 */

import { newId } from './message-shape.js';

const PART_XMLNS = 'urn:claric:identity';
const ID_PATTERN = /^document-[A-Za-z0-9._:-]{1,118}$/;

async function digest(text, cryptoImpl) {
    if (!cryptoImpl?.subtle) return null;
    const result = await cryptoImpl.subtle.digest('SHA-256', new TextEncoder().encode(text));
    return Array.from(new Uint8Array(result), (byte) => byte.toString(16).padStart(2, '0')).join('');
}

function randomDocumentId(cryptoImpl) {
    if (cryptoImpl?.getRandomValues) {
        const bytes = new Uint8Array(32);
        cryptoImpl.getRandomValues(bytes);
        return `document-${Array.from(bytes, (byte) => byte.toString(16).padStart(2, '0')).join('')}`;
    }
    return newId('document');
}

function documentUrl(office) {
    try { return typeof office?.context?.document?.url === 'string' ? office.context.document.url : ''; }
    catch { return ''; }
}

function customXmlParts(office) {
    try { return office?.context?.document?.customXmlParts || null; }
    catch { return null; }
}

function succeeded(result) {
    return !!result && result.status === 'succeeded';
}

function escapeXml(text) {
    return String(text).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function identityXml(documentId, urlHash, createdAt) {
    const payload = { version: 1, documentId, urlHash, createdAt };
    return `<claricIdentity xmlns="${PART_XMLNS}" version="1">${escapeXml(JSON.stringify(payload))}</claricIdentity>`;
}

function unescapeXml(text) {
    return String(text).replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&');
}

const PART_WRAPPER = /^<claricIdentity xmlns="urn:claric:identity" version="1">([\s\S]{0,4096})<\/claricIdentity>$/;

function parseIdentityXml(xml) {
    const match = PART_WRAPPER.exec(String(xml || ''));
    if (!match) return null;
    let payload;
    try { payload = JSON.parse(unescapeXml(match[1])); } catch { return null; }
    if (!payload || payload.version !== 1 || !ID_PATTERN.test(payload.documentId || '')) return null;
    if (payload.urlHash !== null && !/^[a-f0-9]{64}$/.test(payload.urlHash || '')) return null;
    if (!Number.isSafeInteger(payload.createdAt) || payload.createdAt < 0) return null;
    return { documentId: payload.documentId, urlHash: payload.urlHash, createdAt: payload.createdAt };
}

function callAsync(target, method, ...args) {
    return new Promise((resolve) => {
        try { target[method](...args, resolve); }
        catch { resolve(null); }
    });
}

async function readIdentityPart(parts) {
    const listed = await callAsync(parts, 'getByNamespaceAsync', PART_XMLNS);
    if (!succeeded(listed) || !Array.isArray(listed.value) || listed.value.length === 0) return null;
    for (const part of listed.value) {
        const xmlResult = await callAsync(part, 'getXmlAsync');
        if (!succeeded(xmlResult)) continue;
        const parsed = parseIdentityXml(xmlResult.value);
        if (parsed) return { ...parsed, part };
    }
    return null;
}

/**
 * Resolves the identity of the currently open document. Read-only: never
 * modifies the document, so opening a file does not dirty it.
 *
 * @param {object} [deps]
 * @param {object} [deps.office]
 * @param {object} [deps.locationObject]
 * @param {object} [deps.cryptoImpl]
 * @returns {Promise<object>} Frozen identity: workspaceId, documentId,
 *   instanceId, displayName, identityKind ('custom-xml' | 'fork' | 'url' |
 *   'ephemeral'), documentEphemeral, workspaceEphemeral, identityForked
 */
export async function resolveDocumentIdentity({ office = globalThis.Office, locationObject = globalThis.location, cryptoImpl = globalThis.crypto } = {}) {
    const instanceId = newId('instance');
    const origin = locationObject?.origin || '';
    const url = documentUrl(office);
    let workspaceHash;
    let urlHash;
    try {
        workspaceHash = origin ? await digest(origin, cryptoImpl) : null;
        urlHash = url ? await digest(url, cryptoImpl) : null;
    } catch {
        workspaceHash = null;
        urlHash = null;
    }
    const workspaceId = workspaceHash ? `workspace-${workspaceHash}` : newId('workspace');
    const workspaceEphemeral = !workspaceHash;

    let displayName = 'Unsaved or unidentified document';
    if (url) {
        try { displayName = decodeURIComponent(url.split(/[\\/]/).pop().split(/[?#]/)[0]) || 'Open document'; }
        catch { displayName = 'Open document'; }
    }

    let documentId = null;
    let identityKind = 'ephemeral';
    const parts = customXmlParts(office);
    if (parts) {
        const stored = await readIdentityPart(parts);
        if (stored) {
            if (!urlHash || stored.urlHash === urlHash) {
                documentId = stored.documentId;
                identityKind = 'custom-xml';
            } else {
                // Save As / Copy / moved: derive a deterministic fork id so the
                // original keeps its identity and this copy is stable across
                // its own reloads — without writing to the file at startup.
                const forkHash = await digest(`${stored.documentId}:${url}`, cryptoImpl).catch(() => null);
                documentId = forkHash ? `document-${forkHash}` : randomDocumentId(cryptoImpl);
                identityKind = 'fork';
            }
        }
    }
    if (!documentId && urlHash) {
        documentId = `document-${urlHash}`;
        identityKind = 'url';
    }
    if (!documentId) documentId = randomDocumentId(cryptoImpl);

    return Object.freeze({
        workspaceId,
        documentId,
        instanceId,
        displayName: displayName.slice(0, 160),
        identityKind,
        documentEphemeral: identityKind === 'ephemeral',
        workspaceEphemeral,
        identityForked: identityKind === 'fork',
    });
}

/**
 * Lazily persists the document identity into the file so future loads reuse
 * it and forks can be detected. Called after the document actually joins a
 * cross-document workspace — never at startup — because writing a custom XML
 * part marks the document modified.
 *
 * Best effort: resolves false without throwing when the store is unavailable,
 * the document is unsaved/read-only, or the identity is ephemeral or already
 * persisted.
 *
 * @param {object} identity - Result of resolveDocumentIdentity
 * @param {object} [deps]
 * @returns {Promise<boolean>} True when the identity is now (or already was)
 *   durable in the file
 */
export async function persistDocumentIdentity(identity, { office = globalThis.Office, cryptoImpl = globalThis.crypto } = {}) {
    if (!identity || identity.documentEphemeral || identity.identityKind === 'custom-xml') return !!identity && identity.identityKind === 'custom-xml';
    const parts = customXmlParts(office);
    const url = documentUrl(office);
    if (!parts || !url) return false;
    let urlHash;
    try { urlHash = await digest(url, cryptoImpl); } catch { urlHash = null; }
    if (!urlHash) return false;
    const xml = identityXml(identity.documentId, urlHash, Date.now());
    if (identity.identityKind === 'url') {
        const added = await callAsync(parts, 'addAsync', xml);
        return succeeded(added);
    }
    if (identity.identityKind === 'fork') {
        const stored = await readIdentityPart(parts);
        if (!stored?.part) return false;
        const updated = await callAsync(stored.part, 'setXmlAsync', xml);
        return succeeded(updated);
    }
    return false;
}
