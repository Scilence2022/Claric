import { newId } from './message-shape.js';

async function digest(text, cryptoImpl) {
    if (!cryptoImpl?.subtle) return null;
    const result = await cryptoImpl.subtle.digest('SHA-256', new TextEncoder().encode(text));
    return Array.from(new Uint8Array(result), (byte) => byte.toString(16).padStart(2, '0')).join('');
}

export async function resolveDocumentIdentity({ office = globalThis.Office, locationObject = globalThis.location, cryptoImpl = globalThis.crypto } = {}) {
    const instanceId = newId('instance');
    const origin = locationObject?.origin || '';
    let url = '';
    try { url = typeof office?.context?.document?.url === 'string' ? office.context.document.url : ''; } catch { /* Some hosts do not expose a document URL. */ }
    let workspaceHash;
    let documentHash;
    try {
        workspaceHash = origin ? await digest(origin, cryptoImpl) : null;
        documentHash = url ? await digest(url, cryptoImpl) : null;
    } catch {
        workspaceHash = null;
        documentHash = null;
    }
    let displayName = 'Unsaved or unidentified document';
    if (url) {
        try { displayName = decodeURIComponent(url.split(/[\\/]/).pop().split(/[?#]/)[0]) || 'Open document'; }
        catch { displayName = 'Open document'; }
    }
    return Object.freeze({
        workspaceId: workspaceHash ? `workspace-${workspaceHash}` : newId('workspace'),
        documentId: documentHash ? `document-${documentHash}` : newId('document'),
        instanceId,
        displayName: displayName.slice(0, 160),
        identityKind: documentHash ? 'url' : 'ephemeral',
        documentEphemeral: !documentHash,
        workspaceEphemeral: !workspaceHash,
    });
}
