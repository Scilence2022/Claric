import { normalizeDocumentAccess, allowsDocumentAccess } from '../lib/coordination/access.cjs';

export function createWorkspaceAccess({ identity, storage = globalThis.localStorage } = {}) {
    const key = `claric.workspace.access.${identity.workspaceId}.${identity.documentId}`;
    let access;
    try { access = normalizeDocumentAccess(JSON.parse(storage?.getItem(key) || '{}')); }
    catch { access = normalizeDocumentAccess({ readContext: false, requestEdits: false }); }
    return {
        get: () => normalizeDocumentAccess(access),
        allows: (source, operation) => source?.workspaceId === identity.workspaceId && allowsDocumentAccess(access, source.documentId, operation),
        update(documentId, operation, allowed) {
            const next = normalizeDocumentAccess(access);
            if (documentId) {
                let peer = next.peers.find((item) => item.documentId === documentId);
                if (!peer) { peer = { documentId }; next.peers.push(peer); }
                peer[operation] = allowed;
            } else next[operation] = allowed;
            // Persist before granting. Failed storage leaves the old policy intact.
            const normalized = normalizeDocumentAccess(next);
            storage.setItem(key, JSON.stringify(normalized));
            access = normalized;
            return normalizeDocumentAccess(access);
        },
    };
}
