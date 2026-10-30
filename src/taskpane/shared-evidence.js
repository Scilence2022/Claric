/** Explicitly shared, immutable excerpts. No chat history or live Word objects. */
export const SHARED_EVIDENCE_LIMITS = Object.freeze({ records: 24, totalChars: 96000, textChars: 12000, pageChars: 3000, ttlMs: 900000 });
const SCOPES = new Set(['selection', 'document', 'outline', 'cursor']);
const FIELDS = ['workspaceId', 'documentId', 'instanceId'];
const same = (a, b) => a && b && FIELDS.every((field) => a[field] === b[field]);
const clone = (value) => JSON.parse(JSON.stringify(value));
const fail = (code, message) => Object.assign(new Error(message), { code });

export async function evidenceDigest(text, cryptoImpl = globalThis.crypto) {
    if (!cryptoImpl?.subtle) throw fail('DIGEST_UNAVAILABLE', 'Secure evidence fingerprints are unavailable');
    const bytes = new TextEncoder().encode(text);
    const hash = await cryptoImpl.subtle.digest('SHA-256', bytes);
    return Array.from(new Uint8Array(hash), (byte) => byte.toString(16).padStart(2, '0')).join('');
}

export function createSharedEvidenceStore({ identity, clock = Date.now, cryptoImpl = globalThis.crypto } = {}) {
    if (!identity || FIELDS.some((field) => typeof identity[field] !== 'string' || !identity[field])) throw new Error('Evidence requires a complete document identity');
    const owner = Object.freeze(Object.fromEntries(FIELDS.map((field) => [field, identity[field]])));
    const records = new Map();
    let disposed = false;
    function prune() { for (const [id, item] of records) if (item.metadata.expiresAt <= clock()) records.delete(id); }
    function get(artifactId) {
        prune();
        const item = !disposed && records.get(artifactId);
        if (!item) throw fail('EVIDENCE_UNAVAILABLE', 'Shared evidence expired or was withdrawn; ask the source to share it again');
        return item;
    }
    async function publish(snapshot, { title = 'Shared excerpt', ttlMs = SHARED_EVIDENCE_LIMITS.ttlMs } = {}) {
        if (disposed) throw fail('EVIDENCE_UNAVAILABLE', 'Evidence workspace has stopped');
        if (!SCOPES.has(snapshot?.scope) || snapshot.sourceDocumentId !== owner.documentId || snapshot.sourceInstanceId !== owner.instanceId
            || typeof snapshot.snapshotId !== 'string' || !snapshot.snapshotId || typeof snapshot.documentRevision !== 'string' || !snapshot.documentRevision
            || typeof snapshot.text !== 'string' || !snapshot.text.trim() || snapshot.text.length > SHARED_EVIDENCE_LIMITS.textChars) throw fail('INVALID_EVIDENCE', 'Share a non-empty, bounded snapshot from this document');
        if (!Number.isInteger(ttlMs) || ttlMs < 1000 || ttlMs > SHARED_EVIDENCE_LIMITS.ttlMs
            || typeof title !== 'string' || !title.trim() || title.length > 120) throw fail('INVALID_EVIDENCE', 'Invalid evidence title or lifetime');
        const text = snapshot.text;
        const captured = { scope: snapshot.scope, snapshotId: snapshot.snapshotId, documentRevision: snapshot.documentRevision, truncated: snapshot.truncation?.text === true };
        const digest = await evidenceDigest(text, cryptoImpl);
        if (disposed) throw fail('EVIDENCE_UNAVAILABLE', 'Evidence workspace has stopped');
        prune();
        if (records.size >= SHARED_EVIDENCE_LIMITS.records || [...records.values()].reduce((total, item) => total + item.text.length, text.length) > SHARED_EVIDENCE_LIMITS.totalChars) throw fail('EVIDENCE_CAPACITY', 'Withdraw an excerpt before sharing more evidence');
        const metadata = { artifactId: `evidence-${cryptoImpl.randomUUID()}`, digest, title: title.trim(), scope: captured.scope,
            source: owner, sourceSnapshotId: captured.snapshotId, sourceRevision: captured.documentRevision,
            size: new TextEncoder().encode(text).length, textLength: text.length, createdAt: clock(), expiresAt: clock() + ttlMs,
            truncated: captured.truncated };
        records.set(metadata.artifactId, { text, metadata });
        return clone(metadata);
    }
    function catalog() { prune(); return disposed ? [] : [...records.values()].map((item) => clone(item.metadata)); }
    function read({ artifactId, digest, offset = 0, limit = SHARED_EVIDENCE_LIMITS.pageChars, query = '' } = {}) {
        const { metadata, text } = get(artifactId);
        if (digest !== metadata.digest) throw fail('EVIDENCE_VERSION_MISMATCH', 'Evidence fingerprint does not match the advertised version');
        if (!Number.isInteger(offset) || offset < 0 || offset > text.length || !Number.isInteger(limit) || limit < 1 || limit > SHARED_EVIDENCE_LIMITS.pageChars
            || typeof query !== 'string' || query.length > 200) throw fail('INVALID_EVIDENCE_QUERY', 'Invalid evidence page or search query');
        let start = offset;
        if (query) {
            const found = text.toLowerCase().indexOf(query.toLowerCase(), offset);
            if (found < 0) return { ...clone(metadata), text: '', offset, nextOffset: null, matched: false };
            start = Math.max(offset, found - 160);
        }
        if (start > 0 && /[\uDC00-\uDFFF]/.test(text[start])) start--;
        let end = Math.min(start + limit, text.length);
        if (end < text.length && /[\uD800-\uDBFF]/.test(text[end - 1])) end--;
        if (end === start && end < text.length) throw fail('INVALID_EVIDENCE_QUERY', 'Increase the page size to read this character');
        return { ...clone(metadata), text: text.slice(start, end), offset: start, nextOffset: end < text.length ? end : null, matched: true };
    }
    function metadata(artifactId) { return clone(get(artifactId).metadata); }
    function remove(artifactId) { return records.delete(artifactId); }
    function dispose() { disposed = true; records.clear(); }
    return Object.freeze({ publish, catalog, read, metadata, remove, dispose });
}

/** Serve only an explicitly published version, after checking its live source. */
export function createSharedEvidenceReader({ store, agent, identity } = {}) {
    return {
        async readContext(request = {}, options = {}) {
            if (request.scope !== 'shared_evidence') return agent.readContext(request, options);
            if (options.signal?.aborted) throw new DOMException('Evidence read cancelled', 'AbortError');
            if (!Array.isArray(request.artifactIds) || request.artifactIds.length !== 1) throw fail('INVALID_EVIDENCE_QUERY', 'Read one shared artifact at a time');
            const metadata = store.metadata(request.artifactIds[0]);
            if (!same(metadata.source, identity)) throw fail('EVIDENCE_OWNER_MISMATCH', 'Shared evidence belongs to another instance');
            const current = await agent.getDocumentRevision(metadata.scope, options);
            if (current !== metadata.sourceRevision) throw fail('EVIDENCE_SOURCE_CHANGED', 'The source changed; share a fresh excerpt before using this evidence');
            if (options.signal?.aborted) throw new DOMException('Evidence read cancelled', 'AbortError');
            const evidence = store.read({ artifactId: metadata.artifactId, digest: request.digest, offset: request.offset, limit: request.limit, query: request.query });
            return { snapshotId: metadata.sourceSnapshotId, sourceDocumentId: identity.documentId, sourceInstanceId: identity.instanceId,
                scope: 'shared_evidence', documentRevision: metadata.sourceRevision, text: evidence.text, evidence,
                truncation: { text: metadata.truncated || evidence.nextOffset !== null || evidence.offset > 0 } };
        },
    };
}

/** Verify a page against the catalogue before it enters model context. */
export function validateSharedEvidencePage(snapshot, artifact, target, clock = Date.now) {
    const page = snapshot?.evidence;
    if (!page || !same(artifact?.source, target) || !same(page.source, target)
        || !Number.isSafeInteger(artifact.expiresAt) || artifact.expiresAt <= clock() || page.expiresAt !== artifact.expiresAt
        || page.artifactId !== artifact.artifactId || page.digest !== artifact.digest
        || snapshot.sourceDocumentId !== target.documentId || snapshot.sourceInstanceId !== target.instanceId
        || snapshot.documentRevision !== artifact.sourceRevision || page.sourceRevision !== artifact.sourceRevision
        || snapshot.snapshotId !== artifact.sourceSnapshotId || page.sourceSnapshotId !== artifact.sourceSnapshotId
        || typeof page.text !== 'string' || page.text !== snapshot.text || page.text.length > SHARED_EVIDENCE_LIMITS.pageChars
        || !Number.isInteger(page.offset) || page.offset < 0 || page.textLength !== artifact.textLength
        || page.offset + page.text.length > artifact.textLength || page.scope !== artifact.scope
        || (page.matched === false ? page.text !== '' || page.nextOffset !== null
            : page.nextOffset !== (page.offset + page.text.length < artifact.textLength ? page.offset + page.text.length : null))) throw fail('INVALID_EVIDENCE_RESPONSE', 'Shared evidence does not match its source and advertised version');
    return clone(page);
}
