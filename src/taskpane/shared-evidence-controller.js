import { validateSharedEvidencePage } from './shared-evidence.js';

/** Small UI adapter over the existing authenticated coordination transport. */
export function createSharedEvidenceController({ client, agent, store, contextManager, getDocuments, signal,
    addNote = () => {}, clock = Date.now } = {}) {
    const identity = client.identity;
    const share = document.getElementById('sharedEvidenceShareBtn');
    const withdraw = document.getElementById('sharedEvidenceWithdrawBtn');
    const local = document.getElementById('sharedEvidenceLocal');
    const available = document.getElementById('sharedEvidenceRemote');
    const read = document.getElementById('sharedEvidenceReadBtn');
    const status = document.getElementById('sharedEvidenceStatus');
    const pages = new Map();
    let artifacts = new Map();
    let disposed = false;
    let reading = false;
    let sharing = false;
    const active = () => !disposed && !signal.aborted;
    const say = (text) => { if (active() && status) status.textContent = text; };
    const keyFor = (id, offset) => `${id}:${offset}`;
    function render() {
        if (!active()) return;
        if (share) share.disabled = sharing;
        if (local) {
            const selected = local.value;
            local.replaceChildren(new Option('Your shared excerpts', ''));
            for (const item of store.catalog()) local.add(new Option(item.title, item.artifactId));
            local.value = selected;
            if (withdraw) withdraw.disabled = !local.value || sharing;
        }
        if (available) {
            const selected = available.value;
            available.replaceChildren(new Option('Shared evidence from linked documents', ''));
            for (const item of artifacts.values()) {
                if (item.source.instanceId === identity.instanceId) continue;
                const source = getDocuments().get(item.source.instanceId);
                if (!source || source.expiresAt <= clock()) continue;
                available.add(new Option(`${source.title || source.documentId}: ${item.summary || 'Shared excerpt'}`, item.artifactId));
            }
            available.value = selected;
            if (read) read.disabled = !available.value || reading;
        }
    }
    function onSnapshot(snapshot) {
        artifacts = new Map((snapshot.artifacts || []).filter((item) => item.shared === true && item.expiresAt > clock()
            && item.source?.workspaceId === identity.workspaceId).map((item) => [item.artifactId, item]));
        for (const [key, page] of pages) {
            const item = artifacts.get(page.artifactId);
            const source = item && getDocuments().get(item.source.instanceId);
            if (!item || !source || source.expiresAt <= clock() || item.digest !== page.digest) pages.delete(key);
        }
        render();
    }
    async function fetchPage(artifact, { offset = 0, query = '' } = {}) {
        const target = getDocuments().get(artifact.source.instanceId);
        if (!active() || !target || target.expiresAt <= clock()) throw new Error('The evidence source is no longer connected');
        const result = await contextManager.requestContext(target, { scope: 'shared_evidence', artifactIds: [artifact.artifactId],
            digest: artifact.digest, offset, limit: 3000, ...(query ? { query } : {}) }, { signal });
        if (!active()) throw new Error('Evidence workspace disconnected');
        return validateSharedEvidencePage(result.snapshot, artifact, target, clock);
    }
    share?.addEventListener('click', async () => {
        if (!active() || sharing) return;
        sharing = true; render();
        let metadata;
        try {
            const scope = document.getElementById('sharedEvidenceScope')?.value || 'selection';
            const snapshot = await agent.readContext({ scope }, { signal });
            if (!active()) return;
            metadata = await store.publish(snapshot, { title: scope === 'selection' ? 'Selected passage' : 'Document excerpt' });
            if (!active()) return;
            await client.publishEnvelope({ type: 'artifact.announced', target: identity, correlationId: metadata.artifactId,
                ttlMs: metadata.expiresAt - metadata.createdAt, createdAt: metadata.createdAt,
                payload: { artifactId: metadata.artifactId, digest: metadata.digest, size: metadata.size, mediaType: 'text/plain', shared: true,
                    scope: metadata.scope, summary: metadata.title, sourceRevision: metadata.sourceRevision,
                    sourceSnapshotId: metadata.sourceSnapshotId, textLength: metadata.textLength, truncated: metadata.truncated } });
            say(`Shared ${metadata.textLength} characters${metadata.truncated ? ' (excerpt only)' : ''}. Expires in 15 minutes; source changes require sharing again.`);
        } catch (error) {
            if (metadata) store.remove(metadata.artifactId);
            say(`Could not share evidence: ${error.message}`);
        } finally { sharing = false; render(); }
    }, { signal });
    withdraw?.addEventListener('click', async () => {
        if (!active() || sharing || !local?.value) return;
        const id = local.value;
        // Revoke the local reader immediately, even if the server is unreachable.
        store.remove(id); render();
        try {
            await client.publishEnvelope({ type: 'artifact.removed', target: identity, correlationId: id, payload: { artifactId: id } });
            say('Excerpt withdrawn. Previously read copies may remain in other documents’ chats.');
        } catch (error) { say(`Local access revoked; directory removal failed: ${error.message}`); }
    }, { signal });
    local?.addEventListener('change', render, { signal });
    available?.addEventListener('change', render, { signal });
    read?.addEventListener('click', async () => {
        const artifact = artifacts.get(available.value);
        if (!active() || reading || !artifact) return;
        reading = true; render();
        try {
            const query = document.getElementById('sharedEvidenceQuery')?.value?.trim() || '';
            const previous = [...pages.values()].filter((page) => page.artifactId === artifact.artifactId).at(-1);
            if (!query && previous?.nextOffset === null) { say('The shared excerpt has been read completely.'); return; }
            const page = await fetchPage(artifact, { offset: query ? 0 : previous?.nextOffset || 0, query });
            if (!page.matched) { say('No matching text in this shared excerpt.'); return; }
            if (pages.size >= 4 && !pages.has(keyFor(page.artifactId, page.offset))) throw new Error('Four evidence pages are already loaded; reconnect to clear them');
            pages.set(keyFor(page.artifactId, page.offset), page);
            addNote(`Shared evidence ${page.artifactId} from ${artifact.source.documentId}, characters ${page.offset}–${page.offset + page.text.length}:\n${page.text}`);
            say(page.nextOffset === null ? 'Excerpt read. Planning will recheck its source version.' : 'Page read. Read again for the next page, or search within the excerpt.');
        } catch (error) { say(`Could not read shared evidence: ${error.message}`); }
        finally { reading = false; render(); }
    }, { signal });
    render();
    return {
        onSnapshot,
        async getPlanEvidence(expected = [...pages.values()]) {
            const verified = [];
            for (const old of expected) {
                const key = keyFor(old.artifactId, old.offset);
                const artifact = artifacts.get(old.artifactId);
                if (!artifact) { pages.delete(key); throw new Error('Shared evidence was withdrawn or expired; read a fresh excerpt'); }
                try { verified.push(await fetchPage(artifact, { offset: old.offset })); }
                catch (error) { pages.delete(key); throw error; }
            }
            return verified;
        },
        dispose() {
            disposed = true; pages.clear(); artifacts.clear(); store.dispose();
            for (const control of [share, withdraw, read]) if (control) control.disabled = true;
        },
    };
}
