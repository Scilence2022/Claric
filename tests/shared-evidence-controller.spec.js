/** @jest-environment jsdom */
import { createSharedEvidenceController } from '../src/taskpane/shared-evidence-controller.js';
import { createSharedEvidenceStore, createSharedEvidenceReader } from '../src/taskpane/shared-evidence.js';
const { webcrypto } = require('crypto');
const { TextEncoder } = require('util');
const fs = require('fs');
const path = require('path');

const local = { workspaceId: 'w', documentId: 'a', instanceId: 'a-1' };
const remote = { workspaceId: 'w', documentId: 'b', instanceId: 'b-1', title: 'Reference', expiresAt: 60000 };
let listeners;
let workspace;
beforeEach(() => {
    global.TextEncoder = TextEncoder;
    document.body.innerHTML = fs.readFileSync(path.join(__dirname, '../src/taskpane/taskpane.html'), 'utf8').split('<body>')[1].split('</body>')[0];
    listeners = new AbortController();
});
afterEach(() => { listeners.abort(); workspace?.dispose(); });
const flush = async () => { for (let i = 0; i < 30; i++) await Promise.resolve(); await new Promise((resolve) => setTimeout(resolve, 10)); };
const snap = (identity, text) => ({ sourceDocumentId: identity.documentId, sourceInstanceId: identity.instanceId, scope: 'selection', snapshotId: 'snap', documentRevision: 'r1', text });

test('sharing is explicit, metadata-only, and withdrawing revokes the reader even on server failure', async () => {
    const store = createSharedEvidenceStore({ identity: local, cryptoImpl: webcrypto, clock: () => 1000 });
    const agent = { readContext: jest.fn(async () => snap(local, 'Private selected excerpt')) };
    const client = { identity: local, publishEnvelope: jest.fn(async () => ({})) };
    workspace = createSharedEvidenceController({ client, agent, store, signal: listeners.signal, getDocuments: () => new Map(), clock: () => 1000 });
    expect(agent.readContext).not.toHaveBeenCalled();
    document.getElementById('sharedEvidenceShareBtn').click(); await flush();
    expect(store.catalog()).toHaveLength(1);
    expect(client.publishEnvelope).toHaveBeenCalledWith(expect.objectContaining({ type: 'artifact.announced', target: local,
        payload: expect.objectContaining({ shared: true, textLength: 24 }) }));
    expect(JSON.stringify(client.publishEnvelope.mock.calls)).not.toContain('Private selected excerpt');
    const meta = store.catalog()[0];
    const select = document.getElementById('sharedEvidenceLocal');
    select.value = meta.artifactId; select.dispatchEvent(new Event('change'));
    client.publishEnvelope.mockRejectedValue(new Error('Offline'));
    document.getElementById('sharedEvidenceWithdrawBtn').click(); await flush();
    expect(() => store.metadata(meta.artifactId)).toThrow('withdrawn');
    expect(document.getElementById('sharedEvidenceStatus').textContent).toContain('Local access revoked');
});

test('reads pages with provenance, rechecks them for planning and rejects a changed source', async () => {
    const remoteStore = createSharedEvidenceStore({ identity: remote, cryptoImpl: webcrypto, clock: () => 1000 });
    const meta = await remoteStore.publish(snap(remote, 'x'.repeat(3100) + ' evidence'));
    const source = { getDocumentRevision: jest.fn(async () => 'r1') };
    const reader = createSharedEvidenceReader({ store: remoteStore, agent: source, identity: remote });
    const contextManager = { requestContext: jest.fn(async (_target, request, options) => ({ snapshot: await reader.readContext(request, options) })) };
    const store = createSharedEvidenceStore({ identity: local, cryptoImpl: webcrypto, clock: () => 1000 });
    const addNote = jest.fn();
    workspace = createSharedEvidenceController({ client: { identity: local }, store, contextManager, addNote,
        signal: listeners.signal, getDocuments: () => new Map([[remote.instanceId, remote]]), clock: () => 1000 });
    workspace.onSnapshot({ artifacts: [{ ...meta, summary: 'Excerpt', shared: true }] });
    const select = document.getElementById('sharedEvidenceRemote');
    select.value = meta.artifactId; select.dispatchEvent(new Event('change'));
    document.getElementById('sharedEvidenceReadBtn').click(); await flush();
    expect(addNote).toHaveBeenCalledWith(expect.stringContaining('characters 0–3000'));
    document.getElementById('sharedEvidenceReadBtn').click(); await flush();
    expect(addNote).toHaveBeenCalledWith(expect.stringContaining('characters 3000–3109'));
    const planned = await workspace.getPlanEvidence();
    expect(planned).toHaveLength(2);
    source.getDocumentRevision.mockResolvedValue('r2');
    await expect(workspace.getPlanEvidence()).rejects.toThrow('source changed');
    workspace.onSnapshot({ artifacts: [] });
    await expect(workspace.getPlanEvidence(planned)).rejects.toThrow('withdrawn or expired');
    expect(await workspace.getPlanEvidence()).toEqual([]);
    workspace.dispose();
    expect(document.getElementById('sharedEvidenceReadBtn').disabled).toBe(true);
    remoteStore.dispose();
});
