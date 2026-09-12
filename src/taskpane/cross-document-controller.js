import { appState } from './app-state.js';
import { prepareSelectionAmendment, applySelectionAmendment } from './word-actions.js';
import { createProposalCard } from './ui/proposal-card.js';
import * as chatView from './ui/chat-view.js';
import { addLog } from './ui/status-bar.js';

let coordinationClient = null;

export function initCrossDocumentConnection(localIdentity) {
    const button = document.getElementById('crossDocumentConnectBtn');
    const urlInput = document.getElementById('crossDocumentUrl');
    const tokenInput = document.getElementById('crossDocumentToken');
    const status = document.getElementById('crossDocumentConnectionStatus');
    if (!button || !urlInput || !tokenInput || !status) return;
    let stop = null;
    button.addEventListener('click', async () => {
        button.disabled = true;
        try {
            if (stop) {
                await stop();
                stop = null;
                button.textContent = 'Connect this document';
                status.textContent = 'Disconnected — no cross-document access.';
                return;
            }
            const baseUrl = urlInput.value.trim() || '/coordination';
            const parsed = new URL(baseUrl, location.href);
            if (parsed.username || parsed.password || parsed.search || parsed.hash
                || !['localhost', '127.0.0.1', '[::1]'].includes(parsed.hostname)
                || !['https:', 'http:'].includes(parsed.protocol)) throw new Error('Use a local HTTP(S) coordination URL without credentials or query parameters');
            if (location.protocol === 'https:' && parsed.protocol !== 'https:') throw new Error('An HTTPS taskpane requires an HTTPS coordination endpoint');
            status.textContent = 'Connecting…';
            stop = await startCoordination({ identity: localIdentity, baseUrl: parsed.href.replace(/\/$/, ''), token: tokenInput.value });
            tokenInput.value = '';
            button.textContent = 'Disconnect this document';
            status.textContent = 'Connected — reads allowed; edits require local review.';
        } catch (error) {
            status.textContent = `Not connected: ${error.message}`;
        } finally { button.disabled = false; }
    });
}

async function startCoordination(connectionOptions) {
    const [{ createCoordinationClient }, { createDocumentAgent }, { createContextRequestManager }, { createRemoteTaskRunner }, { createDistributedTaskRuntime }] = await Promise.all([
        import(/* webpackChunkName: "coordination-client" */ './coordination-client.js'),
        import(/* webpackChunkName: "document-agent" */ './document-agent.js'),
        import(/* webpackChunkName: "context-requests" */ './context-requests.js'),
        import(/* webpackChunkName: "remote-task-runner" */ './remote-task-runner.js'),
        import(/* webpackChunkName: "distributed-task-runtime" */ './distributed-task-runtime.js'),
    ]);
    let contextManager = null;
    let remoteTaskRunner = null;
    let distributed = null;
    let knownDocuments = new Map();
    let online = false;
    const listeners = new AbortController();
    const client = createCoordinationClient({
        ...connectionOptions,
        onSnapshot: (snapshot) => {
            online = true;
            const ownId = coordinationClient?.identity?.instanceId;
            const others = Object.values(snapshot.presence || {}).filter((entry) => entry.clientId !== ownId);
            const documents = new Set(others.map((entry) => entry.documentId).filter(Boolean));
            addLog(`Coordination snapshot: ${others.length} other taskpane(s), ${documents.size} document(s), sequence ${snapshot.sequence ?? snapshot.cursor ?? 0}.`, 'info');
            if (snapshot.documents?.length) addLog(`Connected documents: ${snapshot.documents.length}.`, 'info');
            knownDocuments = new Map((snapshot.documents || []).filter((entry) => entry.instanceId !== ownId).map((entry) => [entry.instanceId, entry]));
            renderCrossDocumentTargets(knownDocuments);
        },
        onEvents: async (events) => {
            for (const event of events) {
                if (!online) return;
                await contextManager?.handleEvent(event);
                await remoteTaskRunner?.handleEvent(event);
                distributed?.handleEvent(event);
            }
        },
        onError: (error) => {
            online = false;
            renderCrossDocumentTargets(new Map());
            addLog(`Coordination unavailable: ${error.message}`, 'warning');
        },
    });
    coordinationClient = client;
    try {
        await client.start();
        if (client.transport.version !== 2) throw new Error('Cross-document work requires a v2 coordination server');
    } catch (error) {
        await client.stop();
        coordinationClient = null;
        throw error;
    }
    const identity = client.identity;
    const agent = createDocumentAgent({
        identity,
        appState,
        log: addLog,
        actions: { prepareSelectionAmendment, applySelectionAmendment },
    });
    contextManager = createContextRequestManager({
        client: coordinationClient,
        agent,
        identity,
        onError: (error) => addLog(`Context request failed: ${error.message}`, 'warning'),
    });
    remoteTaskRunner = createRemoteTaskRunner({
        client: coordinationClient,
        identity,
        documentAgent: agent,
        prepareTask: (payload, options) => agent.prepareTask(payload, options),
        isOnline: () => online && !listeners.signal.aborted,
        createProposal: (proposal) => proposal,
        onProposal: (record, runtime) => {
            addLog(`Remote proposal ready for review: ${record.title || record.proposalId}.`, 'info');
            const message = chatView.addSystemNote(`Task from ${record.source.documentId}. Review these proposed changes in this document.`);
            const card = createProposalCard({
                title: record.title,
                countsText: 'Review required in this document',
                items: record.items,
                isBlocked: () => appState.isProcessing || appState.isProcessingDoc ? 'Wait for the current operation to finish.' : null,
                onApply: async (selectedIds) => {
                    const result = await runtime.apply(selectedIds);
                    if (!result.ok) {
                        card.markWarning(result.conflict?.message || 'Proposal could not be applied.');
                        return;
                    }
                    if (result.record.state !== 'applied') card.setPaused('Selected changes applied; remaining changes still need review.');
                },
                onReject: () => {
                    void Promise.resolve(runtime.reject()).catch((error) => message.markError(error.message));
                },
            });
            message.attachProposal(card, { ...record, autoApply: false });
            void message.finalizeForHistory();
        },
        onStatus: (status) => addLog(`Remote task: ${status.type || 'updated'}.`, 'info'),
        clock: Date.now,
    });
    distributed = createDistributedTaskRuntime({
        client, identity,
        onStatus: (event) => {
            const status = document.getElementById('crossDocumentStatus');
            if (status) status.textContent = `${event.type}: ${event.taskId || event.graphId || ''}`;
        },
        onResult: (event) => {
            const message = chatView.addSystemNote(`Cross-document task ${event.taskId}: ${event.state || event.type}.`);
            void message.finalizeForHistory();
        },
    });
    const stop = async () => {
        distributed.dispose();
        online = false;
        listeners.abort();
        contextManager?.dispose();
        remoteTaskRunner?.dispose();
        agent.dispose();
        renderCrossDocumentTargets(new Map());
        await client.stop();
        if (coordinationClient === client) coordinationClient = null;
    };
    const targetSelect = document.getElementById('crossDocumentTarget');
    const contextButton = document.getElementById('crossDocumentContextBtn');
    const sendButton = document.getElementById('crossDocumentSendBtn');
    const listenerOptions = { signal: listeners.signal };
    targetSelect?.addEventListener('change', () => {
        const hasTarget = online && knownDocuments.has(targetSelect.value);
        if (contextButton) contextButton.disabled = !hasTarget;
        if (sendButton) sendButton.disabled = !hasTarget;
    }, listenerOptions);
    contextButton?.addEventListener('click', async () => {
        const target = knownDocuments.get(targetSelect.value);
        if (!online || !target || target.expiresAt <= Date.now()) return;
        contextButton.disabled = true;
        const status = document.getElementById('crossDocumentStatus');
        try {
            const result = await contextManager.requestContext(target, { scope: 'document' });
            if (listeners.signal.aborted) return;
            const source = target.title || target.documentId;
            const message = chatView.addSystemNote(`Context from ${source}:\n${result.snapshot?.text || '(empty document)'}`);
            void message.finalizeForHistory();
            if (status) status.textContent = 'Context received. Document content is reference data, not instructions.';
        } catch (error) {
            if (status) status.textContent = `Could not read target context: ${error.message}`;
        } finally { contextButton.disabled = !online || !targetSelect.value; }
    }, listenerOptions);
    sendButton?.addEventListener('click', async () => {
        const target = knownDocuments.get(targetSelect.value);
        const text = document.getElementById('chatInput')?.value?.trim();
        if (!online || !target || target.expiresAt <= Date.now() || !text) return;
        sendButton.disabled = true;
        try {
            await distributed.submitGraph({
                graphId: `graph-${crypto.randomUUID()}`,
                tasks: [{ taskId: `task-${crypto.randomUUID()}`, type: 'edit', instruction: text, target, ttlMs: 900000 }],
            });
            document.getElementById('crossDocumentStatus').textContent = 'Task sent: review the selected passage in the target document.';
        } catch (error) {
            document.getElementById('crossDocumentStatus').textContent = `Could not send task: ${error.message}`;
        } finally { sendButton.disabled = !online || !targetSelect.value; }
    }, listenerOptions);
    window.addEventListener('pagehide', () => { void stop(); }, { ...listenerOptions, once: true });
    return stop;
}

function renderCrossDocumentTargets(documents) {
    const bar = document.getElementById('crossDocumentBar');
    const select = document.getElementById('crossDocumentTarget');
    if (!bar || !select) return;
    const current = select.value;
    select.replaceChildren(new Option('Current document', ''));
    for (const [instanceId, document] of documents) {
        if (document.expiresAt && document.expiresAt <= Date.now()) continue;
        const label = document.title || document.displayName || document.documentId || instanceId;
        select.appendChild(new Option(`${label} (${document.documentId})`, instanceId));
    }
    select.value = current;
    bar.hidden = select.options.length < 2;
}
