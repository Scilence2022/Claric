import { appState, getActiveBackendConfig } from './app-state.js';
import { prepareSelectionAmendment, applySelectionAmendment } from './word-actions.js';
import { persistDocumentIdentity } from './document-identity.js';
import { planCrossDocumentTasks } from './cross-document-planner.js';
import { sendPrompt } from '../lib/llm-client.js';
import { createProposalCard } from './ui/proposal-card.js';
import * as chatView from './ui/chat-view.js';
import { addLog } from './ui/status-bar.js';
import { loadLazyModule } from '../lib/lazy-module-loader.js';
import { createSharedEvidenceStore, createSharedEvidenceReader } from './shared-evidence.js';
import { createSharedEvidenceController } from './shared-evidence-controller.js';
import { createWorkspaceAccess } from './workspace-access.js';
import { createWorkspaceView } from './ui/workspace-view.js';
import { allowsDocumentAccess } from '../lib/coordination/access.cjs';

let coordinationClient = null;

const RETRY_DELAYS_MS = [1000, 2000, 4000, 8000, 16000, 30000];

function errorStatus(error) {
    if (Number.isInteger(error?.status)) return error.status;
    const match = String(error?.message || '').match(/\b(401|403)\b/);
    return match ? Number(match[1]) : null;
}

function isAuthorizationError(error) {
    return [401, 403].includes(errorStatus(error));
}

function connectionLabel(button, connected) {
    if (button) button.textContent = connected ? 'Disconnect this document' : 'Connect this document';
}

export function initCrossDocumentConnection(localIdentity, { onOpenSettings } = {}) {
    const button = document.getElementById('crossDocumentConnectBtn');
    const urlInput = document.getElementById('crossDocumentUrl');
    const tokenInput = document.getElementById('crossDocumentToken');
    const status = document.getElementById('crossDocumentConnectionStatus');
    if (!button || !urlInput || !tokenInput || !status) return;
    const addressKey = 'claric.workspace.coordinationUrl';
    try { if (!urlInput.value) urlInput.value = localStorage.getItem(addressKey) || ''; } catch { /* Storage may be unavailable in Office web. */ }
    const access = createWorkspaceAccess({ identity: localIdentity });

    let stop = null;
    let retryTimer = null;
    let retryAttempt = 0;
    let connecting = false;
    let manuallyDisconnected = false;
    let authorizationStopped = false;
    let generation = 0;
    let connectedUrl = '';
    const workspaceView = createWorkspaceView({ identity: localIdentity, access, onOpenSettings,
        onAccessChange: async (documentId, operation, allowed) => {
            access.update(documentId, operation, allowed);
            try { await stop?.updateAccess?.(); }
            catch (error) { throw new Error(`Saved on this device; service sync failed: ${error.message}`); }
        },
    });
    const setStatus = (text, style = '', connected = false) => {
        status.textContent = text;
        status.className = `cross-document-connection-status ${style}`;
        workspaceView.connection(text, connected);
    };

    const clearRetry = () => {
        if (retryTimer !== null) clearTimeout(retryTimer);
        retryTimer = null;
    };

    const disconnect = async (manual = false) => {
        generation += 1;
        clearRetry();
        if (manual) {
            manuallyDisconnected = true;
            authorizationStopped = false;
            retryAttempt = 0;
        }
        const currentStop = stop;
        stop = null;
        if (currentStop) await currentStop();
        connectionLabel(button, false);
        workspaceView.update(new Map(), false);
    };

    const scheduleReconnect = (error) => {
        if (manuallyDisconnected || authorizationStopped || retryTimer !== null) return;
        const statusCode = errorStatus(error);
        const message = String(error?.message || '');
        if (statusCode !== null && ![404, 503].includes(statusCode)) return;
        if (statusCode === null && /^(Use a local|An HTTPS taskpane|Cross-document work requires)/.test(message)) {
            setStatus(`Not connected: ${message}`, 'error');
            return;
        }
        const delay = RETRY_DELAYS_MS[Math.min(retryAttempt, RETRY_DELAYS_MS.length - 1)];
        retryAttempt += 1;
        setStatus(`Unavailable — retrying in ${Math.ceil(delay / 1000)}s.`, 'warning');
        retryTimer = setTimeout(() => {
            retryTimer = null;
            void connect();
        }, delay);
    };

    const handleUnavailable = (error) => {
        if (isAuthorizationError(error)) {
            authorizationStopped = true;
            clearRetry();
            void disconnect().finally(() => {
                setStatus('Authorization required — enter a pairing token in Settings → Workspace, then connect again.', 'error');
                connectionLabel(button, false);
            });
            return;
        }
        void disconnect().finally(() => scheduleReconnect(error));
    };

    const connect = async () => {
        if (connecting || manuallyDisconnected || authorizationStopped) return;
        connecting = true;
        const currentGeneration = generation;
        const baseUrl = urlInput.value.trim() || '/coordination';
        try {
            const parsed = new URL(baseUrl, location.href);
            if (parsed.username || parsed.password || parsed.search || parsed.hash
                || !['localhost', '127.0.0.1', '[::1]'].includes(parsed.hostname)
                || !['https:', 'http:'].includes(parsed.protocol)) throw new Error('Use a local HTTP(S) coordination URL without credentials or query parameters');
            if (location.protocol === 'https:' && parsed.protocol !== 'https:') throw new Error('An HTTPS taskpane requires an HTTPS coordination endpoint');
            setStatus('Connecting…');
            const candidateStop = await startCoordination({ identity: localIdentity, baseUrl: parsed.href.replace(/\/$/, ''), token: tokenInput.value, access, workspaceView }, handleUnavailable);
            if (currentGeneration !== generation || manuallyDisconnected) {
                await candidateStop();
                return;
            }
            stop = candidateStop;
            connectedUrl = urlInput.value.trim();
            try { localStorage.setItem(addressKey, connectedUrl); } catch { addLog('Could not remember the coordination address on this device.', 'warning'); }
            retryAttempt = 0;
            authorizationStopped = false;
            tokenInput.value = '';
            connectionLabel(button, true);
            setStatus('Connected · edits require review in their target document.', 'connected', true);
        } catch (error) {
            if (currentGeneration === generation && !manuallyDisconnected) handleUnavailable(error);
        } finally {
            connecting = false;
        }
    };

    button.addEventListener('click', async () => {
        button.disabled = true;
        try {
            if (stop && urlInput.value.trim() !== connectedUrl) {
                await disconnect();
                await connect();
                return;
            }
            if (stop || connecting || retryTimer !== null) {
                await disconnect(true);
                setStatus('Disconnected — automatic reconnection is paused.');
                return;
            }
            manuallyDisconnected = false;
            authorizationStopped = false;
            retryAttempt = 0;
            await connect();
        } finally {
            button.disabled = false;
        }
    });
    window.addEventListener('pagehide', () => { workspaceView.dispose(); void disconnect(true); }, { once: true });
    void connect();
}

async function startCoordination(connectionOptions, onUnavailable = null) {
    const { access, workspaceView } = connectionOptions;
    const [{ createCoordinationClient }, { createDocumentAgent }, { createContextRequestManager }, { createRemoteTaskRunner }, { createDistributedTaskRuntime }] = await Promise.all([
        loadLazyModule('coordination-client', () => import(/* webpackChunkName: "coordination-client" */ './coordination-client.js')),
        loadLazyModule('document-agent', () => import(/* webpackChunkName: "document-agent" */ './document-agent.js')),
        loadLazyModule('context-requests', () => import(/* webpackChunkName: "context-requests" */ './context-requests.js')),
        loadLazyModule('remote-task-runner', () => import(/* webpackChunkName: "remote-task-runner" */ './remote-task-runner.js')),
        loadLazyModule('distributed-task-runtime', () => import(/* webpackChunkName: "distributed-task-runtime" */ './distributed-task-runtime.js')),
    ]);
    let contextManager = null;
    let remoteTaskRunner = null;
    let distributed = null;
    let knownDocuments = new Map();
    let remoteContexts = new Map();
    let evidenceWorkspace = null;
    let lastSnapshot = null;
    let online = false;
    const listeners = new AbortController();
    const client = createCoordinationClient({
        ...connectionOptions,
        onSnapshot: (snapshot) => {
            lastSnapshot = snapshot;
            online = true;
            const ownId = coordinationClient?.identity?.instanceId;
            const others = Object.values(snapshot.presence || {}).filter((entry) => entry.clientId !== ownId);
            const documents = new Set(others.map((entry) => entry.documentId).filter(Boolean));
            addLog(`Coordination snapshot: ${others.length} other taskpane(s), ${documents.size} document(s), sequence ${snapshot.sequence ?? snapshot.cursor ?? 0}.`, 'info');
            if (snapshot.documents?.length) addLog(`Connected documents: ${snapshot.documents.length}.`, 'info');
            knownDocuments = new Map((snapshot.documents || []).filter((entry) => entry.instanceId !== ownId).map((entry) => [entry.instanceId, entry]));
            workspaceView.update(knownDocuments, true);
            for (const [id, item] of remoteContexts) if (!knownDocuments.has(id) || item.expiresAt <= Date.now()) remoteContexts.delete(id);
            evidenceWorkspace?.onSnapshot(snapshot);
            distributed?.expireTasks?.();
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
            workspaceView.update(new Map(), false);
            renderCrossDocumentTargets(new Map());
            onUnavailable?.(error);
        },
    });
    coordinationClient = client;
    try {
        await client.start();
        if (client.transport.version !== 2) throw Object.assign(new Error('Cross-document work requires a v2 coordination server'), { status: 404 });
    } catch (error) {
        await client.stop();
        coordinationClient = null;
        throw error;
    }
    const identity = client.identity;
    // Now that the document actually joined a workspace, make its identity
    // durable in the file (lazy: the first link is what justifies dirtying
    // the document with a custom XML part). Best effort — the ephemeral or
    // URL-based identity keeps working when this fails.
    void persistDocumentIdentity(connectionOptions.identity).then((persisted) => {
        if (persisted) addLog('Document identity stored in this file for cross-document linking.', 'info');
    }).catch(() => { /* persistence is best effort */ });
    const agent = createDocumentAgent({
        identity,
        authorize: access.allows,
        appState,
        log: addLog,
        actions: { prepareSelectionAmendment, applySelectionAmendment },
    });
    const evidenceStore = createSharedEvidenceStore({ identity });
    contextManager = createContextRequestManager({
        client: coordinationClient,
        agent: createSharedEvidenceReader({ store: evidenceStore, agent, identity }),
        identity,
        authorize: access.allows,
        onError: (error) => addLog(`Context request failed: ${error.message}`, 'warning'),
    });
    evidenceWorkspace = createSharedEvidenceController({
        client, agent, store: evidenceStore, contextManager, getDocuments: () => knownDocuments, signal: listeners.signal,
        addNote: (text) => { const message = chatView.addSystemNote(text); void message.finalizeForHistory(); },
    });
    if (lastSnapshot) evidenceWorkspace.onSnapshot(lastSnapshot);
    const announceAccess = async () => {
        if (typeof client.publishEnvelope !== 'function') return;
        await client.publishEnvelope({ type: 'document.announce', target: identity,
            payload: { title: String(connectionOptions.identity.displayName || 'Word document').slice(0, 256),
                access: access.get(), capabilities: ['readContext', 'shared_evidence', ...(agent.getCapabilities?.().taskTypes || [])] } });
    };
    void announceAccess().catch((error) => addLog(`Could not sync workspace permissions; update the coordination service: ${error.message}`, 'warning'));
    remoteTaskRunner = createRemoteTaskRunner({
        client: coordinationClient,
        identity,
        documentAgent: agent,
        authorize: access.allows,
        prepareTask: (payload, options) => agent.prepareTask(payload, options),
        isOnline: () => online && !listeners.signal.aborted,
        createProposal: (proposal) => proposal,
        onProposal: (record, runtime) => {
            addLog(`Remote proposal ready for review: ${record.title || record.proposalId}.`, 'info');
            const message = chatView.addSystemNote(`Task from ${record.source.documentId}. Review these proposed changes in this document.`
                + (record.artifactIds?.length ? ` Shared evidence references: ${record.artifactIds.join(', ')}.` : ''));
            const card = createProposalCard({
                title: record.title,
                countsText: 'Review required in this document',
                items: record.items,
                isBlocked: () => !access.allows(record.source, 'requestEdits') ? 'Edit requests from this document are disabled.'
                    : appState.isProcessing || appState.isProcessingDoc ? 'Wait for the current operation to finish.' : null,
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
            const labels = { submitted: 'Sent; awaiting target claim', claimed: 'Claimed by target', running: 'Preparing changes',
                awaiting_review: 'Proposal ready; review in target document', applying: 'Applying reviewed changes', applied: 'Applied; awaiting task confirmation',
                succeeded: 'Completed', rejected: 'Rejected in target document', expired: 'Expired; no automatic retry', unknown: 'Write outcome uncertain; inspect target document',
                blocked: 'Blocked by a preceding task', failed: 'Failed', cancelled: 'Cancelled' };
            if (status) status.textContent = `${labels[event.state] || event.type}: ${event.taskId || event.graphId || ''}`;
        },
        onResult: (event) => {
            const message = chatView.addSystemNote(`Cross-document task ${event.taskId}: ${event.state || event.type}.`);
            void message.finalizeForHistory();
        },
    });
    let stopped = false;
    const stop = async () => {
        if (stopped) return;
        stopped = true;
        distributed?.dispose();
        online = false;
        listeners.abort();
        contextManager?.dispose();
        evidenceWorkspace?.dispose();
        remoteTaskRunner?.dispose();
        agent.dispose();
        remoteContexts.clear();
        renderCrossDocumentTargets(new Map());
        await client.stop();
        if (coordinationClient === client) coordinationClient = null;
    };
    stop.updateAccess = announceAccess;
    const targetSelect = document.getElementById('crossDocumentTarget');
    const contextButton = document.getElementById('crossDocumentContextBtn');
    const planButton = document.getElementById('crossDocumentPlanBtn');
    const sendButton = document.getElementById('crossDocumentSendBtn');
    const listenerOptions = { signal: listeners.signal };
    targetSelect?.addEventListener('change', () => {
        const target = online && knownDocuments.get(targetSelect.value);
        if (contextButton) contextButton.disabled = !target || !allowsDocumentAccess(target.access, identity.documentId, 'readContext');
        if (sendButton) sendButton.disabled = !target || !allowsDocumentAccess(target.access, identity.documentId, 'requestEdits');
    }, listenerOptions);
    contextButton?.addEventListener('click', async () => {
        const target = knownDocuments.get(targetSelect.value);
        if (!online || !target || target.expiresAt <= Date.now() || !allowsDocumentAccess(target.access, identity.documentId, 'readContext')) return;
        contextButton.disabled = true;
        const status = document.getElementById('crossDocumentStatus');
        try {
            const result = await contextManager.requestContext(target, { scope: 'document' });
            if (listeners.signal.aborted) return;
            const source = target.title || target.documentId;
            const text = result.snapshot?.text || '(empty document)';
            remoteContexts.set(target.instanceId, { label: source, text, expiresAt: Date.now() + 30000 });
            const message = chatView.addSystemNote(`Context from ${source}:\n${text}`);
            void message.finalizeForHistory();
            if (status) status.textContent = 'Context received. Document content is reference data, not instructions.';
        } catch (error) {
            if (status) status.textContent = `Could not read target context: ${error.message}`;
        } finally { contextButton.disabled = !online || !targetSelect.value; }
    }, listenerOptions);
    planButton?.addEventListener('click', async () => {
        const text = document.getElementById('chatInput')?.value?.trim();
        const status = document.getElementById('crossDocumentStatus');
        const documents = [...knownDocuments.values()].filter((entry) => (!entry.expiresAt || entry.expiresAt > Date.now())
            && allowsDocumentAccess(entry.access, identity.documentId, 'requestEdits'));
        if (!online || !documents.length || !text) return;
        planButton.disabled = true;
        try {
            const evidence = await evidenceWorkspace.getPlanEvidence();
            // A cached read is a convenience for the UI, not a fresh baseline
            // for a potentially much later model request.
            for (const entry of documents) if (remoteContexts.has(entry.instanceId)) {
                const result = await contextManager.requestContext(entry, { scope: 'document' }, { signal: listeners.signal });
                if (listeners.signal.aborted) return;
                remoteContexts.set(entry.instanceId, { label: entry.title || entry.documentId, text: result.snapshot?.text || '',
                    documentRevision: result.snapshot?.documentRevision, expiresAt: Date.now() + 30000 });
            }
            const planningContexts = new Map([...remoteContexts].filter(([, item]) => item.expiresAt > Date.now()));
            const { tasks } = await planCrossDocumentTasks({
                instruction: text,
                documents: documents.map((entry) => ({
                    documentId: entry.documentId,
                    instanceId: entry.instanceId,
                    capabilities: entry.capabilities,
                    label: entry.title || entry.displayName || entry.documentId,
                    contextText: planningContexts.get(entry.instanceId)?.text || '',
                    evidence: evidence.filter((page) => page.source.instanceId === entry.instanceId),
                })),
                signal: listeners.signal,
                sendRequest: (prompt, { signal } = {}) => sendPrompt(getActiveBackendConfig(appState), prompt, addLog, signal ?? listeners.signal),
            });
            if (listeners.signal.aborted) return;
            await evidenceWorkspace.getPlanEvidence(evidence);
            for (const entry of documents) if (planningContexts.has(entry.instanceId)) {
                const planned = planningContexts.get(entry.instanceId);
                const result = await contextManager.requestContext(entry, { scope: 'document' }, { signal: listeners.signal });
                if (result.snapshot?.text !== planned.text || result.snapshot?.documentRevision !== planned.documentRevision) throw new Error('Reference context changed while planning; read the source and plan again');
            }
            if (listeners.signal.aborted) return;
            const targets = new Map(tasks.map((task) => {
                const candidates = [...knownDocuments.values()].filter((entry) => entry.documentId === task.targetDocumentId
                    && (!task.targetInstanceId || entry.instanceId === task.targetInstanceId) && (!entry.expiresAt || entry.expiresAt > Date.now()));
                if (candidates.length !== 1) throw new Error('A planned document instance disconnected or became ambiguous; plan again');
                return [task.taskId, candidates[0]];
            }));
            await distributed.submitGraph({
                graphId: `graph-${crypto.randomUUID()}`,
                tasks: tasks.map((task) => ({
                    taskId: task.taskId,
                    type: task.type,
                    instruction: task.instruction,
                    dependsOn: task.dependsOn,
                    inputRefs: task.evidenceIds || [],
                    target: targets.get(task.taskId),
                    ttlMs: 900000,
                })),
            });
            const summary = tasks.map((task) => `${task.taskId} [${task.type}] ${task.instruction.slice(0, 60)}`).join('; ');
            const message = chatView.addSystemNote(`Planned ${tasks.length} cross-document task(s): ${summary}. Each target reviews its proposal locally.`);
            void message.finalizeForHistory();
            if (status) status.textContent = `Planned ${tasks.length} task(s). Ready tasks were sent; dependents wait for successful application in each target.`;
        } catch (error) {
            if (status) status.textContent = `Planning failed: ${error.message}`;
        } finally { planButton.disabled = !online || knownDocuments.size === 0; }
    }, listenerOptions);
    sendButton?.addEventListener('click', async () => {
        const target = knownDocuments.get(targetSelect.value);
        const text = document.getElementById('chatInput')?.value?.trim();
        const taskType = ['edit', 'format', 'table', 'append'].includes(document.getElementById('crossDocumentTaskType')?.value)
            ? document.getElementById('crossDocumentTaskType').value : 'edit';
        if (!online || !target || target.expiresAt <= Date.now() || !text || !allowsDocumentAccess(target.access, identity.documentId, 'requestEdits')) return;
        sendButton.disabled = true;
        try {
            await distributed.submitGraph({
                graphId: `graph-${crypto.randomUUID()}`,
                tasks: [{ taskId: `task-${crypto.randomUUID()}`, type: taskType, instruction: text, target, ttlMs: 900000 }],
            });
            const statusText = { edit: 'review the selected passage', format: 'review the formatting ops', table: 'review the table preview', append: 'review appended content' }[taskType];
            document.getElementById('crossDocumentStatus').textContent = `Task sent: ${statusText} in the target document.`;
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
    const target = documents.get(select.value);
    const sourceId = coordinationClient?.identity?.documentId;
    const contextButton = document.getElementById('crossDocumentContextBtn');
    const sendButton = document.getElementById('crossDocumentSendBtn');
    if (contextButton) contextButton.disabled = !target || !allowsDocumentAccess(target.access, sourceId, 'readContext');
    if (sendButton) sendButton.disabled = !target || !allowsDocumentAccess(target.access, sourceId, 'requestEdits');
    const planButton = document.getElementById('crossDocumentPlanBtn');
    if (planButton) planButton.disabled = ![...documents.values()].some((entry) => entry.expiresAt > Date.now()
        && allowsDocumentAccess(entry.access, sourceId, 'requestEdits'));
}
