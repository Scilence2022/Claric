/** @jest-environment jsdom */
import { initCrossDocumentConnection } from '../src/taskpane/cross-document-controller.js';
import { createCoordinationClient } from '../src/taskpane/coordination-client.js';
import { createDocumentAgent } from '../src/taskpane/document-agent.js';
import { createContextRequestManager } from '../src/taskpane/context-requests.js';
import { createRemoteTaskRunner } from '../src/taskpane/remote-task-runner.js';
import { createDistributedTaskRuntime } from '../src/taskpane/distributed-task-runtime.js';
import { createProposalCard } from '../src/taskpane/ui/proposal-card.js';
import * as chatView from '../src/taskpane/ui/chat-view.js';
const fs = require('fs');
const path = require('path');

jest.mock('../src/taskpane/coordination-client.js', () => ({ createCoordinationClient: jest.fn() }));
jest.mock('../src/taskpane/document-agent.js', () => ({ createDocumentAgent: jest.fn() }));
jest.mock('../src/taskpane/context-requests.js', () => ({ createContextRequestManager: jest.fn() }));
jest.mock('../src/taskpane/remote-task-runner.js', () => ({ createRemoteTaskRunner: jest.fn() }));
jest.mock('../src/taskpane/distributed-task-runtime.js', () => ({ createDistributedTaskRuntime: jest.fn() }));
jest.mock('../src/taskpane/ui/proposal-card.js', () => ({ createProposalCard: jest.fn() }));
jest.mock('../src/taskpane/ui/chat-view.js', () => ({ addSystemNote: jest.fn() }));
jest.mock('../src/taskpane/ui/status-bar.js', () => ({ addLog: jest.fn() }));

const local = { workspaceId: 'w', documentId: 'a', instanceId: 'a-local' };
const target = { workspaceId: 'w', documentId: 'b', instanceId: 'b-server', title: 'Target', expiresAt: Date.now() + 60000 };
const flush = () => new Promise((resolve) => setTimeout(resolve, 0));
const flushMicrotasks = async () => {
    for (let index = 0; index < 12; index += 1) await Promise.resolve();
};
let options;
let client;
let agent;
let runtime;
let context;
let remote;
let message;
beforeEach(() => {
    jest.clearAllMocks();
    const html = fs.readFileSync(path.join(__dirname, '../src/taskpane/taskpane.html'), 'utf8');
    document.body.innerHTML = html.slice(html.indexOf('<body>') + 6, html.indexOf('</body>'));
    client = { identity: { ...local, instanceId: 'a-server' }, transport: { version: 2 }, stop: jest.fn(async () => {}) };
    client.start = jest.fn(async () => options.onSnapshot({ documents: [target], presence: {}, cursor: 1 }));
    createCoordinationClient.mockImplementation((value) => { options = value; return client; });
    agent = { prepareTask: jest.fn(), dispose: jest.fn() };
    context = { dispose: jest.fn(), handleEvent: jest.fn(), requestContext: jest.fn(async () => ({ snapshot: { text: 'Target data' } })) };
    remote = { dispose: jest.fn(), handleEvent: jest.fn() };
    runtime = { dispose: jest.fn(), handleEvent: jest.fn(), submitGraph: jest.fn(async () => {}) };
    message = { attachProposal: jest.fn(), finalizeForHistory: jest.fn(), markError: jest.fn() };
    createDocumentAgent.mockReturnValue(agent);
    createContextRequestManager.mockReturnValue(context);
    createRemoteTaskRunner.mockReturnValue(remote);
    createDistributedTaskRuntime.mockReturnValue(runtime);
    chatView.addSystemNote.mockReturnValue(message);
    createProposalCard.mockReturnValue({ markWarning: jest.fn(), setPaused: jest.fn() });
});

afterEach(() => {
    window.dispatchEvent(new Event('pagehide'));
    jest.clearAllTimers();
    jest.useRealTimers();
});

test('automatically connects, routes source tasks, and disposes access on disconnect', async () => {
    initCrossDocumentConnection(local);
    await flushMicrotasks();
    expect(createCoordinationClient).toHaveBeenCalledTimes(1);
    expect(options.identity).toEqual(local);
    expect(options.baseUrl).toBe(new URL('/coordination', location.href).href.replace(/\/$/, ''));
    const select = document.getElementById('crossDocumentTarget');
    select.value = target.instanceId; select.dispatchEvent(new Event('change'));
    document.getElementById('chatInput').value = 'Polish';
    document.getElementById('crossDocumentSendBtn').click(); await flush();
    expect(runtime.submitGraph).toHaveBeenCalledWith(expect.objectContaining({ tasks: [expect.objectContaining({ target, instruction: 'Polish' })] }));
    document.getElementById('crossDocumentContextBtn').click(); await flush();
    expect(context.requestContext).toHaveBeenCalledWith(target, { scope: 'document' });
    expect(chatView.addSystemNote).toHaveBeenCalledWith(expect.stringContaining('Target data'));
    document.getElementById('crossDocumentConnectBtn').click(); await flush();
    expect(client.stop).toHaveBeenCalledTimes(1);
    expect(agent.dispose).toHaveBeenCalledTimes(1);
    expect(document.getElementById('crossDocumentBar').hidden).toBe(true);
    expect(createRemoteTaskRunner.mock.calls[0][0].isOnline()).toBe(false);
    await new Promise((resolve) => setTimeout(resolve, 1100));
    expect(createCoordinationClient).toHaveBeenCalledTimes(1);
});

test('retries a temporary coordination failure and releases the failed client first', async () => {
    jest.useFakeTimers();
    const failure = Object.assign(new Error('Coordination request failed (503)'), { status: 503 });
    const failedClient = { identity: { ...local, instanceId: 'failed' }, transport: { version: 2 }, start: jest.fn(async () => { throw failure; }), stop: jest.fn(async () => {}) };
    const recoveredClient = { identity: { ...local, instanceId: 'recovered' }, transport: { version: 2 }, stop: jest.fn(async () => {}) };
    recoveredClient.start = jest.fn(async () => options.onSnapshot({ documents: [target], presence: {}, cursor: 2 }));
    let attempts = 0;
    createCoordinationClient.mockImplementation((value) => {
        options = value;
        attempts += 1;
        return attempts === 1 ? failedClient : recoveredClient;
    });

    initCrossDocumentConnection(local);
    await flushMicrotasks();
    expect(failedClient.stop).toHaveBeenCalledTimes(1);
    expect(createCoordinationClient).toHaveBeenCalledTimes(1);
    expect(document.getElementById('crossDocumentConnectionStatus').textContent).toContain('retrying in 1s');

    await jest.advanceTimersByTimeAsync(1000);
    await flushMicrotasks();
    expect(createCoordinationClient).toHaveBeenCalledTimes(2);
    expect(recoveredClient.start).toHaveBeenCalledTimes(1);
    expect(document.getElementById('crossDocumentConnectionStatus').textContent).toContain('Connected');
});

test('stops automatic retries for a pairing-token authorization failure', async () => {
    jest.useFakeTimers();
    const failure = Object.assign(new Error('Coordination request failed (401)'), { status: 401 });
    const deniedClient = { identity: { ...local, instanceId: 'denied' }, transport: { version: 2 }, start: jest.fn(async () => { throw failure; }), stop: jest.fn(async () => {}) };
    createCoordinationClient.mockImplementation((value) => { options = value; return deniedClient; });

    initCrossDocumentConnection(local);
    await flushMicrotasks();
    expect(deniedClient.stop).toHaveBeenCalledTimes(1);
    expect(document.getElementById('crossDocumentConnectionStatus').textContent).toContain('Authorization required');
    await jest.advanceTimersByTimeAsync(30000);
    expect(createCoordinationClient).toHaveBeenCalledTimes(1);
});

test('binds review to the target runtime with no auto-apply', async () => {
    initCrossDocumentConnection(local);
    await flushMicrotasks();
    const record = { source: target, title: 'Edit', items: [{ id: 'text-1' }] };
    const review = { apply: jest.fn(async () => ({ ok: false, conflict: { message: 'Stale' } })), reject: jest.fn(async () => ({})) };
    createRemoteTaskRunner.mock.calls[0][0].onProposal(record, review);
    expect(message.attachProposal).toHaveBeenCalledWith(expect.anything(), expect.objectContaining({ autoApply: false }));
    expect(review.apply).not.toHaveBeenCalled();
    await createProposalCard.mock.calls[0][0].onApply(['text-1']);
    expect(createProposalCard.mock.results[0].value.markWarning).toHaveBeenCalledWith('Stale');
});
