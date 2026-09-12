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
    document.getElementById('crossDocumentUrl').value = 'http://127.0.0.1:3010/coordination';
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

test('requires explicit connect, routes source tasks, and disposes access on disconnect', async () => {
    initCrossDocumentConnection(local);
    expect(createCoordinationClient).not.toHaveBeenCalled();
    document.getElementById('crossDocumentConnectBtn').click(); await flush();
    expect(options.identity).toEqual(local);
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
});

test('binds review to the target runtime with no auto-apply', async () => {
    initCrossDocumentConnection(local);
    document.getElementById('crossDocumentConnectBtn').click(); await flush();
    const record = { source: target, title: 'Edit', items: [{ id: 'text-1' }] };
    const review = { apply: jest.fn(async () => ({ ok: false, conflict: { message: 'Stale' } })), reject: jest.fn(async () => ({})) };
    createRemoteTaskRunner.mock.calls[0][0].onProposal(record, review);
    expect(message.attachProposal).toHaveBeenCalledWith(expect.anything(), expect.objectContaining({ autoApply: false }));
    expect(review.apply).not.toHaveBeenCalled();
    await createProposalCard.mock.calls[0][0].onApply(['text-1']);
    expect(createProposalCard.mock.results[0].value.markWarning).toHaveBeenCalledWith('Stale');
    document.getElementById('crossDocumentConnectBtn').click(); await flush();
});
