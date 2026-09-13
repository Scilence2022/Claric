import { createDocumentAgent } from '../src/taskpane/document-agent.js';
const { webcrypto } = require('crypto');

describe('document agent', () => {
    const identity = { workspaceId: 'workspace', documentId: 'document-b', instanceId: 'instance-b' };
    const create = (options = {}) => createDocumentAgent({ identity, cryptoImpl: webcrypto, ...options });

    test('bounds CJK tokens, strips arbitrary metadata and handles cursor contextText', async () => {
        const agent = create({ actions: {
            readSelectionContent: async () => ({ text: '中'.repeat(100), range: {}, images: [{ base64: 'private' }], outline: [{ text: 'unsafe', token: 'secret' }] }),
            readCursorContext: async () => ({ kind: 'cursor', contextText: 'Cursor paragraph' }),
        } });
        const result = await agent.readContext({ scope: 'selection', maxChars: 20, maxTokens: 10 });
        expect(result.text).toBe('中'.repeat(10));
        expect(result.truncation.text).toBe(true);
        expect(result.documentRevision).toMatch(/^revision-[a-f0-9]{64}$/);
        expect(JSON.stringify(result)).not.toMatch(/private|secret|unsafe/);
        expect((await agent.readContext({ scope: 'cursor' })).text).toBe('Cursor paragraph');
    });

    test('reads the document and outline but rejects unknown scopes', async () => {
        const agent = create({ actions: { extractDocumentStructured: async () => '# Title\nText\n## Part\nMore' } });
        expect((await agent.readContext({ scope: 'document' })).text).toContain('More');
        expect((await agent.readContext({ scope: 'outline' })).text).toBe('# Title\n## Part');
        await expect(agent.getDocumentRevision('full-ooxml')).rejects.toMatchObject({ code: 'UNSUPPORTED_SCOPE' });
    });

    test('prepares with the local model, writes locally once, and never serializes runtime callbacks', async () => {
        let text = 'Original text';
        const prepared = { selectionText: text, amendedText: 'Edited text' };
        const prepare = jest.fn(async () => prepared);
        const apply = jest.fn(async () => { text = prepared.amendedText; });
        const appState = {};
        const agent = create({ appState, actions: {
            readSelectionContent: async () => ({ text }), prepareSelectionAmendment: prepare, applySelectionAmendment: apply,
        } });
        const record = await agent.prepareTask({ taskId: 'task-1', taskType: 'edit', instruction: 'Polish the passage' });
        record.target = identity;
        expect(prepare).toHaveBeenCalledWith(expect.objectContaining({ appState }), expect.objectContaining({ promptTemplate: 'Polish the passage' }));
        expect(record.items[0]).toMatchObject({ before: 'Original text', after: 'Edited text' });
        expect(JSON.parse(JSON.stringify(record))).toEqual(record);
        const result = await agent.applyProposal(['text-1'], record);
        expect(result.appliedItemIds).toEqual(['text-1']);
        expect(appState.isProcessingDoc).toBe(false);
        await expect(agent.applyProposal(['text-1'], record)).rejects.toMatchObject({ code: 'ALREADY_ATTEMPTED' });
        expect(apply).toHaveBeenCalledTimes(1);
    });

    test('rejects changes during model preparation and unsupported tasks instead of fabricating a proposal', async () => {
        let text = 'Before';
        const agent = create({ actions: {
            readSelectionContent: async () => ({ text }),
            prepareSelectionAmendment: async () => { text = 'User changed it'; return { selectionText: 'Before', amendedText: 'Model edit' }; },
        } });
        await expect(agent.prepareTask({ taskId: 'task', taskType: 'edit', instruction: 'Edit' })).rejects.toMatchObject({ code: 'STALE' });
        await expect(agent.prepareTask({ taskId: 'other', taskType: 'arbitrary', instruction: 'Edit' })).rejects.toMatchObject({ code: 'UNSUPPORTED_TASK' });
    });

    test('does not mark skipped or uncertain writes successful or automatically retry them', async () => {
        const apply = jest.fn(async () => ({ skipped: true, reason: 'Selection moved' }));
        const agent = create({ actions: {
            readSelectionContent: async () => ({ text: 'Before' }),
            prepareSelectionAmendment: async () => ({ selectionText: 'Before', amendedText: 'After' }),
            applySelectionAmendment: apply,
        } });
        const record = { ...await agent.prepareTask({ taskId: 'task', taskType: 'edit', instruction: 'Edit' }), target: identity };
        await expect(agent.applyProposal(['text-1'], { ...record, target: { ...identity, documentId: 'wrong' } })).rejects.toMatchObject({ code: 'TARGET_MISMATCH' });
        await expect(agent.applyProposal(['text-1'], record)).rejects.toMatchObject({ code: 'WRITE_NOT_COMPLETED' });
        await expect(agent.applyProposal(['text-1'], record)).rejects.toMatchObject({ code: 'ALREADY_ATTEMPTED' });
        expect(apply).toHaveBeenCalledTimes(1);
    });

    test('rejects busy and cancelled work before invoking readers', async () => {
        const read = jest.fn();
        const agent = create({ appState: { isProcessing: true }, actions: { readSelectionContent: read } });
        await expect(agent.prepareTask({})).rejects.toMatchObject({ code: 'DOCUMENT_BUSY' });
        const controller = new AbortController(); controller.abort();
        await expect(agent.readContext({}, { signal: controller.signal })).rejects.toMatchObject({ name: 'AbortError' });
        expect(read).not.toHaveBeenCalled();
    });
});
