import { resolveDocumentIdentity } from '../src/taskpane/document-identity.js';
const { webcrypto } = require('crypto');

const resolve = (url, cryptoImpl = webcrypto) => resolveDocumentIdentity({
    office: { context: { document: { url } } }, locationObject: { origin: 'https://localhost:3000' }, cryptoImpl,
});

test('uses a stable full digest for known document URLs without exposing paths', async () => {
    const a = await resolve('file:///private/project/Document.docx');
    const b = await resolve('file:///private/project/Document.docx');
    const c = await resolve('file:///private/other/Document.docx');
    expect(a.documentId).toBe(b.documentId);
    expect(a.instanceId).not.toBe(b.instanceId);
    expect(c.documentId).not.toBe(a.documentId);
    expect(a.workspaceId).toBe(c.workspaceId);
    expect(a.documentId).toMatch(/^document-[a-f0-9]{64}$/);
    expect(JSON.stringify(a)).not.toContain('/private/');
});

test('marks unidentified documents ephemeral instead of equating unsaved documents', async () => {
    const a = await resolve('');
    const b = await resolve('');
    expect(a.documentEphemeral).toBe(true);
    expect(a.documentId).not.toBe(b.documentId);
    expect(a.workspaceId).toBe(b.workspaceId);
    expect((await resolve('file:///test.docx', {})).documentEphemeral).toBe(true);
});
