import { resolveDocumentIdentity, persistDocumentIdentity } from '../src/taskpane/document-identity.js';
const { webcrypto } = require('crypto');

const ORIGIN = { origin: 'https://localhost:3000' };
const resolve = (url, cryptoImpl = webcrypto, extra = {}) => resolveDocumentIdentity({
    office: { context: { document: { url, ...extra } } }, locationObject: ORIGIN, cryptoImpl,
});

function makeParts(initialXml = null) {
    let xml = initialXml;
    const part = {
        getXmlAsync: (cb) => cb({ status: 'succeeded', value: xml }),
        setXmlAsync: jest.fn((next, cb) => { xml = next; cb({ status: 'succeeded' }); }),
    };
    const parts = {
        getByNamespaceAsync: jest.fn((ns, cb) => cb({ status: 'succeeded', value: xml ? [part] : [] })),
        addAsync: jest.fn((next, cb) => { xml = next; cb({ status: 'succeeded', value: { id: 'part-1' } }); }),
    };
    return { parts, getXml: () => xml, part };
}

async function urlHash(url) {
    const digest = await webcrypto.subtle.digest('SHA-256', new TextEncoder().encode(url));
    return Array.from(new Uint8Array(digest), (byte) => byte.toString(16).padStart(2, '0')).join('');
}

function identityXml(documentId, hash, createdAt = 1) {
    return `<claricIdentity xmlns="urn:claric:identity" version="1">${JSON.stringify({ version: 1, documentId, urlHash: hash, createdAt })}</claricIdentity>`;
}

test('uses a stable full digest for known document URLs without exposing paths', async () => {
    const a = await resolve('file:///private/project/Document.docx');
    const b = await resolve('file:///private/project/Document.docx');
    const c = await resolve('file:///private/other/Document.docx');
    expect(a.documentId).toBe(b.documentId);
    expect(a.instanceId).not.toBe(b.instanceId);
    expect(c.documentId).not.toBe(a.documentId);
    expect(a.workspaceId).toBe(c.workspaceId);
    expect(a.documentId).toMatch(/^document-[a-f0-9]{64}$/);
    expect(a.identityKind).toBe('url');
    expect(a.identityForked).toBe(false);
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

test('reuses the stored custom-XML identity when the URL still matches', async () => {
    const url = 'file:///docs/Contract.docx';
    const { parts, getXml } = makeParts(identityXml('document-stored-identity', await urlHash(url)));
    const a = await resolve(url, webcrypto, { customXmlParts: parts });
    const b = await resolve(url, webcrypto, { customXmlParts: parts });
    expect(a.identityKind).toBe('custom-xml');
    expect(a.documentId).toBe('document-stored-identity');
    expect(b.documentId).toBe('document-stored-identity');
    expect(a.documentEphemeral).toBe(false);
    expect(parts.addAsync).not.toHaveBeenCalled();
    expect(getXml()).toContain('document-stored-identity');
});

test('derives a deterministic fork identity on Save As without rewriting the file at startup', async () => {
    const originalUrl = 'file:///docs/Contract.docx';
    const forkUrl = 'file:///docs/Contract-v2.docx';
    const storedId = 'document-stored-identity';
    const { parts, part } = makeParts(identityXml(storedId, await urlHash(originalUrl)));
    const a = await resolve(forkUrl, webcrypto, { customXmlParts: parts });
    const b = await resolve(forkUrl, webcrypto, { customXmlParts: parts });
    expect(a.identityKind).toBe('fork');
    expect(a.identityForked).toBe(true);
    expect(a.documentId).toBe(b.documentId);
    expect(a.documentId).not.toBe(storedId);
    expect(a.documentId).toMatch(/^document-[a-f0-9]{64}$/);
    expect(part.setXmlAsync).not.toHaveBeenCalled();
    expect(parts.addAsync).not.toHaveBeenCalled();
    // The original file identity is untouched and differs from the fork.
    const original = await resolve(originalUrl, webcrypto, { customXmlParts: parts });
    expect(original.documentId).toBe(storedId);
    expect(original.documentId).not.toBe(a.documentId);
});

test('keeps the stored identity when the host exposes no URL', async () => {
    const { parts } = makeParts(identityXml('document-stored-identity', null));
    const a = await resolve('', webcrypto, { customXmlParts: parts });
    expect(a.identityKind).toBe('custom-xml');
    expect(a.documentId).toBe('document-stored-identity');
    expect(a.documentEphemeral).toBe(false);
});

test('falls back to the URL hash when the stored part is malformed', async () => {
    const url = 'file:///docs/Contract.docx';
    for (const xml of ['not xml at all', '<claricIdentity xmlns="urn:claric:identity" version="1">{"version":2}</claricIdentity>', `<claricIdentity xmlns="urn:claric:identity" version="1">${JSON.stringify({ version: 1, documentId: 'evil/../../path', urlHash: null, createdAt: 1 })}</claricIdentity>`]) {
        const { parts } = makeParts(xml);
        const identity = await resolve(url, webcrypto, { customXmlParts: parts });
        expect(identity.identityKind).toBe('url');
        expect(identity.documentId).toBe(`document-${await urlHash(url)}`);
    }
});

test('persists a url identity lazily without leaking the path', async () => {
    const url = 'file:///private/docs/Contract.docx';
    const { parts, getXml } = makeParts(null);
    const identity = await resolve(url, webcrypto, { customXmlParts: parts });
    expect(identity.identityKind).toBe('url');
    await expect(persistDocumentIdentity(identity, { office: { context: { document: { url, customXmlParts: parts } } }, cryptoImpl: webcrypto })).resolves.toBe(true);
    expect(parts.addAsync).toHaveBeenCalledTimes(1);
    expect(getXml()).toContain(identity.documentId);
    expect(getXml()).not.toContain('/private/');
    expect(getXml()).toContain(await urlHash(url));
    // Once persisted, reloads resolve through the stored identity.
    const reloaded = await resolve(url, webcrypto, { customXmlParts: parts });
    expect(reloaded.identityKind).toBe('custom-xml');
    expect(reloaded.documentId).toBe(identity.documentId);
    await expect(persistDocumentIdentity(reloaded, { office: { context: { document: { url, customXmlParts: parts } } }, cryptoImpl: webcrypto })).resolves.toBe(true);
});

test('rewrites the stored part after persisting a fork identity', async () => {
    const originalUrl = 'file:///docs/Contract.docx';
    const forkUrl = 'file:///docs/Contract-v2.docx';
    const { parts, getXml, part } = makeParts(identityXml('document-stored-identity', await urlHash(originalUrl)));
    const fork = await resolve(forkUrl, webcrypto, { customXmlParts: parts });
    await expect(persistDocumentIdentity(fork, { office: { context: { document: { url: forkUrl, customXmlParts: parts } } }, cryptoImpl: webcrypto })).resolves.toBe(true);
    expect(part.setXmlAsync).toHaveBeenCalledTimes(1);
    expect(getXml()).toContain(fork.documentId);
    expect(getXml()).toContain(await urlHash(forkUrl));
    const reloaded = await resolve(forkUrl, webcrypto, { customXmlParts: parts });
    expect(reloaded.identityKind).toBe('custom-xml');
    expect(reloaded.documentId).toBe(fork.documentId);
});

test('persist is a best-effort no-op for ephemeral identities and unavailable stores', async () => {
    const ephemeral = await resolve('');
    await expect(persistDocumentIdentity(ephemeral)).resolves.toBe(false);
    const { parts } = makeParts(null);
    const urlIdentity = await resolve('file:///docs/a.docx', webcrypto, { customXmlParts: parts });
    await expect(persistDocumentIdentity(urlIdentity, { office: { context: { document: { url: '' } } }, cryptoImpl: webcrypto })).resolves.toBe(false);
    await expect(persistDocumentIdentity(urlIdentity, { office: { context: { document: { url: 'file:///docs/a.docx' } } }, cryptoImpl: webcrypto })).resolves.toBe(false);
    await expect(persistDocumentIdentity(null)).resolves.toBe(false);
    expect(parts.addAsync).not.toHaveBeenCalled();
});
