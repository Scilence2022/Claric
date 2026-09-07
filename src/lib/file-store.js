import { ATTACHMENT_LIMITS, detectAttachmentKind, parseAttachment, validateAttachment } from './file-attachments.js';

export const FILE_STORE_LIMITS = Object.freeze({
    MAX_FILES: 200,
    MAX_TOTAL_BYTES: 100 * 1024 * 1024,
    MAX_TEXT_CHARS: 2_000_000,
    MAX_PAGE_ITEMS: 50,
    MAX_READ_CHARS: 12000,
    MAX_QUERY_CHARS: 200,
    MAX_NAME_CHARS: 255,
});

const DB_NAME = 'claric-file-resources';
const STORE_NAME = 'files';
const IMAGE_MIME = Object.freeze({ png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', webp: 'image/webp' });

function opaqueId(prefix) {
    if (typeof crypto !== 'undefined' && typeof crypto.randomUUID === 'function') return `${prefix}_${crypto.randomUUID()}`;
    return `${prefix}_${Date.now().toString(36)}_${Math.random().toString(36).slice(2, 12)}`;
}

function validateId(id) {
    if (typeof id !== 'string' || !/^[a-zA-Z0-9_-]{1,128}$/.test(id)) throw new Error('A valid opaque fileId is required.');
}

function fileName(name) {
    if (typeof name !== 'string' || !name.trim() || name.length > FILE_STORE_LIMITS.MAX_NAME_CHARS
        || [...name].some((char) => char.charCodeAt(0) < 32 || char.charCodeAt(0) === 127)) {
        throw new Error('File name must contain 1–255 characters without control characters.');
    }
    return name.trim();
}

function pageNumber(value, fallback, max, min = 0) {
    if (value === undefined) return fallback;
    const number = Number(value);
    if (!Number.isFinite(number) || number < min) throw new Error('Invalid paging value.');
    return Math.min(Math.floor(number), max);
}

function metadata(record) {
    if (!record) return null;
    const { fileId, versionId, name, kind, mimeType, size, source, createdAt, updatedAt, parseStatus, parseError, textLength, persistence } = record;
    return { fileId, versionId, name, kind, mimeType, size, source, createdAt, updatedAt, parseStatus, parseError, textLength, persistence };
}

function checkCapacity(records, record) {
    if (records.length >= FILE_STORE_LIMITS.MAX_FILES) throw new Error('File library is full. Delete a file before uploading.');
    if (records.reduce((sum, item) => sum + item.size, 0) + record.size > FILE_STORE_LIMITS.MAX_TOTAL_BYTES) {
        throw new Error('File library exceeds the 100 MB storage limit.');
    }
}

export function createIndexedDbFileAdapter({ indexedDB: factory = globalThis.indexedDB } = {}) {
    let dbPromise;
    function open() {
        if (!factory) return Promise.reject(new Error('File storage is unavailable: IndexedDB is not supported.'));
        if (!dbPromise) {
            dbPromise = new Promise((resolve, reject) => {
                let settled = false;
                const req = factory.open(DB_NAME, 1);
                req.onupgradeneeded = () => {
                    if (!req.result.objectStoreNames.contains(STORE_NAME)) req.result.createObjectStore(STORE_NAME, { keyPath: 'fileId' });
                };
                req.onerror = () => { settled = true; reject(req.error || new Error('File storage could not be opened.')); };
                req.onblocked = () => { settled = true; reject(new Error('File storage is blocked by another window. Close it and retry.')); };
                req.onsuccess = () => {
                    if (settled) { req.result.close(); return; }
                    const db = req.result;
                    db.onversionchange = () => { db.close(); dbPromise = undefined; };
                    resolve(db);
                };
            }).catch((error) => { dbPromise = undefined; throw error; });
        }
        return dbPromise;
    }
    async function transaction(mode, action) {
        const db = await open();
        return new Promise((resolve, reject) => {
            const tx = db.transaction(STORE_NAME, mode);
            let value;
            let failure;
            tx.oncomplete = () => resolve(value);
            tx.onabort = () => reject(failure || tx.error || new Error('File storage transaction aborted.'));
            tx.onerror = () => { failure = tx.error || new Error('File storage transaction failed.'); };
            const fail = (error) => { failure = error; tx.abort(); };
            try { action(tx.objectStore(STORE_NAME), (result) => { value = result; }, fail); }
            catch (error) { fail(error); }
        });
    }
    return {
        persistence: 'indexeddb',
        list: () => transaction('readonly', (store, done) => {
            const items = [];
            const req = store.openCursor();
            req.onsuccess = () => {
                const cursor = req.result;
                if (!cursor) { done(items); return; }
                items.push(metadata(cursor.value));
                cursor.continue();
            };
        }),
        get: (id) => transaction('readonly', (store, done) => {
            const req = store.get(id);
            req.onsuccess = () => done(req.result);
        }),
        add: (record) => transaction('readwrite', (store, done, fail) => {
            const items = [];
            const req = store.openCursor();
            req.onsuccess = () => {
                const cursor = req.result;
                if (cursor) { items.push(metadata(cursor.value)); cursor.continue(); return; }
                try { checkCapacity(items, record); store.add(record); done(record); }
                catch (error) { fail(error); }
            };
        }),
        update: (id, transform) => transaction('readwrite', (store, done, fail) => {
            const req = store.get(id);
            req.onsuccess = () => {
                try {
                    const next = transform(req.result);
                    if (next) store.put(next);
                    else store.delete(id);
                    done(next);
                } catch (error) { fail(error); }
            };
        }),
    };
}

export function createMemoryFileAdapter() {
    const records = new Map();
    return {
        persistence: 'memory',
        async list() { return [...records.values()].map(metadata); },
        async get(id) { const record = records.get(id); return record ? { ...record } : undefined; },
        async add(record) { checkCapacity([...records.values()], record); records.set(record.fileId, { ...record }); return record; },
        async update(id, transform) {
            const next = transform(records.has(id) ? { ...records.get(id) } : undefined);
            if (next) records.set(id, { ...next });
            else records.delete(id);
            return next;
        },
    };
}

export function createFileStore({ adapter = createIndexedDbFileAdapter(), parse = parseAttachment } = {}) {
    function requireRecord(record, versionId) {
        if (!record) throw new Error('File resource not found. It may have been deleted.');
        if (versionId && record.versionId !== versionId) throw new Error('File version is no longer available.');
        return record;
    }
    async function recordFor(id, versionId) {
        validateId(id);
        return requireRecord(await adapter.get(id), versionId);
    }
    function requireParsed(record) {
        if (record.parseStatus !== 'ready' && record.parseStatus !== 'empty') throw new Error(record.parseError || 'File text has not been extracted. Re-upload this file.');
    }
    async function imageAttachment(record) {
        const parsed = await parse(new File([record.body], record.originalName || record.name, { type: record.mimeType }));
        if (!/^data:image\/(?:png|jpeg|gif|webp);base64,[a-zA-Z0-9+/]*={0,2}$/.test(parsed.dataUrl || '')
            || parsed.dataUrl.length > 6 * 1024 * 1024) throw new Error('Image input is unsafe or exceeds the image budget.');
        return { dataUrl: parsed.dataUrl };
    }
    async function saveFile(file, { name = file?.name, source = 'upload' } = {}) {
        if (!(file instanceof Blob)) throw new Error('A File or Blob is required.');
        const nameValue = fileName(name);
        const kind = detectAttachmentKind(nameValue, file.type);
        const verdict = validateAttachment({ name: nameValue, kind, size: file.size });
        if (verdict.ok === false) throw new Error(verdict.error);
        if (!['upload', 'library', 'attachment'].includes(source)) throw new Error('Invalid file source.');
        let mimeType = file.type || '';
        if (kind === 'image') {
            mimeType = IMAGE_MIME[nameValue.split('.').pop().toLowerCase()] || mimeType;
            if (!Object.values(IMAGE_MIME).some((mime) => mime === mimeType)) throw new Error('Only PNG, JPEG, GIF and WebP images are supported.');
        }
        const body = file.slice(0, file.size, mimeType);
        const now = new Date().toISOString();
        const record = {
            fileId: opaqueId('file'), versionId: opaqueId('version'), name: nameValue, originalName: nameValue,
            kind, mimeType, size: body.size, source, createdAt: now, updatedAt: now, body,
            parseStatus: 'ready', parseError: null, textLength: 0, persistence: adapter.persistence,
        };
        try {
            const parsed = await parse(new File([body], nameValue, { type: mimeType }));
            if (kind !== 'image') {
                const text = typeof parsed.text === 'string' ? parsed.text : '';
                if (text.length > FILE_STORE_LIMITS.MAX_TEXT_CHARS) throw new Error('Extracted text exceeds the 2,000,000 character limit.');
                record.text = text;
                record.textLength = text.length;
                if (!text.trim()) {
                    record.parseStatus = 'empty';
                    record.parseError = kind === 'pdf' ? 'No text extracted. This PDF may be scanned; OCR is not available.' : 'No text extracted from this file.';
                }
            } else if (!/^data:image\/(?:png|jpeg|gif|webp);base64,[a-zA-Z0-9+/]*={0,2}$/.test(parsed.dataUrl || '') || parsed.dataUrl.length > 6 * 1024 * 1024) {
                throw new Error('Image input is unsafe or exceeds the image budget.');
            }
        } catch (error) {
            record.parseStatus = 'error';
            record.parseError = String(error.message || 'File could not be parsed.').slice(0, 300);
        }
        await adapter.add(record);
        return metadata(record);
    }
    /** @param {{query?: string, kind?: string, limit?: number, offset?: number, allowedIds?: string[]}} options */
    async function listFiles({ query = '', kind = '', limit, offset, allowedIds } = {}) {
        if (typeof query !== 'string' || query.length > FILE_STORE_LIMITS.MAX_QUERY_CHARS) throw new Error('Search query must be at most 200 characters.');
        if (kind && !['text', 'image', 'pdf', 'docx'].includes(kind)) throw new Error('Invalid file kind.');
        const start = pageNumber(offset, 0, Number.MAX_SAFE_INTEGER);
        const count = pageNumber(limit, 20, FILE_STORE_LIMITS.MAX_PAGE_ITEMS, 1);
        const allowed = allowedIds === undefined ? null : new Set(allowedIds);
        const q = query.toLowerCase();
        const records = (await adapter.list()).filter((item) => (!allowed || allowed.has(item.fileId))
            && (!kind || item.kind === kind) && (!q || item.name.toLowerCase().includes(q)));
        records.sort((a, b) => b.updatedAt.localeCompare(a.updatedAt) || a.fileId.localeCompare(b.fileId));
        const items = records.slice(start, start + count).map(metadata);
        return { items, total: records.length, offset: start, limit: count, hasMore: start + items.length < records.length };
    }
    async function getFile(id, { versionId = '' } = {}) { return metadata(await recordFor(id, versionId)); }
    async function readFile(id, { offset = 0, limit = FILE_STORE_LIMITS.MAX_READ_CHARS, versionId = '' } = {}) {
        const start = pageNumber(offset, 0, Number.MAX_SAFE_INTEGER);
        const count = pageNumber(limit, FILE_STORE_LIMITS.MAX_READ_CHARS, FILE_STORE_LIMITS.MAX_READ_CHARS, 1);
        const record = await recordFor(id, versionId);
        requireParsed(record);
        if (record.kind === 'image') return { file: metadata(record), attachment: await imageAttachment(record) };
        const text = record.text || '';
        return { file: metadata(record), text: text.slice(start, start + count), offset: start, limit: count, totalChars: text.length, hasMore: start + count < text.length, nextOffset: start + count < text.length ? start + count : null };
    }
    async function getAttachment(id, { versionId = '' } = {}) {
        const record = await recordFor(id, versionId);
        requireParsed(record);
        if (record.kind === 'image') return { ...metadata(record), ...await imageAttachment(record) };
        const text = record.text || '';
        return { ...metadata(record), text: text.slice(0, ATTACHMENT_LIMITS.MAX_CONTEXT_CHARS), truncated: text.length > ATTACHMENT_LIMITS.MAX_CONTEXT_CHARS };
    }
    async function renameFile(id, name) {
        validateId(id);
        const value = fileName(name);
        return metadata(await adapter.update(id, (record) => ({ ...requireRecord(record), name: value, updatedAt: new Date().toISOString() })));
    }
    async function deleteFile(id) {
        validateId(id);
        await adapter.update(id, (record) => { requireRecord(record); return null; });
        return { fileId: id, deleted: true };
    }
    async function searchFiles(query, options = {}) { return listFiles({ ...options, query }); }
    return { saveFile, listFiles, getFile, readFile, getAttachment, renameFile, deleteFile, searchFiles };
}

const defaultStore = createFileStore();
export const { saveFile, listFiles, getFile, readFile, getAttachment, renameFile, deleteFile, searchFiles } = defaultStore;
