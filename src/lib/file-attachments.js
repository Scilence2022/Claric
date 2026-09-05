/**
 * File Attachments Module
 *
 * Pure logic behind the chat input bar's file upload: type detection, size/
 * count validation, and parsing (.txt/.md/.csv/... via File.text(), images
 * via base64 data URLs, .docx via mammoth, .pdf via pdf.js — the latter two
 * loaded lazily so the parsers never inflate the first paint). Also builds
 * the prompt context block that prepends attached text to a chat turn.
 *
 * No DOM access: the UI layer (ui/input-bar.js) owns chips and the file
 * picker; this module only sees File-like objects ({ name, size, type,
 * text(), arrayBuffer() }).
 *
 * @module lib/file-attachments
 */

/** Attachment kinds emitted by detectAttachmentKind. */
export const ATTACHMENT_KIND = Object.freeze({
    TEXT: 'text',
    IMAGE: 'image',
    DOCX: 'docx',
    PDF: 'pdf',
});

/**
 * Upload limits. The image cap mirrors agent-actions' read_image ceiling
 * (6M base64 chars ≈ 4.5 MiB binary); MAX_CONTEXT_CHARS bounds the total
 * text ever injected into one prompt so a 10 MB upload cannot blow up the
 * LLM request.
 */
export const ATTACHMENT_LIMITS = Object.freeze({
    MAX_FILES: 5,
    MAX_TEXT_FILE_BYTES: 10 * 1024 * 1024,
    MAX_IMAGE_FILE_BYTES: 4.5 * 1024 * 1024, // ≈ 6M base64 chars on the wire
    MAX_TOTAL_BYTES: 10 * 1024 * 1024,
    MAX_CONTEXT_CHARS: 200000,
});

const TEXT_EXTENSIONS = ['.txt', '.md', '.markdown', '.csv', '.json', '.log', '.xml', '.yaml', '.yml'];
const IMAGE_EXTENSIONS = ['.png', '.jpg', '.jpeg', '.gif', '.webp'];

/**
 * Classifies a file by extension first, then by MIME type.
 *
 * @param {string} name - File name (extension wins over MIME: a .md file
 *   served as application/octet-stream is still text)
 * @param {string} [mimeType] - File.type when available
 * @returns {'text'|'image'|'docx'|'pdf'|null} Null for unsupported types.
 */
export function detectAttachmentKind(name, mimeType) {
    const lower = String(name || '').toLowerCase();
    const dot = lower.lastIndexOf('.');
    const ext = dot === -1 ? '' : lower.slice(dot);
    if (ext === '.docx') return ATTACHMENT_KIND.DOCX;
    if (ext === '.pdf') return ATTACHMENT_KIND.PDF;
    if (TEXT_EXTENSIONS.includes(ext)) return ATTACHMENT_KIND.TEXT;
    if (IMAGE_EXTENSIONS.includes(ext)) return ATTACHMENT_KIND.IMAGE;
    const mime = String(mimeType || '').toLowerCase();
    if (mime === 'application/pdf') return ATTACHMENT_KIND.PDF;
    if (mime === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document') return ATTACHMENT_KIND.DOCX;
    if (mime.startsWith('image/')) return ATTACHMENT_KIND.IMAGE;
    if (mime.startsWith('text/') || mime === 'application/json') return ATTACHMENT_KIND.TEXT;
    return null;
}

/**
 * Validates one candidate against the limits and the already-attached list.
 *
 * @param {{name: string, size: number, kind: string|null}} file - Candidate
 *   (kind from detectAttachmentKind; null is a rejection, not an error here)
 * @param {Array<{name: string, size: number}>} existing - Accepted so far
 * @returns {{ok: true} | {ok: false, error: string}}
 */
export function validateAttachment(file, existing = []) {
    const name = (file && file.name) || '(unnamed)';
    if (!file || !file.kind) {
        return { ok: false, error: `${name}: unsupported file type (use text, image, .docx or .pdf).` };
    }
    const size = Number(file.size) || 0;
    const perFileCap = file.kind === ATTACHMENT_KIND.IMAGE
        ? ATTACHMENT_LIMITS.MAX_IMAGE_FILE_BYTES
        : ATTACHMENT_LIMITS.MAX_TEXT_FILE_BYTES;
    if (size > perFileCap) {
        return { ok: false, error: `${name}: ${formatBytes(size)} exceeds the ${formatBytes(perFileCap)} per-file limit.` };
    }
    const list = Array.isArray(existing) ? existing : [];
    if (list.length >= ATTACHMENT_LIMITS.MAX_FILES) {
        return { ok: false, error: `${name}: at most ${ATTACHMENT_LIMITS.MAX_FILES} attachments per message.` };
    }
    const used = list.reduce((sum, f) => sum + (Number(f.size) || 0), 0);
    if (used + size > ATTACHMENT_LIMITS.MAX_TOTAL_BYTES) {
        return { ok: false, error: `${name}: attachments total would exceed ${formatBytes(ATTACHMENT_LIMITS.MAX_TOTAL_BYTES)}.` };
    }
    return { ok: true };
}

/**
 * Human-readable byte count for error messages and chip tooltips.
 * @param {number} bytes
 * @returns {string}
 */
export function formatBytes(bytes) {
    const n = Number(bytes) || 0;
    if (n < 1024) return `${n} B`;
    if (n < 1024 * 1024) return `${(n / 1024).toFixed(1)} KB`;
    return `${(n / (1024 * 1024)).toFixed(1)} MB`;
}

/**
 * Parses one validated File into an attachment object:
 *   text/docx/pdf → { name, kind, size, text }
 *   image         → { name, kind, size, dataUrl }
 * Throws with a user-facing message when reading or parsing fails.
 *
 * mammoth and pdf.js are dynamic imports: multi-hundred-KB parsers stay out
 * of the initial bundle and only load when such a file is actually attached.
 *
 * @param {File} file
 * @returns {Promise<{name: string, kind: string, size: number, text?: string, dataUrl?: string}>}
 */
export async function parseAttachment(file) {
    const kind = detectAttachmentKind(file && file.name, file && file.type);
    const base = { name: (file && file.name) || '(unnamed)', kind, size: Number(file && file.size) || 0 };
    try {
        if (kind === ATTACHMENT_KIND.TEXT) {
            return { ...base, text: await _readText(file) };
        }
        if (kind === ATTACHMENT_KIND.IMAGE) {
            return { ...base, dataUrl: await _readAsDataUrl(file) };
        }
        if (kind === ATTACHMENT_KIND.DOCX) {
            return { ...base, text: await _extractDocxText(file) };
        }
        if (kind === ATTACHMENT_KIND.PDF) {
            return { ...base, text: await _extractPdfText(file) };
        }
    } catch (err) {
        throw new Error(`${base.name}: ${err && err.message ? err.message : 'could not be read'}`);
    }
    throw new Error(`${base.name}: unsupported file type.`);
}

/**
 * Builds the prompt context block for one turn. Text-bearing attachments
 * (text/docx/pdf) become labeled sections; images are listed by name (their
 * bytes travel as image_url parts on vision-capable QA turns instead).
 * Total output is capped at MAX_CONTEXT_CHARS with an explicit truncation
 * note so the model knows the content was cut.
 *
 * @param {Array<{name: string, kind: string, text?: string, dataUrl?: string}>} attachments
 * @returns {string} '' when nothing attachable, else a block starting with \n\n
 */
export function buildAttachmentContext(attachments) {
    const list = Array.isArray(attachments) ? attachments.filter(Boolean) : [];
    if (list.length === 0) return '';
    let out = '';
    let truncated = false;
    for (const att of list) {
        let block;
        if (att.kind === ATTACHMENT_KIND.IMAGE) {
            block = `\n\n--- ATTACHED IMAGE: ${att.name} (sent as an image input to vision-capable models) ---`;
        } else {
            block = `\n\n--- ATTACHED FILE: ${att.name} ---\n${typeof att.text === 'string' ? att.text : '(no text extracted)'}`;
        }
        if (out.length + block.length > ATTACHMENT_LIMITS.MAX_CONTEXT_CHARS) {
            truncated = true;
            break;
        }
        out += block;
    }
    if (truncated) {
        out += '\n\n(Additional attachment content omitted — attachment context limit reached.)';
    }
    return out;
}

/**
 * Splits parsed attachments into prompt text sources and image payloads.
 * @param {Array<object>} attachments
 * @returns {{textAttachments: Array<object>, imageAttachments: Array<{name: string, dataUrl: string}>}}
 */
export function splitAttachments(attachments) {
    const list = Array.isArray(attachments) ? attachments.filter(Boolean) : [];
    return {
        textAttachments: list.filter((a) => a.kind !== ATTACHMENT_KIND.IMAGE),
        imageAttachments: list
            .filter((a) => a.kind === ATTACHMENT_KIND.IMAGE && typeof a.dataUrl === 'string' && a.dataUrl)
            .map((a) => ({ name: a.name, dataUrl: a.dataUrl })),
    };
}

/**
 * Strips attachments down to display/persistence metadata. Image data URLs
 * and extracted text never enter the session store (localStorage is ~5 MB);
 * the chip list and history only need name/kind/size.
 *
 * This is the single definition of that reduction: taskpane/message-shape.js
 * re-exports it for the session save/load legs, so a chip row, a persisted
 * message, and a restored message can never disagree about the shape.
 * Entries without a name are dropped — parseAttachment always assigns one, so
 * a nameless entry is corrupt input, not an unnamed file.
 *
 * @param {Array<{name: string, kind: string, size: number}>} attachments
 * @returns {Array<{name: string, kind: string, size: number}>}
 */
export function attachmentMeta(attachments) {
    return (Array.isArray(attachments) ? attachments : [])
        .filter((a) => a && typeof a === 'object' && a.name)
        .map((a) => ({
            name: String(a.name),
            kind: typeof a.kind === 'string' && a.kind ? a.kind : ATTACHMENT_KIND.TEXT,
            size: Number(a.size) || 0,
        }));
}

/**
 * Reads a File as text. Falls back to FileReader for jsdom (its Blob lacks
 * .text()). @private
 */
async function _readText(file) {
    if (typeof file.text === 'function') return file.text();
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve(reader.result);
        reader.onerror = () => reject(reader.error || new Error('read failed'));
        reader.readAsText(file);
    });
}

/**
 * Reads a File as an ArrayBuffer with the same FileReader fallback.
 * @private
 */
async function _readArrayBuffer(file) {
    if (typeof file.arrayBuffer === 'function') return file.arrayBuffer();
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve(reader.result);
        reader.onerror = () => reject(reader.error || new Error('read failed'));
        reader.readAsArrayBuffer(file);
    });
}

/**
 * Reads a File as a base64 data URL without FileReader's readAsDataURL when
 * possible (keeps node tests FileReader-free): arrayBuffer → chunked binary
 * string → btoa. Chunked because String.fromCharCode(...bytes) overflows
 * the call stack on MB-sized files.
 * @private
 */
async function _readAsDataUrl(file) {
    const bytes = new Uint8Array(await _readArrayBuffer(file));
    let binary = '';
    const CHUNK = 0x8000;
    for (let i = 0; i < bytes.length; i += CHUNK) {
        binary += String.fromCharCode.apply(null, bytes.subarray(i, i + CHUNK));
    }
    const mime = file.type || 'application/octet-stream';
    return `data:${mime};base64,${btoa(binary)}`;
}

/**
 * Extracts raw text from a .docx via mammoth's prebuilt browser bundle
 * (UMD — webpack interop may wrap the export in .default).
 * @private
 */
async function _extractDocxText(file) {
    const mod = await import(/* webpackChunkName: "mammoth" */ 'mammoth/mammoth.browser.min.js');
    const mammoth = mod.default || mod;
    const arrayBuffer = await _readArrayBuffer(file);
    const result = await mammoth.extractRawText({ arrayBuffer });
    return (result && typeof result.value === 'string' ? result.value : '').trim();
}

/**
 * Extracts text from a .pdf via pdf.js legacy build (WebView2-compatible).
 * The worker script is copied to dist/ by webpack (see webpack.config.cjs);
 * workerSrc resolves against the document base so dev server and sideloaded
 * builds both find it. Items join with newlines at hasEOL boundaries —
 * good enough for prompt context, no layout fidelity attempted.
 * @private
 */
async function _extractPdfText(file) {
    const pdfjs = await import(/* webpackChunkName: "pdfjs" */ 'pdfjs-dist/legacy/build/pdf.min.mjs');
    if (pdfjs.GlobalWorkerOptions && !pdfjs.GlobalWorkerOptions.workerSrc) {
        const base = (typeof document !== 'undefined' && document.baseURI) || '';
        try {
            pdfjs.GlobalWorkerOptions.workerSrc = base
                ? new URL('pdf.worker.min.mjs', base).toString()
                : 'pdf.worker.min.mjs';
        } catch (_err) {
            pdfjs.GlobalWorkerOptions.workerSrc = 'pdf.worker.min.mjs';
        }
    }
    const loadingTask = pdfjs.getDocument({ data: await _readArrayBuffer(file) });
    try {
        const doc = await loadingTask.promise;
        const parts = [];
        for (let p = 1; p <= doc.numPages; p++) {
            const page = await doc.getPage(p);
            const reader = page.streamTextContent().getReader();
            const chunks = [];
            try {
                while (true) {
                    const { done, value } = await reader.read();
                    if (done) break;
                    for (const item of value.items) {
                        if (typeof item.str === 'string') {
                            chunks.push(item.str + (item.hasEOL ? '\n' : ''));
                        }
                    }
                }
            } finally {
                try {
                    reader.releaseLock();
                } catch (_err) {
                    // Cleanup must not replace the extraction result or error.
                }
            }
            parts.push(chunks.join(''));
        }
        return parts.join('\n\n').trim();
    } finally {
        try {
            await loadingTask.destroy();
        } catch (_err) {
            // Cleanup must not replace the extraction result or error.
        }
    }
}
