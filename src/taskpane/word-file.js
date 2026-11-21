/** Collect a bounded native Office file snapshot and release its handle. */
const TIMEOUT_MS = 30000;
const SLICE_BYTES = 65536;

function officeCall(invoke, label, signal, onLateResult) {
    signal?.throwIfAborted();
    return new Promise((resolve, reject) => {
        let settled = false;
        const done = (error, value) => {
            if (settled) return;
            settled = true;
            clearTimeout(timer);
            signal?.removeEventListener('abort', abort);
            if (error) reject(error); else resolve(value);
        };
        const abort = () => done(new DOMException('Word file export cancelled.', 'AbortError'));
        const timer = setTimeout(() => done(new Error(`${label} export timed out.`)), TIMEOUT_MS);
        signal?.addEventListener('abort', abort, { once: true });
        try {
            invoke((result) => {
                if (settled) { onLateResult?.(result); return; }
                if (result.status !== Office.AsyncResultStatus.Succeeded) {
                    done(new Error(result.error?.message || `${label} export failed.`));
                } else done(null, result.value);
            });
        } catch (error) { done(error); }
    });
}

export async function readWordFile(type, { label, maxBytes, signal, log = () => {} }) {
    const file = await officeCall(
        (callback) => Office.context.document.getFileAsync(type, { sliceSize: SLICE_BYTES }, callback),
        label, signal, (result) => result.value?.closeAsync?.(),
    );
    try {
        signal?.throwIfAborted();
        if (!Number.isInteger(file.size) || file.size < 5 || file.size > maxBytes
            || !Number.isInteger(file.sliceCount) || file.sliceCount < 1 || file.sliceCount > Math.ceil(maxBytes / SLICE_BYTES)) {
            throw new Error(`${label} is empty, invalid or exceeds the ${maxBytes / (1024 * 1024)} MB inspection limit.`);
        }
        const bytes = new Uint8Array(file.size);
        let offset = 0;
        for (let index = 0; index < file.sliceCount; index++) {
            const slice = await officeCall((callback) => file.getSliceAsync(index, callback), label, signal);
            signal?.throwIfAborted();
            const data = slice.data instanceof ArrayBuffer ? new Uint8Array(slice.data)
                : slice.data instanceof Uint8Array ? slice.data
                    : Array.isArray(slice.data) && slice.data.every((value) => Number.isInteger(value) && value >= 0 && value <= 255)
                        ? Uint8Array.from(slice.data) : null;
            if (!data || slice.index !== index || !data.length || data.length > SLICE_BYTES || offset + data.length > bytes.length) {
                throw new Error(`Word returned an invalid ${label} slice.`);
            }
            bytes.set(data, offset);
            offset += data.length;
        }
        if (offset !== bytes.length) throw new Error(`Word did not return a complete ${label}.`);
        return bytes;
    } finally {
        try { await officeCall((callback) => file.closeAsync(callback), label); }
        catch (error) { log(`${label} handle cleanup failed: ${error.message}`, 'warning'); }
    }
}
