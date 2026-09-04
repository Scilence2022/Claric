/**
 * SVG source store: keeps the original SVG markup of Claric-designed
 * illustrations inside the document, so later edits can re-edit the vector
 * source instead of redrawing from the rasterized pixels.
 *
 * Why this exists: Word's insertInlinePictureFromBase64 rejects SVG, so
 * every illustration is rasterized to PNG on insertion and
 * getBase64ImageSrc() reads back the raster — the source is otherwise
 * unrecoverable once the proposal apply finishes.
 *
 * Mechanism: shared-API custom XML parts (Office.context.document.
 * customXmlParts — available in Word 2013+, no manifest requirement-set
 * bump). Each illustration gets one part holding its SVG; the picture
 * carries the part id in its alt-text TITLE ('claric-svg:<id>'). The title
 * field is not what screen readers speak (that is the description), so the
 * marker stays unobtrusive. Everything degrades gracefully: when the store
 * is unavailable or the marker was edited away, callers fall back to the
 * vision-grounded redesign path.
 *
 * @module svg-source-store
 */

const PART_XMLNS = 'urn:claric:illustration-source';

/** Alt-text title prefix linking a picture to its SVG source part. */
export const SVG_SOURCE_TITLE_PREFIX = 'claric-svg:';

/**
 * True when the shared custom-XML-parts API exists on this host. Runtime
 * gate instead of a manifest MinVersion bump: the add-in keeps loading on
 * WordApi 1.3 hosts and simply skips source persistence there.
 *
 * @returns {boolean}
 */
export function isSvgSourceStoreAvailable() {
    try {
        return typeof Office !== 'undefined'
            && !!Office.context
            && !!Office.context.document
            && !!Office.context.document.customXmlParts;
    } catch {
        return false;
    }
}

function _escapeXml(text) {
    return String(text)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;');
}

function _succeeded(result) {
    // The enum value of Office.AsyncResultStatus.Succeeded is the string
    // 'succeeded'; comparing the literal keeps this module usable in tests
    // where Office.AsyncResultStatus itself is not mocked.
    return !!result && result.status === 'succeeded';
}

/**
 * Saves one SVG source as a custom XML part.
 *
 * @param {string} svg - Sanitized SVG markup
 * @returns {Promise<string|null>} The new part id, or null on failure
 */
export async function saveSvgSource(svg) {
    if (!isSvgSourceStoreAvailable() || typeof svg !== 'string' || !svg.trim().startsWith('<svg')) {
        return null;
    }
    const xml = `<claricIllustration xmlns="${PART_XMLNS}" version="1">${_escapeXml(svg.trim())}</claricIllustration>`;
    try {
        return await new Promise((resolve) => {
            Office.context.document.customXmlParts.addAsync(xml, (result) => {
                resolve(_succeeded(result) && result.value ? result.value.id : null);
            });
        });
    } catch {
        return null;
    }
}

/**
 * Loads the SVG source held by a custom XML part.
 *
 * @param {string} partId - Part id from svgSourceIdFromPicture
 * @returns {Promise<string|null>} The SVG markup, or null when the part is
 *   gone, malformed, or does not hold an SVG
 */
export async function loadSvgSource(partId) {
    if (!isSvgSourceStoreAvailable() || !partId) return null;
    try {
        const xml = await new Promise((resolve) => {
            Office.context.document.customXmlParts.getByIdAsync(partId, (result) => {
                if (!_succeeded(result) || !result.value) { resolve(null); return; }
                result.value.getXmlAsync((xmlResult) => {
                    resolve(_succeeded(xmlResult) ? xmlResult.value : null);
                });
            });
        });
        if (!xml) return null;
        const doc = new DOMParser().parseFromString(String(xml), 'text/xml');
        if (doc.getElementsByTagName('parsererror').length > 0) return null;
        const root = doc.documentElement;
        if (!root || root.localName !== 'claricIllustration') return null;
        const svg = (root.textContent || '').trim();
        return svg.startsWith('<svg') ? svg : null;
    } catch {
        return null;
    }
}

/**
 * Deletes a source part (best-effort — used when its picture is replaced or
 * deleted so orphaned markup does not accumulate in the file).
 *
 * @param {string} partId
 * @returns {Promise<void>}
 */
export async function deleteSvgSource(partId) {
    if (!isSvgSourceStoreAvailable() || !partId) return;
    try {
        await new Promise((resolve) => {
            Office.context.document.customXmlParts.getByIdAsync(partId, (result) => {
                if (!_succeeded(result) || !result.value) { resolve(); return; }
                result.value.deleteAsync(() => resolve());
            });
        });
    } catch {
        // Best-effort cleanup — an orphaned part is harmless.
    }
}

/**
 * Reads the SVG-source part id off a picture proxy (altTextTitle must have
 * been loaded and synced).
 *
 * @param {object} picture - Word.InlinePicture-shaped proxy
 * @returns {string|null}
 */
export function svgSourceIdFromPicture(picture) {
    const title = picture && picture.altTextTitle;
    if (typeof title !== 'string' || !title.startsWith(SVG_SOURCE_TITLE_PREFIX)) return null;
    const id = title.slice(SVG_SOURCE_TITLE_PREFIX.length).trim();
    return id || null;
}

/**
 * High-level attach: stores the SVG and links the picture to the new part
 * via its alt-text title. The title assignment is queued on the picture's
 * Word context — the caller owns the context.sync() that commits it.
 * Never throws; returns false when the store is unavailable or the save
 * failed (callers simply proceed without persistence).
 *
 * @param {object} picture - Word.InlinePicture proxy (inside an open Word.run)
 * @param {string} svg - Sanitized SVG markup
 * @returns {Promise<boolean>} True when the title link was queued
 */
export async function attachSvgSource(picture, svg) {
    if (!picture || !isSvgSourceStoreAvailable()) return false;
    const id = await saveSvgSource(svg);
    if (!id) return false;
    try {
        picture.altTextTitle = SVG_SOURCE_TITLE_PREFIX + id;
    } catch {
        return false;
    }
    return true;
}
