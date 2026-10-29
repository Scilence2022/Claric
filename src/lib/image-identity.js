/**
 * Builds a stable, metadata-based identity for one snapshot image. It is not
 * intended to identify two visually identical images; callers must treat
 * duplicate keys as ambiguous rather than silently choosing the first one.
 *
 * @param {object} image
 * @returns {string}
 */
export function imageIdentityKey(image = {}) {
    if (image && typeof image.fingerprint === 'string' && image.fingerprint) {
        return `fingerprint:${image.fingerprint}`;
    }
    return JSON.stringify([
        Number.isFinite(Number(image.width)) ? Number(image.width) : null,
        Number.isFinite(Number(image.height)) ? Number(image.height) : null,
        String(image.altText ?? image.altTextDescription ?? ''),
        String(image.title ?? image.altTextTitle ?? ''),
        String(image.hyperlink ?? ''),
        image.lockAspectRatio === undefined ? null : !!image.lockAspectRatio,
        image.alignment === undefined ? null : String(image.alignment).toLowerCase(),
        String(image.sourceId ?? image.svgSourceId ?? ''),
    ]);
}

