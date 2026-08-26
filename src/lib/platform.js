/**
 * Host Platform Detection
 *
 * Word behaves differently across hosts for table row-level revisions:
 * Word desktop (PC/Mac) records table row insertions/deletions as tracked
 * changes (w:ins/w:del markers in w:trPr) while Track Changes is on, but
 * Word for the web does not track table structure changes — the row edit
 * lands without a revision mark. Table-scope edits therefore differentiate
 * by platform: desktop runs row ops under trackAll; the web host applies
 * them untracked and the caller must tell the user.
 *
 * Detection uses Office.context.platform (Office.PlatformType), which the
 * shared API has exposed since Office.js 1.1 — older than every WordApi
 * feature this add-in uses, so it is always available at runtime.
 *
 * Module-level code is side-effect free (safe to import under Jest/node).
 *
 * @module platform
 */

/**
 * Reads the current Office host platform.
 *
 * @param {object} [officeObj] - Injectable Office global (tests); defaults
 *   to the real Office global when present.
 * @returns {string} Office.PlatformType value ('PC' | 'Mac' | 'OfficeOnline'
 *   | 'iOS' | 'Android' | 'Universal'), or 'unknown' when Office is absent
 *   (Jest/node) or the read fails.
 */
export function getHostPlatform(officeObj) {
    const office = officeObj !== undefined
        ? officeObj
        : (typeof Office !== 'undefined' ? Office : undefined);
    try {
        return (office && office.context && office.context.platform) || 'unknown';
    } catch (_err) {
        return 'unknown';
    }
}

/**
 * Whether the host records table row insertions/deletions as tracked
 * changes. Only Word desktop (PC/Mac) does; web, mobile, and Universal
 * hosts apply row structure edits without revision marks, so callers must
 * warn instead of promising redlines. Unknown/Universal fail closed:
 * silently untracked structural edits are worse than an untracked edit the
 * user was told about.
 *
 * @param {string} platform - Value from getHostPlatform()
 * @returns {boolean}
 */
export function supportsTrackedRowOps(platform) {
    return platform === 'PC' || platform === 'Mac';
}
