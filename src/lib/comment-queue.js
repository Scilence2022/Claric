
/**
 * Comment Queue Module
 *
 * Manages the lifecycle of async comment requests: pending state tracking,
 * bookmark name generation, and Word API interactions for capturing selections
 * and inserting comments.
 *
 * @module comment-queue
 */

class CommentQueue {
    constructor(log) {
        this._pending = [];
        this._log = log || (() => {});
    }

    addPending({ id, bookmarkName, selectionPreview }) {
        this._pending.push({ id, bookmarkName, selectionPreview, status: 'pending' });
        const count = this._pending.length;
        if (count >= 5) {
            this._log(`${count} comments queued -- LLM may slow down`, 'warning');
        }
        return count;
    }

    removePending(id) {
        const idx = this._pending.findIndex(c => c.id === id);
        if (idx !== -1) this._pending.splice(idx, 1);
        return this._pending.length;
    }

    getPendingCount() {
        return this._pending.length;
    }

    getPending() {
        return [...this._pending];
    }

    hasPending(id) {
        return this._pending.some(c => c.id === id);
    }

    /**
     * Captures the current Word selection as a hidden bookmark.
     * Each call uses its own Word.run context (no shared state across async boundaries).
     *
     * @param {string} bookmarkName - Valid bookmark name (_cq prefix, alphanumeric+underscore, max 40 chars)
     * @returns {Promise<string>} The selection text at capture time
     */
    async captureSelectionAsBookmark(bookmarkName) {
        let selectionText = '';
        await Word.run(async (context) => {
            const selection = context.document.getSelection();
            selection.load('text');
            await context.sync();
            selectionText = selection.text;
            selection.insertBookmark(bookmarkName);
            await context.sync();
        });
        return selectionText;
    }

    /**
     * Retrieves a bookmarked range, inserts a comment, and cleans up the bookmark
     * in a single Word.run batch. Handles the case where the bookmark is lost
     * (user deleted the text) by returning success: false with null rangeText.
     *
     * @param {string} bookmarkName - The bookmark name to look up
     * @param {string} commentText - The comment text to insert (raw LLM response)
     * @returns {Promise<{success: boolean, rangeText: string|null}>}
     */
    async insertCommentOnBookmark(bookmarkName, commentText) {
        let result = { success: false, rangeText: null };
        await Word.run(async (context) => {
            const range = context.document.getBookmarkRangeOrNullObject(bookmarkName);
            range.load('isNullObject,text');
            await context.sync();

            if (range.isNullObject) {
                // Bookmark lost -- user deleted the text
                result = { success: false, rangeText: null };
                return;
            }

            result.rangeText = range.text;
            range.insertComment(commentText);
            context.document.deleteBookmark(bookmarkName);
            await context.sync();
            result.success = true;
        });
        return result;
    }
}

/**
 * Monotonic counter making bookmark names unique within a taskpane session.
 *
 * A timestamp+random name is NOT enough: several requests fired in the same
 * millisecond only differ by 4 base36 chars (~1.7M combinations), and
 * `Math.random().toString(36).slice(2, 6)` narrows that further because the
 * old `[^a-z0-9] -> 'a'` fold collapsed every illegal char onto one letter.
 * Collisions were observable at 100 names/ms and would make two concurrent
 * comment requests share a bookmark — the second overwrites the first's
 * captured range. The counter removes the birthday problem entirely; the
 * random suffix stays so names from different sessions (each starting at 0)
 * still do not collide inside one document.
 */
let bookmarkSeq = 0;

/**
 * Generates a unique hidden bookmark name for comment range capture.
 * Format: _cq + hex timestamp + base36 sequence + 4 random base36 chars.
 * Hidden (underscore prefix), max 40 chars, alphanumeric + underscore only.
 */
function generateBookmarkName() {
    const timestamp = Date.now().toString(16);
    const seq = (bookmarkSeq++).toString(36);
    // padStart keeps the suffix at a fixed 4 chars: toString(36) drops
    // trailing zeros, so a bare slice() can yield fewer and shrink the space.
    const random = Math.random().toString(36).slice(2, 6).padStart(4, '0');
    return `_cq${timestamp}${seq}${random}`;
}

export { CommentQueue, generateBookmarkName };
