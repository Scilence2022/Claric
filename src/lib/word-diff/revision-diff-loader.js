/** Load native edit verification before any host reads or writes. Keeping
 * the locator off the startup path leaves room for equation-aware readers.
 * @param {Word.RequestContext} context @param {Word.Range} range
 * @param {string} before @param {string} after @param {function} [log] @param {object} [options]
 */
export async function tryRevisionDiff(context, range, before, after, log = () => {}, options = {}) {
    const engine = await import(/* webpackChunkName: "word-revision-diff" */ './revision-diff.js');
    return engine.tryRevisionDiff(context, range, before, after, log, options);
}
