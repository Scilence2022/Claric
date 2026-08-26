
import createDOMPurify from 'dompurify';

/**
 * Sanitize Module
 *
 * Shared lazy DOMPurify factory for modules that sanitize LLM-generated
 * markup (document-generator HTML summaries, illustration SVGs).
 *
 * @module sanitize
 */

// The dompurify default export is a ready instance when a global window
// exists at bundle time (browser/WebView2), but resolves to the factory
// under babel/jest CommonJS interop. Initialized lazily so importing this
// module in a no-DOM environment (node test specs) never touches window.
let purifierInstance = null;

/**
 * Returns the DOMPurify instance for the current environment.
 *
 * @returns {{ sanitize: (string, object) => string }}
 */
export function getPurifier() {
    if (!purifierInstance) {
        purifierInstance = typeof createDOMPurify.sanitize === 'function'
            ? createDOMPurify
            : createDOMPurify(window);
    }
    return purifierInstance;
}
