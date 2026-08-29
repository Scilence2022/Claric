/**
 * JSON Utilities
 *
 * Shared parsing helpers for the strict-JSON LLM protocols (tool loop,
 * table patch, table creation, task planner, format ops). Each layer used
 * to grow its own fence-stripping + brace-scanning extractor; this module
 * is the single implementation.
 *
 * Two invariants matter for document integrity:
 * 1. Valid JSON is never rewritten — trailing-comma cleanup only runs as a
 *    recovery attempt after a failed parse.
 * 2. Cleanup is string-aware — a naive `,\s*([}\]])` regex also fires on
 *    `, ]` sequences INSIDE string literals, silently deleting characters
 *    from cell text / tool args before they reach the document.
 *
 * Pure module — no DOM, no Word API, no network.
 *
 * @module json-utils
 */

/**
 * Removes trailing commas that appear OUTSIDE string literals, e.g.
 * `{"a":1,}` → `{"a":1}`. Commas inside JSON string values (including
 * escaped quotes) are never touched.
 *
 * @param {string} text
 * @returns {string}
 */
export function stripTrailingCommas(text) {
    const src = String(text || '');
    let out = '';
    let inString = false;
    let escaped = false;

    for (let i = 0; i < src.length; i += 1) {
        const char = src[i];
        if (inString) {
            out += char;
            if (escaped) {
                escaped = false;
            } else if (char === '\\') {
                escaped = true;
            } else if (char === '"') {
                inString = false;
            }
            continue;
        }
        if (char === '"') {
            inString = true;
            out += char;
            continue;
        }
        if (char === ',') {
            // Lookahead past whitespace: a comma directly before a closing
            // brace/bracket is a trailing comma; drop it, keep the gap.
            let j = i + 1;
            while (j < src.length && /\s/.test(src[j])) j += 1;
            if (j < src.length && (src[j] === '}' || src[j] === ']')) {
                continue;
            }
        }
        out += char;
    }
    return out;
}

/**
 * Extracts balanced `{...}` / `[...]` candidates from prose, honoring
 * string literals so braces inside strings never open a candidate.
 * Mismatched closings terminate the current candidate.
 *
 * @param {string} text
 * @returns {Array<{text: string, kind: '{'|'['}>}
 */
export function balancedJsonCandidates(text) {
    const candidates = [];
    let start = -1;
    let stack = [];
    let inString = false;
    let escaped = false;

    for (let index = 0; index < text.length; index += 1) {
        const char = text[index];
        if (start === -1) {
            if (char === '{' || char === '[') {
                start = index;
                stack = [char];
            }
            continue;
        }
        if (inString) {
            if (escaped) {
                escaped = false;
            } else if (char === '\\') {
                escaped = true;
            } else if (char === '"') {
                inString = false;
            }
            continue;
        }
        if (char === '"') {
            inString = true;
        } else if (char === '{' || char === '[') {
            stack.push(char);
        } else if (char === '}' || char === ']') {
            const opener = stack[stack.length - 1];
            const matches = (opener === '{' && char === '}') || (opener === '[' && char === ']');
            if (!matches) {
                candidates.push({ text: text.slice(start, index + 1), kind: /** @type {'{'|'['} */ (text[start]) });
                start = -1;
                stack = [];
                continue;
            }
            stack.pop();
            if (stack.length === 0) {
                candidates.push({ text: text.slice(start, index + 1), kind: /** @type {'{'|'['} */ (text[start]) });
                start = -1;
            }
        }
    }
    return candidates;
}

/** @private */
function _isObject(value) {
    return Boolean(value) && typeof value === 'object' && !Array.isArray(value);
}

/** @private */
function _tryParse(candidate) {
    try {
        return { ok: true, value: JSON.parse(candidate) };
    } catch (err) {
        return { ok: false, message: err.message };
    }
}

/**
 * Parses a JSON fragment; on failure retries once with string-aware
 * trailing-comma cleanup. Valid JSON is never rewritten.
 *
 * @private
 * @returns {{ok: boolean, value?: *, message?: string}}
 */
function _parseWithCleanup(text) {
    const direct = _tryParse(text);
    if (direct.ok) return direct;
    const cleaned = _tryParse(stripTrailingCommas(text));
    if (cleaned.ok) return cleaned;
    return direct;
}

/** @private */
function _fenceSources(text) {
    const sources = [];
    const fencePattern = /```(?:json)?\s*([\s\S]*?)```/gi;
    let fence;
    while ((fence = fencePattern.exec(text)) !== null) sources.push(fence[1].trim());
    sources.push(text);
    return sources;
}

/**
 * Extracts the first JSON OBJECT from a raw model reply, tolerating code
 * fences and surrounding prose. Scans balanced candidates (not a greedy
 * first-`{`…last-`}` slice), so prose containing a second brace pair no
 * longer poisons the slice. Trailing commas are cleaned as a recovery
 * attempt only.
 *
 * @param {string} raw
 * @param {object} [options]
 * @param {string} [options.noObjectMessage] - Error when nothing object-like exists
 * @param {string} [options.parseFailedPrefix] - Prefix for parse-failure errors
 * @returns {object} Parsed object
 * @throws {Error} When no JSON object can be located or parsed
 */
export function extractJsonObject(raw, {
    noObjectMessage = 'Reply contains no JSON object',
    parseFailedPrefix = 'JSON parse failed: ',
} = {}) {
    const text = String(raw || '').trim();
    if (!text) {
        throw new Error(noObjectMessage);
    }

    let sawArray = false;
    let sawContainerMarker = false;
    let lastFailure = '';

    for (const source of _fenceSources(text)) {
        if (!source) continue;

        const direct = _tryParse(source);
        if (direct.ok) {
            if (_isObject(direct.value)) return direct.value;
            // A valid array/primitive is a complete response, not an object
            // embedded in prose. Do not recover an object nested in an array.
            if (Array.isArray(direct.value)) sawArray = true;
        } else {
            lastFailure = direct.message;
        }

        if (source.includes('{') || source.includes('[')) sawContainerMarker = true;

        for (const candidate of balancedJsonCandidates(source)) {
            const parsed = _parseWithCleanup(candidate.text);
            if (parsed.ok) {
                if (_isObject(parsed.value)) return parsed.value;
                if (Array.isArray(parsed.value)) sawArray = true;
            } else {
                lastFailure = parsed.message;
            }
        }
    }

    if (sawArray) {
        throw new Error(`${parseFailedPrefix}not an object`);
    }
    if (sawContainerMarker) {
        throw new Error(parseFailedPrefix + (lastFailure || 'malformed JSON'));
    }
    throw new Error(noObjectMessage);
}

/**
 * Extracts the first JSON ARRAY from a raw model reply, tolerating code
 * fences and surrounding prose (same recovery rules as extractJsonObject).
 *
 * @param {string} raw
 * @returns {{value: Array<*>, error: string|null}} `value` is null when no
 *   array could be extracted; `error` then explains why.
 */
export function extractJsonArray(raw) {
    const text = String(raw || '').trim();
    if (!text) {
        return { value: null, error: 'no JSON array found in the model response' };
    }

    let sawContainerMarker = false;
    let sawNonArrayJson = false;
    let lastFailure = '';

    for (const source of _fenceSources(text)) {
        if (!source) continue;

        const direct = _tryParse(source);
        if (direct.ok) {
            if (Array.isArray(direct.value)) return { value: direct.value, error: null };
            // A complete object/primitive response is not an array embedded
            // in prose; do not fish arrays out of an object's fields.
            sawNonArrayJson = true;
        } else {
            lastFailure = direct.message;
        }

        if (source.includes('{') || source.includes('[')) sawContainerMarker = true;

        for (const candidate of balancedJsonCandidates(source)) {
            if (candidate.kind !== '[') continue;
            const parsed = _parseWithCleanup(candidate.text);
            if (parsed.ok) {
                if (Array.isArray(parsed.value)) return { value: parsed.value, error: null };
                sawNonArrayJson = true;
            } else {
                lastFailure = parsed.message;
            }
        }
    }

    if (sawNonArrayJson) {
        return { value: null, error: 'no JSON array found in the model response' };
    }
    if (sawContainerMarker) {
        return { value: null, error: `response is not valid JSON (${lastFailure || 'malformed JSON'})` };
    }
    return { value: null, error: 'no JSON array found in the model response' };
}
