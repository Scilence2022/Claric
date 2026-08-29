/**
 * Table Creation Protocol
 *
 * Pure helpers for creating a native Word table. The protocol keeps table
 * structure explicit: cells are represented by a rectangular string matrix,
 * while placement and table options live in a small validated allowlist.
 *
 * This module has no DOM, Office.js, or network dependencies. The caller is
 * responsible for mapping the validated spec to the Word API.
 *
 * @module table-ops
 */

import { balancedJsonCandidates, stripTrailingCommas } from './json-utils.js';

/** Conservative limits for model output and Word/taskpane usability. */
export const TABLE_CREATION_LIMITS = Object.freeze({
    MAX_ROWS: 50,
    MAX_COLUMNS: 20,
    MAX_CELLS: 500,
    MAX_CELL_CHARS: 2000,
    MAX_TOTAL_CHARS: 50000,
    MAX_RESPONSE_CHARS: 100000,
});

/** Placement values understood by the Word-bound table creator. */
export const TABLE_POSITIONS = Object.freeze(['start', 'end', 'before', 'after']);

/** Native table styles accepted by this protocol. */
export const TABLE_STYLES = Object.freeze(['tableGrid']);

/** Defaults applied when optional model fields are omitted. */
export const DEFAULT_TABLE_CREATION_OPTIONS = Object.freeze({
    position: 'end',
    headerRowCount: 0,
    style: 'tableGrid',
    autoFit: true,
});

const SPEC_KEYS = new Set(['rows', 'position', 'headerRowCount', 'style', 'autoFit']);
const CHINESE_NUMBER_TOKEN = '[0-9零〇一二两三四五六七八九十百千]+';
const CHINESE_DIGITS = Object.freeze({
    零: 0,
    〇: 0,
    一: 1,
    二: 2,
    两: 2,
    三: 3,
    四: 4,
    五: 5,
    六: 6,
    七: 7,
    八: 8,
    九: 9,
});
const CHINESE_UNITS = Object.freeze({ 十: 10, 百: 100, 千: 1000 });

/**
 * @typedef {object} TableCreationSpec
 * @property {string[][]} rows - Non-empty rectangular plain-text cell matrix
 * @property {'start'|'end'|'before'|'after'} position - Insertion position;
 *   before/after are relative to the caller's active range
 * @property {number} headerRowCount - Leading rows treated as headers
 * @property {'tableGrid'} style - Allowlisted native table style identifier
 * @property {boolean} autoFit - Whether the caller should auto-fit the table
 */

/**
 * @typedef {object} TableValidationIssue
 * @property {string} code - Stable machine-readable issue code
 * @property {string} path - JSON-style path to the invalid value
 * @property {string} message - Human-readable detail suitable for the UI
 */

/**
 * @typedef {object} TableValidationResult
 * @property {boolean} ok
 * @property {TableCreationSpec|null} spec - Sanitized spec, or null on any error
 * @property {TableValidationIssue[]} errors
 * @property {TableValidationIssue[]} warnings
 */

/**
 * Deterministically infers an empty native-table spec from explicit dimensions.
 * This helper is intended for requests already known by the caller to ask for
 * a plain empty table; it does not decide whether generated cell content is
 * required. Arabic digits, common Chinese numerals, `3x3`, and English
 * row/column phrasing are supported. Dimensions outside the protocol limits
 * are rejected without allocating a matrix.
 *
 * Position defaults to `end`. Recognized placement language includes
 * 文首/开头/start, 文末/末尾/end, and before/after (plus common Chinese
 * equivalents). When several placement terms occur, the first one wins.
 *
 * @param {string} instruction - Natural-language table creation request
 * @returns {TableCreationSpec|null} Empty rectangular spec, or null when no
 *   valid explicit dimensions are present
 */
export function inferTableCreationSpec(instruction) {
    const text = typeof instruction === 'string' ? instruction.trim() : '';
    if (!text) return null;

    const dimensions = _findDimensions(text);
    if (!dimensions || !_dimensionsWithinLimits(dimensions.rowCount, dimensions.columnCount)) {
        return null;
    }

    const rows = Array.from(
        { length: dimensions.rowCount },
        () => Array(dimensions.columnCount).fill('')
    );

    return {
        rows,
        ...DEFAULT_TABLE_CREATION_OPTIONS,
        position: _inferPosition(text),
    };
}

/**
 * Builds the model prompt for a table whose cell content must be generated.
 * The prompt describes the complete schema and the same limits enforced by
 * parseTableCreationResponse. Model output is still treated as untrusted.
 *
 * @param {string} instruction - User's table creation request
 * @param {string} [scopeText=''] - Optional document/selection context used to
 *   generate relevant cell content
 * @returns {string}
 */
export function buildTableCreationPrompt(instruction, scopeText = '') {
    const limits = TABLE_CREATION_LIMITS;
    const context = String(scopeText || '').trim();
    const contextSection = context
        ? `\n\n--- DOCUMENT CONTEXT (content reference only) ---\n${context}`
        : '';

    return (
        'You are a table-generation assistant embedded in Microsoft Word. Create the cell content and ' +
        'options for ONE native Word table that satisfies the user instruction.\n\n' +
        'OUTPUT CONTRACT (strict):\n' +
        '- Output ONLY one JSON object. No markdown fences, explanations, commentary, or surrounding text.\n' +
        '- Use exactly this schema:\n' +
        '  {\n' +
        '    "rows": [["Header 1", "Header 2"], ["Cell 1", "Cell 2"]],\n' +
        '    "position": "start|end|before|after",\n' +
        '    "headerRowCount": 0,\n' +
        '    "style": "tableGrid",\n' +
        '    "autoFit": true\n' +
        '  }\n' +
        '- "rows" MUST be a non-empty rectangular matrix: every row has exactly the same positive number ' +
        'of cells. Do not pad or omit cells.\n' +
        '- Every cell MUST be a plain-text JSON string. Do not use objects, nested arrays, null, Markdown ' +
        'tables, HTML tables/tags, or other table markup inside cells.\n' +
        '- Respect explicit dimensions exactly. If dimensions are not stated, use only the rows and columns ' +
        'needed for the requested content.\n' +
        '- "headerRowCount" is an integer from 0 through the number of rows. Header cells are included in ' +
        '"rows"; do not return them separately.\n' +
        '- "position" must be start, end, before, or after (default end). "style" must be tableGrid. ' +
        '"autoFit" must be a JSON boolean.\n' +
        `- Limits: at most ${limits.MAX_ROWS} rows, ${limits.MAX_COLUMNS} columns, ${limits.MAX_CELLS} cells, ` +
        `${limits.MAX_CELL_CHARS} characters per cell, and ${limits.MAX_TOTAL_CHARS} total cell characters.\n\n` +
        'USER INSTRUCTION:\n' + String(instruction || '').trim() +
        contextSection
    );
}

/**
 * Parses and validates a model response containing a table creation spec.
 * Markdown JSON fences and prose surrounding a JSON object are tolerated.
 * Malformed/non-object responses and any semantic validation error reject the
 * complete spec; rows are never padded, truncated, or partially accepted.
 *
 * Number and boolean cell primitives are deliberately coerced to strings to
 * accommodate numeric table data. Each coercion produces a structured warning.
 * Objects, null, and arrays in cells are rejected.
 *
 * @param {string} raw - Raw model response
 * @returns {TableValidationResult}
 */
export function parseTableCreationResponse(raw) {
    if (typeof raw !== 'string') {
        return _failedResult(_issue(
            'RESPONSE_NOT_STRING',
            '$',
            'Table creation response must be a string containing one JSON object'
        ));
    }
    if (raw.length > TABLE_CREATION_LIMITS.MAX_RESPONSE_CHARS) {
        return _failedResult(_issue(
            'RESPONSE_TOO_LARGE',
            '$',
            `Table creation response exceeds ${TABLE_CREATION_LIMITS.MAX_RESPONSE_CHARS} characters`
        ));
    }

    const extracted = _extractJsonObject(raw);
    if (!extracted.value) return _failedResult(extracted.error);
    return validateTableCreationSpec(extracted.value);
}

/**
 * Validates and sanitizes an already-parsed table creation candidate.
 * Missing options receive DEFAULT_TABLE_CREATION_OPTIONS. Unknown top-level
 * fields are ignored with warnings; unsupported values and invalid field types
 * are errors. Any error makes `spec` null, preserving all-or-nothing semantics.
 *
 * @param {*} candidate - Parsed JSON candidate
 * @returns {TableValidationResult}
 */
export function validateTableCreationSpec(candidate) {
    const errors = [];
    const warnings = [];

    if (!_isObject(candidate)) {
        errors.push(_issue('SPEC_NOT_OBJECT', '$', 'Table creation spec must be a JSON object'));
        return _validationResult(null, errors, warnings);
    }

    for (const key of Object.keys(candidate)) {
        if (!SPEC_KEYS.has(key)) {
            warnings.push(_issue(
                'UNKNOWN_PROPERTY',
                key,
                `Unknown table creation property "${key}" was ignored`
            ));
        }
    }

    const position = candidate.position === undefined
        ? DEFAULT_TABLE_CREATION_OPTIONS.position
        : candidate.position;
    if (!TABLE_POSITIONS.includes(position)) {
        errors.push(_issue(
            'UNSUPPORTED_POSITION',
            'position',
            `Unsupported table position "${String(position)}"; expected ${TABLE_POSITIONS.join('|')}`
        ));
    }

    const style = candidate.style === undefined
        ? DEFAULT_TABLE_CREATION_OPTIONS.style
        : candidate.style;
    if (!TABLE_STYLES.includes(style)) {
        errors.push(_issue(
            'UNSUPPORTED_STYLE',
            'style',
            `Unsupported table style "${String(style)}"; expected ${TABLE_STYLES.join('|')}`
        ));
    }

    const autoFit = candidate.autoFit === undefined
        ? DEFAULT_TABLE_CREATION_OPTIONS.autoFit
        : candidate.autoFit;
    if (typeof autoFit !== 'boolean') {
        errors.push(_issue(
            'AUTO_FIT_NOT_BOOLEAN',
            'autoFit',
            '"autoFit" must be a JSON boolean'
        ));
    }

    const headerRowCount = candidate.headerRowCount === undefined
        ? DEFAULT_TABLE_CREATION_OPTIONS.headerRowCount
        : candidate.headerRowCount;
    if (!Number.isInteger(headerRowCount) || headerRowCount < 0) {
        errors.push(_issue(
            'INVALID_HEADER_ROW_COUNT',
            'headerRowCount',
            '"headerRowCount" must be a non-negative integer'
        ));
    }

    const matrix = _validateRows(candidate.rows, errors, warnings);
    if (matrix.rowCount !== null && Number.isInteger(headerRowCount) && headerRowCount > matrix.rowCount) {
        errors.push(_issue(
            'HEADER_ROW_COUNT_OUT_OF_RANGE',
            'headerRowCount',
            `"headerRowCount" (${headerRowCount}) exceeds the table row count (${matrix.rowCount})`
        ));
    }

    const spec = errors.length === 0
        ? { rows: matrix.rows, position, headerRowCount, style, autoFit }
        : null;
    return _validationResult(spec, errors, warnings);
}

/** @private */
function _validateRows(rawRows, errors, warnings) {
    if (!Array.isArray(rawRows)) {
        errors.push(_issue('ROWS_NOT_ARRAY', 'rows', '"rows" must be a rectangular array of arrays'));
        return { rows: [], rowCount: null };
    }

    const rowCount = rawRows.length;
    if (rowCount < 1 || rowCount > TABLE_CREATION_LIMITS.MAX_ROWS) {
        errors.push(_issue(
            'ROW_COUNT_OUT_OF_RANGE',
            'rows',
            `Table must contain 1-${TABLE_CREATION_LIMITS.MAX_ROWS} rows (got ${rowCount})`
        ));
    }

    let expectedColumns = null;
    let totalCells = 0;
    let totalChars = 0;
    const rows = [];

    for (let rowIndex = 0; rowIndex < rawRows.length; rowIndex += 1) {
        const rawRow = rawRows[rowIndex];
        const rowPath = `rows[${rowIndex}]`;
        if (!Array.isArray(rawRow)) {
            errors.push(_issue('ROW_NOT_ARRAY', rowPath, `${rowPath} must be an array`));
            rows.push([]);
            continue;
        }

        if (expectedColumns === null) expectedColumns = rawRow.length;
        if (rawRow.length !== expectedColumns) {
            errors.push(_issue(
                'RAGGED_ROWS',
                rowPath,
                `${rowPath} has ${rawRow.length} cells; expected ${expectedColumns}`
            ));
        }
        totalCells += rawRow.length;

        const row = [];
        for (let columnIndex = 0; columnIndex < rawRow.length; columnIndex += 1) {
            const cellPath = `${rowPath}[${columnIndex}]`;
            const normalized = _normalizeCell(rawRow[columnIndex], cellPath, errors, warnings);
            row.push(normalized === null ? '' : normalized);
            if (normalized !== null) totalChars += normalized.length;
        }
        rows.push(row);
    }

    if (expectedColumns !== null && (
        expectedColumns < 1 || expectedColumns > TABLE_CREATION_LIMITS.MAX_COLUMNS
    )) {
        errors.push(_issue(
            'COLUMN_COUNT_OUT_OF_RANGE',
            'rows',
            `Table must contain 1-${TABLE_CREATION_LIMITS.MAX_COLUMNS} columns (got ${expectedColumns})`
        ));
    }
    if (totalCells > TABLE_CREATION_LIMITS.MAX_CELLS) {
        errors.push(_issue(
            'CELL_COUNT_OUT_OF_RANGE',
            'rows',
            `Table contains ${totalCells} cells; maximum is ${TABLE_CREATION_LIMITS.MAX_CELLS}`
        ));
    }
    if (totalChars > TABLE_CREATION_LIMITS.MAX_TOTAL_CHARS) {
        errors.push(_issue(
            'TOTAL_TEXT_TOO_LARGE',
            'rows',
            `Table cell text contains ${totalChars} characters; maximum is ${TABLE_CREATION_LIMITS.MAX_TOTAL_CHARS}`
        ));
    }

    return { rows, rowCount };
}

/** @private */
function _normalizeCell(value, path, errors, warnings) {
    let text;
    if (typeof value === 'string') {
        text = value;
    } else if (typeof value === 'number' && Number.isFinite(value)) {
        text = String(value);
        warnings.push(_issue(
            'CELL_PRIMITIVE_COERCED',
            path,
            `${path} numeric value was converted to plain text`
        ));
    } else if (typeof value === 'boolean') {
        text = String(value);
        warnings.push(_issue(
            'CELL_PRIMITIVE_COERCED',
            path,
            `${path} boolean value was converted to plain text`
        ));
    } else {
        errors.push(_issue(
            'CELL_NOT_TEXT',
            path,
            `${path} must be a string, number, or boolean primitive`
        ));
        return null;
    }

    if (text.length > TABLE_CREATION_LIMITS.MAX_CELL_CHARS) {
        errors.push(_issue(
            'CELL_TEXT_TOO_LARGE',
            path,
            `${path} contains ${text.length} characters; maximum is ${TABLE_CREATION_LIMITS.MAX_CELL_CHARS}`
        ));
    }
    return text;
}

/** @private */
function _findDimensions(text) {
    const chineseRowsFirst = new RegExp(
        `(${CHINESE_NUMBER_TOKEN})\\s*行\\s*(?:[xX×*、,，]|和|与|及)?\\s*(${CHINESE_NUMBER_TOKEN})\\s*列`
    );
    const chineseColumnsFirst = new RegExp(
        `(${CHINESE_NUMBER_TOKEN})\\s*列\\s*(?:[xX×*、,，]|和|与|及)?\\s*(${CHINESE_NUMBER_TOKEN})\\s*行`
    );

    let match = text.match(chineseRowsFirst);
    if (match) return _dimensionsFromTokens(match[1], match[2]);

    match = text.match(chineseColumnsFirst);
    if (match) return _dimensionsFromTokens(match[2], match[1]);

    match = text.match(/(\d+)\s*rows?\s*(?:(?:and|by|x|×)\s*)?(\d+)\s*(?:columns?|cols?)\b/i);
    if (match) return _dimensionsFromTokens(match[1], match[2]);

    match = text.match(/(\d+)\s*(?:columns?|cols?)\s*(?:(?:and|by|x|×)\s*)?(\d+)\s*rows?\b/i);
    if (match) return _dimensionsFromTokens(match[2], match[1]);

    match = text.match(/(?:^|[^\d])(\d+)\s*[xX×]\s*(\d+)(?!\d)/);
    if (match) return _dimensionsFromTokens(match[1], match[2]);

    return null;
}

/** @private */
function _dimensionsFromTokens(rowToken, columnToken) {
    const rowCount = _parseDimensionToken(rowToken);
    const columnCount = _parseDimensionToken(columnToken);
    if (!Number.isSafeInteger(rowCount) || !Number.isSafeInteger(columnCount)) return null;
    return { rowCount, columnCount };
}

/** @private */
function _parseDimensionToken(token) {
    if (/^\d+$/.test(token)) return Number(token);
    if (!new RegExp(`^${CHINESE_NUMBER_TOKEN}$`).test(token)) return NaN;

    const chars = [...token];
    const hasUnit = chars.some((char) => CHINESE_UNITS[char] !== undefined);
    if (!hasUnit) {
        const digits = chars.map((char) => CHINESE_DIGITS[char]);
        if (digits.some((digit) => digit === undefined)) return NaN;
        return Number(digits.join(''));
    }

    let total = 0;
    let currentDigit = 0;
    for (const char of chars) {
        if (CHINESE_DIGITS[char] !== undefined) {
            currentDigit = CHINESE_DIGITS[char];
            continue;
        }
        const unit = CHINESE_UNITS[char];
        if (!unit) return NaN;
        total += (currentDigit || 1) * unit;
        currentDigit = 0;
    }
    return total + currentDigit;
}

/** @private */
function _dimensionsWithinLimits(rowCount, columnCount) {
    return Number.isInteger(rowCount)
        && Number.isInteger(columnCount)
        && rowCount >= 1
        && rowCount <= TABLE_CREATION_LIMITS.MAX_ROWS
        && columnCount >= 1
        && columnCount <= TABLE_CREATION_LIMITS.MAX_COLUMNS
        && rowCount * columnCount <= TABLE_CREATION_LIMITS.MAX_CELLS;
}

/** @private */
function _inferPosition(text) {
    /** @type {Array<{position: 'start'|'end'|'before'|'after', pattern: RegExp}>} */
    const patterns = [
        { position: 'start', pattern: /文首|开头|起始位置|最前面|\b(?:start|beginning)\b/i },
        { position: 'end', pattern: /文末|末尾|结尾|\b(?:end|ending)\b/i },
        { position: 'before', pattern: /之前|前面|\bbefore\b/i },
        { position: 'after', pattern: /之后|后面|\bafter\b/i },
    ];

    let selected = null;
    for (const entry of patterns) {
        const index = text.search(entry.pattern);
        if (index !== -1 && (!selected || index < selected.index)) {
            selected = { position: entry.position, index };
        }
    }
    return selected ? selected.position : /** @type {'end'} */ (DEFAULT_TABLE_CREATION_OPTIONS.position);
}

/**
 * Table-creation variant of json-utils' extractJsonObject. NOT delegated to
 * json-utils on purpose: the caller triages three distinct failure shapes
 * (NO_JSON_OBJECT / MALFORMED_JSON / non-object reply) into typed issues
 * that the table-ops spec asserts on, while json-utils throws a single
 * undifferentiated Error. If json-utils ever exposes that triage, this
 * wrapper can collapse onto it.
 *
 * @private
 */
function _extractJsonObject(raw) {
    const text = raw.trim();
    if (!text) {
        return {
            value: null,
            error: _issue('NO_JSON_OBJECT', '$', 'Table creation response is empty'),
        };
    }

    const sources = [];
    const fencePattern = /```(?:json)?\s*([\s\S]*?)```/gi;
    let fence;
    while ((fence = fencePattern.exec(text)) !== null) sources.push(fence[1].trim());
    sources.push(text);

    let sawNonObjectJson = false;
    let sawContainerMarker = false;
    let parseMessage = '';

    for (const source of sources) {
        if (!source) continue;
        try {
            const parsed = JSON.parse(source);
            if (_isObject(parsed)) return { value: parsed, error: null };
            // A valid array/primitive is a complete response, not an object
            // embedded in prose. Do not recover an object nested in an array.
            if (Array.isArray(parsed)) return _notObjectExtraction();
            sawNonObjectJson = true;
        } catch (error) {
            parseMessage = error.message;
        }

        if (source.includes('{') || source.includes('[')) sawContainerMarker = true;
        for (const candidate of balancedJsonCandidates(source)) {
            // Valid JSON parses as-is; only a failed parse retries once with
            // string-aware trailing-comma cleanup (valid JSON is never
            // rewritten, so cell text like "a, b, ]" survives intact).
            let parsedCandidate;
            try {
                parsedCandidate = JSON.parse(candidate.text);
            } catch (_error) {
                try {
                    parsedCandidate = JSON.parse(stripTrailingCommas(candidate.text));
                } catch (retryError) {
                    parseMessage = retryError.message;
                    continue;
                }
            }
            if (candidate.kind === '[' || Array.isArray(parsedCandidate)) return _notObjectExtraction();
            if (_isObject(parsedCandidate)) return { value: parsedCandidate, error: null };
            sawNonObjectJson = true;
        }
    }

    if (sawContainerMarker) {
        return {
            value: null,
            error: _issue(
                'MALFORMED_JSON',
                '$',
                `Table creation response contains malformed JSON${parseMessage ? ` (${parseMessage})` : ''}`
            ),
        };
    }
    if (sawNonObjectJson) return _notObjectExtraction();
    return {
        value: null,
        error: _issue('NO_JSON_OBJECT', '$', 'Table creation response contains no JSON object'),
    };
}

/** @private */
function _notObjectExtraction() {
    return {
        value: null,
        error: _issue('RESPONSE_NOT_OBJECT', '$', 'Table creation response JSON must be an object'),
    };
}

/** @private */
function _isObject(value) {
    return Boolean(value) && typeof value === 'object' && !Array.isArray(value);
}

/** @private */
function _issue(code, path, message) {
    return { code, path, message };
}

/** @private */
function _failedResult(error) {
    return _validationResult(null, [error], []);
}

/** @private */
function _validationResult(spec, errors, warnings) {
    return { ok: errors.length === 0, spec, errors, warnings };
}
