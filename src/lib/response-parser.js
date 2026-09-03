/**
 * Response Parser Module
 *
 * Parses LLM responses that contain delimited sections for amendment and comment.
 * When the LLM is instructed to return both an amendment and a comment, it uses
 * ===AMENDMENT=== and ===COMMENT=== delimiters. This module extracts those sections.
 *
 * Also provides a fallback classification prompt builder for cases where the LLM
 * doesn't follow the delimiter format.
 *
 * Owns BOTH ends of the delimiter protocol: the markers the prompt asks the
 * model to emit (AMENDMENT_MARKER / COMMENT_MARKER, used by orchestrator when
 * composing chunk messages) and the parsing of them here. Keeping the literals
 * in one module means changing a delimiter cannot desync the writer from the
 * reader.
 *
 * @module response-parser
 */

/** Delimiter the model wraps its rewritten text in. */
export const AMENDMENT_MARKER = '===AMENDMENT===';

/** Delimiter the model wraps its comment in. */
export const COMMENT_MARKER = '===COMMENT===';

/**
 * Input-framing markers the orchestrator wraps chunk text in. Listed here
 * (next to the response delimiters) because they share one failure mode:
 * document text that reproduces a marker.
 */
export const FRAMING_MARKERS = Object.freeze([
    '[AMEND THIS TEXT]',
    '[END TEXT]',
    '[CONTEXT - DO NOT AMEND]',
    '[END CONTEXT]',
]);

/**
 * Matches every protocol marker — response delimiters and input framing —
 * wherever it appears, tolerating extra `=` runs and internal whitespace so
 * near-miss reproductions are caught too.
 */
const PROTOCOL_MARKER_RE = new RegExp(
    [
        '={2,}\\s*(?:AMENDMENT|COMMENT)\\s*={2,}',
        '\\[\\s*AMEND\\s+THIS\\s+TEXT\\s*\\]',
        '\\[\\s*END\\s+TEXT\\s*\\]',
        '\\[\\s*CONTEXT\\s*-\\s*DO\\s+NOT\\s+AMEND\\s*\\]',
        '\\[\\s*END\\s+CONTEXT\\s*\\]',
    ].join('|'),
    'gi'
);

/**
 * Matches the exact marker forms produced by defangProtocolMarkers. The
 * optional zero-width space is deliberately limited to the known marker words;
 * arbitrary zero-width characters in user text must remain untouched.
 */
const DEFANGED_PROTOCOL_MARKER_RE = new RegExp(
    [
        '={2,}\\s*A\\u200b?MENDMENT\\s*={2,}',
        '={2,}\\s*C\\u200b?OMMENT\\s*={2,}',
        '\\[\\s*A\\u200b?MEND\\s+THIS\\s+TEXT\\s*\\]',
        '\\[\\s*E\\u200b?ND\\s+TEXT\\s*\\]',
        '\\[\\s*C\\u200b?ONTEXT\\s*-\\s*DO\\s+NOT\\s+AMEND\\s*\\]',
        '\\[\\s*E\\u200b?ND\\s+CONTEXT\\s*\\]',
    ].join('|'),
    'gi'
);

/**
 * Restores protocol markers that were defanged before being sent to a model.
 * Only the known marker shapes are changed; unrelated zero-width characters
 * in document text are preserved.
 *
 * @param {string} text - Model output that may echo defanged markers
 * @returns {string} Text with known defanged markers restored
 */
export function restoreProtocolMarkers(text) {
    if (typeof text !== 'string' || !text) return '';
    return text.replace(DEFANGED_PROTOCOL_MARKER_RE, (match) => match.replace(/\u200b/g, ''));
}

/**
 * Neutralizes protocol markers inside untrusted text before it is framed into
 * a prompt.
 *
 * Document text is interpolated into `[AMEND THIS TEXT]...[END TEXT]` and the
 * model is told to answer with `===AMENDMENT===` / `===COMMENT===`. Any
 * document that happens to contain those strings — a contract template, a
 * spec quoting this add-in's own protocol, or text a previous run wrote back —
 * could close the framing early and make the remainder of the document read as
 * instructions, or make `parseDelimitedResponse` treat body text as a comment
 * to insert into the document. Mirrors the fence-marker stripping that
 * context-extractor already applies to its reference block.
 *
 * Markers are replaced (not deleted) so the text stays readable and keeps its
 * character positions roughly intact for diffing.
 *
 * @param {string} text - Untrusted document-derived text
 * @returns {string} Text with protocol markers neutralized
 */
export function defangProtocolMarkers(text) {
    // Non-strings and empty input carry no markers; normalizing to '' keeps the
    // declared string return type honest for callers that interpolate it.
    if (typeof text !== 'string' || !text) return '';
    return text.replace(PROTOCOL_MARKER_RE, (match) => {
        // The zero-width space goes inside the marker's ALPHABETIC core, not
        // after its first character. Inserting after char 0 is not enough for
        // the `=`-delimited markers: the run length is variable, so defanging
        // `====AMENDMENT====` as `=<zwsp>===AMENDMENT====` still leaves a
        // complete `===AMENDMENT===` in the tail, and the parser would split
        // on it exactly as if nothing had been done. Splitting the word itself
        // ('A<zwsp>MENDMENT') cannot be rebuilt by any surrounding `=` run,
        // and works the same way for the bracketed framing markers.
        const letterIdx = match.search(/[A-Za-z]/);
        const cut = letterIdx === -1 ? 1 : letterIdx + 1;
        return `${match.slice(0, cut)}\u200b${match.slice(cut)}`;
    });
}

/**
 * Parses a delimited LLM response into amendment and comment sections.
 *
 * Looks for ===AMENDMENT=== and ===COMMENT=== markers. Extracts text between/after
 * them. If neither delimiter is found, returns nulls with the raw response.
 *
 * @param {string} responseText - The raw LLM response
 * @returns {{ amendment: string|null, comment: string|null, raw: string }}
 */
export function parseDelimitedResponse(responseText) {
    // Transport adapters normally return a string, but normalize malformed or
    // empty values here so one bad backend payload cannot crash the parser
    // before the caller can place the chunk on its retry path.
    const raw = typeof responseText === 'string' ? responseText : String(responseText || '');
    const amendmentMarker = AMENDMENT_MARKER;
    const commentMarker = COMMENT_MARKER;

    const amendmentIdx = raw.indexOf(amendmentMarker);
    const commentIdx = raw.indexOf(commentMarker);

    // Neither delimiter found
    if (amendmentIdx === -1 && commentIdx === -1) {
        return { amendment: null, comment: null, raw };
    }

    let amendment = null;
    let comment = null;

    if (amendmentIdx !== -1) {
        const afterAmendment = raw.substring(amendmentIdx + amendmentMarker.length);
        if (commentIdx !== -1 && commentIdx > amendmentIdx) {
            // Both markers present: amendment is between them
            amendment = raw.substring(
                amendmentIdx + amendmentMarker.length,
                commentIdx
            ).trim();
        } else {
            // Only amendment marker: everything after it
            amendment = afterAmendment.trim();
        }
    }

    if (commentIdx !== -1) {
        // Mirror the amendment branch. When the model emits ===COMMENT===
        // FIRST, the comment ends where ===AMENDMENT=== begins; taking
        // everything after the marker unconditionally made the comment
        // swallow the ===AMENDMENT=== marker plus the entire rewritten body,
        // and that whole string was then inserted into the document as a
        // Word comment.
        if (amendmentIdx !== -1 && amendmentIdx > commentIdx) {
            comment = raw.substring(
                commentIdx + commentMarker.length,
                amendmentIdx
            ).trim();
        } else {
            comment = raw.substring(commentIdx + commentMarker.length).trim();
        }
    }

    return { amendment: amendment || null, comment: comment || null, raw };
}

/**
 * Builds a fallback classification prompt for when the LLM response
 * doesn't contain the expected delimiters.
 *
 * @param {string} rawResponse - The original LLM response without delimiters
 * @param {string} originalSelection - The original selected text from the document
 * @returns {Array<{role: string, content: string}>} Messages array for chat completions
 */
export function buildFallbackClassificationPrompt(rawResponse, originalSelection) {
    return [
        {
            role: 'system',
            content: 'You are a response formatter. The following text was generated as both an amendment and a comment for a document clause. Split it into the amendment (the rewritten text) and the comment (the analysis/feedback). Use the exact delimiters shown.'
        },
        {
            role: 'user',
            // Both values are untrusted model/document text. Defang protocol
            // markers before interpolating them into the formatter request so
            // an echoed delimiter cannot become a new section boundary in the
            // fallback response or prompt the classifier to follow document
            // text as instructions.
            content: `Original clause:\n${defangProtocolMarkers(originalSelection)}\n\nLLM response to split:\n${defangProtocolMarkers(rawResponse)}\n\nReformat the response using these exact delimiters:\n===AMENDMENT===\n[The rewritten/amended text]\n===COMMENT===\n[The analysis/feedback comment]`
        }
    ];
}
