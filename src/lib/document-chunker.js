/**
 * Document Chunker Module
 *
 * Splits a parsed document model (from document-parser.js) into
 * token-budgeted, structure-aware chunks suitable for LLM processing.
 *
 * Key behaviors:
 * - Splits at H1/H2 heading boundaries as primary break points
 * - Enforces maxTokens limit per chunk (default 12000)
 * - Excludes table paragraphs (inTable=true) entirely: a table acts as a
 *   hard chunk boundary and its cell text is never sent to the LLM —
 *   flattened cell lines invite the model to reorganize/echo them, and the
 *   paragraph alignment then pollutes the document with phantom paragraphs
 *   (tables remain editable via the selection routes)
 * - Falls back to paragraph-boundary splitting when no headings exist
 * - Includes overlap context from previous chunk for continuity
 *
 * Pure JavaScript -- no Word API dependency.
 *
 * @module document-chunker
 */

/**
 * @typedef {import('./document-parser.js').ParsedParagraph} ParsedParagraph
 * @typedef {import('./document-parser.js').DocumentModel} DocumentModel
 */

/**
 * @typedef {Object} DocumentChunk
 * @property {string} id - Unique chunk identifier (e.g., 'chunk-0', 'chunk-1')
 * @property {ParsedParagraph[]} paragraphs - Paragraphs in this chunk
 * @property {number} startIndex - First paragraph index in document
 * @property {number} endIndex - Last paragraph index in document
 * @property {number} tokenCount - Total estimated tokens in chunk
 * @property {string} sectionTitle - Nearest heading text (for logging)
 * @property {string} overlapBefore - Text from previous chunk's last paragraph(s) for context
 */

/**
 * Splits a document model into token-budgeted, structure-aware chunks.
 *
 * Algorithm:
 *   1. Iterate paragraphs in document order
 *   2. On H1/H2 heading: finalize current chunk (if >= minTokens), start new
 *   3. On maxTokens exceeded: finalize current chunk, start new
 *   4. Table paragraphs: skipped entirely; the table finalizes the current
 *      chunk as a hard boundary and sets a barrier the tiny-chunk merge
 *      never crosses (a merged bookmark range would span the table)
 *   5. After all paragraphs: finalize last chunk
 *   6. Merge tiny trailing chunks (< minTokens) into previous chunk
 *   7. Add overlap: for each chunk after first, set overlapBefore
 *   8. Assign sectionTitle, unique IDs
 *
 * @param {DocumentModel} docModel - Output from parseDocument()
 * @param {Object} [options]
 * @param {number} [options.maxTokens=12000] - Max tokens per chunk
 * @param {number} [options.minTokens=500] - Min tokens before creating a chunk
 * @param {number} [options.overlapParagraphs=1] - Paragraphs to overlap between chunks
 * @returns {DocumentChunk[]}
 */
export function chunkDocument(docModel, options = {}) {
    const {
        maxTokens = 12000,
        minTokens = 500,
        overlapParagraphs = 1
    } = options;

    const { paragraphs } = docModel;
    if (!paragraphs || paragraphs.length === 0) {
        return [];
    }

    const rawChunks = [];
    let currentParas = [];
    let currentTokens = 0;
    // Set when a skipped table run stands immediately before the next chunk:
    // the trailing-chunk merge must never cross it (the merged chunk's
    // bookmark range would physically span the table).
    let tableBarrier = false;

    function finalizeCurrentChunk() {
        if (currentParas.length > 0) {
            rawChunks.push({
                paragraphs: currentParas,
                tokenCount: currentTokens,
                barrierBefore: tableBarrier
            });
            currentParas = [];
            currentTokens = 0;
            tableBarrier = false;
        }
    }

    let i = 0;
    while (i < paragraphs.length) {
        const para = paragraphs[i];

        // Table paragraphs never enter amendment chunks. Cell text flattened
        // into prose lines invites the model to reorganize or echo it, and
        // the paragraph alignment then pollutes the document (phantom
        // paragraphs from inserts) or risks striking cell content. The table
        // instead splits chunks as a hard boundary; tables remain editable
        // through the selection routes (table patch / mixed selection).
        if (para.inTable) {
            finalizeCurrentChunk();
            while (i < paragraphs.length && paragraphs[i].inTable) i++;
            tableBarrier = true;
            continue;
        }

        // H1/H2 heading starts a new chunk (if current chunk has content)
        if ((para.headingLevel === 1 || para.headingLevel === 2) && currentParas.length > 0) {
            if (currentTokens >= minTokens) {
                finalizeCurrentChunk();
            }
        }

        // Would this paragraph push us over the limit?
        if (currentTokens + para.tokenEstimate > maxTokens && currentParas.length > 0) {
            if (currentTokens >= minTokens) {
                finalizeCurrentChunk();
            }
        }

        currentParas.push(para);
        currentTokens += para.tokenEstimate;
        i++;
    }

    // Don't forget the last chunk
    finalizeCurrentChunk();

    // Merge tiny trailing chunk (below minTokens) into previous chunk —
    // unless a skipped table stands between them (barrier): merging would
    // give the combined chunk a bookmark range that spans the table.
    if (rawChunks.length > 1) {
        const lastChunk = rawChunks[rawChunks.length - 1];
        if (lastChunk.tokenCount < minTokens && !lastChunk.barrierBefore) {
            const prevChunk = rawChunks[rawChunks.length - 2];
            prevChunk.paragraphs.push(...lastChunk.paragraphs);
            prevChunk.tokenCount += lastChunk.tokenCount;
            rawChunks.pop();
        }
    }

    // Build final chunk objects with metadata
    const chunks = rawChunks.map((raw, idx) => {
        const firstPara = raw.paragraphs[0];
        const lastPara = raw.paragraphs[raw.paragraphs.length - 1];

        // Find nearest heading in this chunk for sectionTitle
        const headingPara = raw.paragraphs.find(p => p.headingLevel > 0);
        const sectionTitle = headingPara ? headingPara.text : '';

        // Build overlap from previous chunk
        let overlapBefore = '';
        if (idx > 0) {
            const prevParas = rawChunks[idx - 1].paragraphs;
            const overlapCount = Math.min(overlapParagraphs, prevParas.length);
            const overlapParas = prevParas.slice(prevParas.length - overlapCount);
            overlapBefore = overlapParas.map(p => p.text).join('\n');
        }

        return {
            id: `chunk-${idx}`,
            paragraphs: raw.paragraphs,
            startIndex: firstPara.index,
            endIndex: lastPara.index,
            tokenCount: raw.tokenCount,
            sectionTitle,
            overlapBefore
        };
    });

    return chunks;
}
