/** Executable local Word capabilities. Planning text is generated from this list. */
export const CAPABILITY_CATALOG = Object.freeze([
    { type: 'comment_management', effect: 'document', draft: 'comments', description: 'Delete all Word comment threads and replies in the document, or in an explicitly requested text selection. Use an explicit instruction such as "delete all comments" or "delete all comments in the selection". Filters by author, content or status, resolving, replying, and editing comments are unsupported. Never substitute a prose rewrite for a comment action.' },
    { type: 'document_edit', effect: 'document', draft: 'prose', description: 'Read the article, insert or integrate plain body prose at a chosen location, revise nearby body text, and format newly inserted paragraphs with bold or italic in one reviewed draft. Existing headings, tables and images are read-only.' },
    { type: 'insert', effect: 'document', draft: 'format', description: 'Add a short structural title or heading using the formatting proposal.' },
    { type: 'format', effect: 'document', draft: 'format', description: 'Format existing text or paragraphs, including font, color, alignment and heading styles, and remove verified redundant empty paragraphs within the same scope and proposal. Combine formatting and blank-line cleanup in one format task; it cannot see unapplied prose from another proposal.' },
    { type: 'edit', effect: 'document', draft: 'text', description: 'Rewrite or polish existing selected text or document sections, including cleanup of redundant spaces in body text. Preserve meaningful spaces, words and structure when only whitespace cleanup is requested. This produces a separate proposal and cannot see another unapplied proposal. Coordinate table-cell changes with table_management.' },
    { type: 'append', effect: 'document', draft: 'append', description: 'Generate new long-form text at the document end.' },
    { type: 'table', effect: 'document', draft: 'table', description: 'Create one native Word table; content and placement must be specified in its own proposal.' },
    { type: 'illustration', effect: 'document', draft: 'image', description: 'Design and insert an illustration.' },
    { type: 'image_management', effect: 'document', draft: 'image', description: 'Inspect or modify existing Word images and visible figure captions using image tools.' },
    { type: 'table_management', effect: 'document', draft: 'table', description: 'Inspect or modify existing native Word tables, including cell text/whitespace, rows, merges, font, shading and native table/row borders. Supports all document tables in one task and academic three-line tables: remove other borders and retain the top rule, header-bottom rule and bottom rule. It produces a separate proposal and cannot see another unapplied proposal.' },
    { type: 'qa', effect: 'answer', draft: null, description: 'Answer a question in chat using the document or supplied references; no Word change.' },
]);

export const CAPABILITY_BY_TYPE = new Map(CAPABILITY_CATALOG.map((item) => [item.type, item]));
export const capabilityTypeList = () => CAPABILITY_CATALOG.map((item) => item.type);
export const capabilityPrompt = () => CAPABILITY_CATALOG.map((item) =>
    `- "${item.type}" (${item.effect}): ${item.description}`).join('\n');
