/** Executable local Word capabilities. Planning text is generated from this list. */
export const CAPABILITY_CATALOG = Object.freeze([
    { type: 'document_edit', effect: 'document', draft: 'prose', description: 'Read the article, insert or integrate plain body prose at a chosen location, revise nearby body text, and format newly inserted paragraphs with bold or italic in one reviewed draft. Existing headings, tables and images are read-only.' },
    { type: 'insert', effect: 'document', draft: 'format', description: 'Add a short structural title or heading using the formatting proposal.' },
    { type: 'format', effect: 'document', draft: 'format', description: 'Format existing text or paragraphs, including font, color, alignment and heading styles; it cannot see unapplied prose from another proposal.' },
    { type: 'edit', effect: 'document', draft: 'text', description: 'Rewrite or polish existing selected text or document sections; this produces a separate proposal and cannot see another unapplied proposal.' },
    { type: 'append', effect: 'document', draft: 'append', description: 'Generate new long-form text at the document end.' },
    { type: 'table', effect: 'document', draft: 'table', description: 'Create one native Word table; content and placement must be specified in its own proposal.' },
    { type: 'illustration', effect: 'document', draft: 'image', description: 'Design and insert an illustration.' },
    { type: 'image_management', effect: 'document', draft: 'image', description: 'Inspect or modify existing Word images and visible figure captions using image tools.' },
    { type: 'table_management', effect: 'document', draft: 'table', description: 'Inspect or modify existing native Word tables, including cells, rows, merges and styling.' },
    { type: 'qa', effect: 'answer', draft: null, description: 'Answer a question in chat using the document or supplied references; no Word change.' },
]);

export const CAPABILITY_BY_TYPE = new Map(CAPABILITY_CATALOG.map((item) => [item.type, item]));
export const capabilityTypeList = () => CAPABILITY_CATALOG.map((item) => item.type);
export const capabilityPrompt = () => CAPABILITY_CATALOG.map((item) =>
    `- "${item.type}" (${item.effect}): ${item.description}`).join('\n');
