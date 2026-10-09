/** @jest-environment jsdom */
const { describeFormatParagraph: describe, resolveFormatParagraphs: resolve,
    requestsBodyFormatting, explicitBodyAlignmentOp } = require('../src/lib/format-targets.js');
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const prose = 'The model observes the document structure before it proposes an edit. This paragraph describes the complete mechanism and its limitations.';
const xml = (text, properties = '', extra = '') => `<w:p xmlns:w="${W}"><w:pPr>${properties}</w:pPr><w:r><w:t>${text}</w:t></w:r>${extra}</w:p>`;
const paragraph = (options = {}) => describe({ index: 0, text: prose, style: 'Normal', styleBuiltIn: 'Normal',
    ooxml: xml(prose), inTable: false, withinScope: true, ...options });

test('localized Normal and native body styles identify prose without guessing localized names', () => {
    const native = paragraph({ style: '正文' });
    expect(native).toMatchObject({ id: 'p1', index: 0, text: prose, role: 'body', eligible: true, verified: true });
    expect(resolve([native], { paragraphStyle: 'normal' }).targets).toEqual([native]);
    for (const styleBuiltIn of ['BodyText', 'Body Text 2', 'NoSpacing']) {
        expect(paragraph({ style: '自定义显示名称', styleBuiltIn })).toMatchObject({ role: 'body', eligible: true });
    }
});

test('custom complete prose is included but arbitrary short custom text remains unresolved', () => {
    expect(paragraph({ style: 'Journal Main', styleBuiltIn: 'Other' })).toMatchObject({ role: 'body', eligible: true });
    const text = 'These results are consistent with the comparison.';
    expect(paragraph({ text, ooxml: xml(text), style: '正文段落', styleBuiltIn: 'Other' })).toMatchObject({ role: 'body' });
    expect(paragraph({ text: 'Jane Smith', ooxml: xml('Jane Smith'), style: 'Author Names', styleBuiltIn: 'Other' }))
        .toMatchObject({ role: 'unknown', eligible: false });
});

test.each([
    ['numbered manual heading', '6 1 What the mechanism study establishes', '', 'heading'],
    ['named heading', 'Discussion', '', 'heading'],
    ['figure caption', 'Figure 2. Native document observations', '', 'caption'],
    ['Chinese table caption', '表 3：比较结果', '', 'caption'],
    ['manual outline heading', 'Mechanism evaluation', '<w:outlineLvl w:val="1"/>', 'heading'],
    ['manual keep-next heading', 'Mechanism evaluation', '<w:keepNext/>', 'heading'],
])('Normal %s is excluded by semantic body targeting', (_name, text, properties, role) => {
    const item = paragraph({ text, ooxml: xml(text, properties) });
    expect(item.role).toBe(role);
    expect(resolve([item], { paragraphRole: 'body' }).targets).toEqual([]);
});

test('Normal whole-paragraph manual bold excludes a short heading but not complete bold prose', () => {
    const boldXml = (text) => `<w:p xmlns:w="${W}"><w:r><w:rPr><w:b/></w:rPr><w:t>${text}</w:t></w:r></w:p>`;
    expect(paragraph({ text: 'Mechanism evaluation', ooxml: boldXml('Mechanism evaluation') })).toMatchObject({ role: 'heading' });
    const longProse = `${prose} The final observation adds an additional sentence so this is clearly a complete body paragraph rather than a short heading.`;
    expect(paragraph({ text: longProse, ooxml: boldXml(longProse) })).toMatchObject({ role: 'body' });
    expect(paragraph({ text: 'Mechanism evaluation', ooxml: xml('Mechanism evaluation'), bold: true })).toMatchObject({ role: 'heading' });
});

test('long unpunctuated titles and short bold sentence-like text remain unresolved rather than declared body/headings', () => {
    const title = 'Mechanisms of document agent structure and progression '.repeat(4).trim();
    expect(paragraph({ text: title, ooxml: xml(title) })).toMatchObject({ role: 'unknown', eligible: false });
    const heading = 'Why do mechanisms remain important?';
    expect(paragraph({ text: heading, ooxml: xml(heading), bold: true })).toMatchObject({ role: 'unknown', eligible: false, verified: true });
    const shortProse = paragraph({ bold: true });
    expect(shortProse).toMatchObject({ role: 'unknown', eligible: false, verified: true });
    expect(resolve([shortProse], { paragraphRole: 'body' }).exclusions[0]).toMatchObject({ role: 'unknown' });
    expect(resolve([shortProse], { paragraphIds: ['p1'] }).targets).toEqual([shortProse]);
});

test('Normal table cells and image-bearing paragraphs cannot enter the body write set', () => {
    const table = paragraph({ inTable: true });
    const image = paragraph({ index: 1, ooxml: xml(prose, '', '<w:r><w:drawing/></w:r>') });
    const body = paragraph({ index: 2 });
    expect(table.role).toBe('table');
    expect(image.role).toBe('object');
    expect(resolve([table, image, body], { paragraphStyle: 'normal' }, { bodyOnly: true }).targets).toEqual([body]);
    expect(resolve([table, image, body], {}, { bodyOnly: true }).exclusions.map((item) => item.role)).toEqual(['table', 'object']);
});

test('complete body prose with an inline reference or field remains eligible for scalar paragraph formatting', () => {
    for (const markup of ['<w:r><w:footnoteReference w:id="1"/></w:r>',
        '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>REF example</w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>']) {
        expect(paragraph({ ooxml: xml(prose, '', markup) })).toMatchObject({ role: 'body', eligible: true });
    }
});

test.each([
    ['missing XML', { ooxml: undefined }],
    ['invalid XML', { ooxml: '<invalid' }],
    ['multiple paragraphs', { ooxml: `<w:body xmlns:w="${W}">${xml(prose)}${xml(prose)}</w:body>` }],
    ['unreadable membership', { inTable: undefined }],
    ['partial scope', { withinScope: false }],
    ['different native text', { text: `${prose} Edited` }],
    ['tracked deleted text', { ooxml: xml(prose, '', '<w:del><w:r><w:delText>old</w:delText></w:r></w:del>') }],
    ['unknown native markup', { ooxml: xml(prose, '', '<custom xmlns="urn:custom"/>') }],
])('%s stays excluded rather than defaulting to body', (_name, changes) => {
    const item = paragraph(changes);
    expect(item.eligible).toBe(false);
    expect(resolve([item], { paragraphRole: 'body' }).targets).toEqual([]);
    if (!item.verified || item.role === 'protected') expect(resolve([item], {}).targets).toEqual([]);
});

test('structurally verified ambiguous text remains available for explicit generic formatting', () => {
    const item = paragraph({ text: 'Hello', ooxml: xml('Hello') });
    expect(item).toMatchObject({ verified: true, eligible: false, role: 'unknown' });
    for (const op of [{}, { match: 'Hello' }, { paragraphStyle: 'normal' }, { paragraphIds: ['p1'] }]) {
        expect(resolve([item], op).targets).toEqual([item]);
        expect(resolve([item], op, { bodyOnly: true }).targets).toEqual([]);
    }
    expect(resolve([item], { paragraphRole: 'body' }).targets).toEqual([]);
    const bibliography = paragraph({ styleBuiltIn: 'Bibliography' });
    expect(resolve([bibliography], { paragraphIds: ['p1'] }).targets).toEqual([bibliography]);
    expect(resolve([bibliography], { paragraphRole: 'body' }).targets).toEqual([]);
});

test('native title/caption/heading styles are preserved even if their text reads like prose', () => {
    for (const [styleBuiltIn, role] of [['Title', 'title'], ['Caption', 'caption'], ['Heading1', 'heading']]) {
        const item = paragraph({ styleBuiltIn });
        expect(item.role).toBe(role);
        expect(resolve([item], { paragraphIds: [item.id] }, { bodyOnly: true }).targets).toEqual([]);
        expect(resolve([item], { paragraphStyle: styleBuiltIn }).targets).toEqual([item]);
    }
});

test('exact IDs preserve captured native paragraph identity without widening to similar text', () => {
    const inventory = [paragraph(), paragraph({ index: 1 })];
    expect(resolve(inventory, { paragraphIds: ['p2'] }).targets).toEqual([inventory[1]]);
    for (const paragraphIds of [['p3'], ['p1', 'p1'], [], [' p1'], ['p0']]) {
        expect(() => resolve(inventory, { paragraphIds })).toThrow(/IDs/);
    }
    expect(() => resolve(inventory, { paragraphIds: ['p1'], paragraphRole: 'body' })).toThrow(/conflicting/);
    expect(() => resolve(inventory, { paragraphRole: 'heading' })).toThrow(/role/);
    expect(() => resolve([inventory[0], inventory[0]], {})).toThrow(/duplicate/);
});

test('unscoped, style, match and explicit IDs cannot bypass body-only authorization', () => {
    const body = paragraph();
    const heading = paragraph({ index: 1, styleBuiltIn: 'Heading1' });
    for (const op of [{}, { paragraphStyle: 'heading1' }, { match: prose }, { paragraphIds: ['p2'] }]) {
        const result = resolve([body, heading], op, { bodyOnly: true });
        expect(result.targets).not.toContain(heading);
        expect(result.exclusions.some((item) => item.id === 'p2')).toBe(true);
    }
});

test('built-in matching never accepts a localized/custom name that only looks like the requested built-in', () => {
    const custom = paragraph({ style: 'Normal', styleBuiltIn: 'Other' });
    expect(resolve([custom], { paragraphStyle: 'normal' }).targets).toEqual([]);
    const heading = paragraph({ style: '标题 1', styleBuiltIn: 'Heading1' });
    expect(resolve([heading], { paragraphStyle: 'heading 1' }).targets).toEqual([heading]);
});

test('explicit body alignment is deterministic only when it covers the entire request', () => {
    expect(requestsBodyFormatting('正文修改为两端对齐')).toBe(true);
    expect(requestsBodyFormatting('Set body text to justified')).toBe(true);
    expect(requestsBodyFormatting('Align all headings')).toBe(false);
    for (const instruction of ['标题居中，不修改正文', '正文保持不变，标题居中', '不要将正文修改为两端对齐',
        'Center headings without changing body text', 'Center headings and leave the body text unchanged']) {
        expect(requestsBodyFormatting(instruction)).toBe(false);
    }
    for (const instruction of ['正文修改为两端对齐', '请将全文正文设置为两端对齐', 'Set body text to justified', 'please align the body paragraphs to left']) {
        expect(explicitBodyAlignmentOp(instruction)).toEqual({ paragraphRole: 'body', paragraph: { alignment: /left/.test(instruction) ? 'left' : 'justified' } });
    }
    for (const instruction of ['正文两端对齐，并清理空格', 'Set body text to justified and remove blank lines',
        '不要将正文修改为两端对齐', '正文修改为两端对齐，不修改标题', 'Center all headings', '正文采用分散对齐']) {
        expect(explicitBodyAlignmentOp(instruction)).toBeNull();
    }
});

test('body-only scope inference never narrows a compound request that also formats non-body targets', () => {
    for (const instruction of ['正文两端对齐，标题加粗', '全文优化正文和标题格式',
        'Format body text and center headings', '正文两端对齐，不修改标题并将表格改为三线表']) {
        expect(requestsBodyFormatting(instruction)).toBe(false);
    }
    for (const instruction of ['将正文两端对齐，不修改标题、表格、图片或其他非正文内容',
        'Justify body text; do not modify headings, tables or images']) {
        expect(requestsBodyFormatting(instruction)).toBe(true);
    }
    expect(requestsBodyFormatting('调整非正文格式')).toBe(false);
});
