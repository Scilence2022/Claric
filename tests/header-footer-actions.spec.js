/** @jest-environment jsdom */
import JSZip from 'jszip';
import { applyHeaderFooterProposal, discardHeaderFooterProposal, prepareHeaderFooterProposal,
    readHeaderFooter } from '../src/taskpane/header-footer-actions.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const RELS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const text = (value) => `<w:r><w:t>${value}</w:t></w:r>`;
const paragraph = (inner) => `<w:p>${inner}</w:p>`;
const field = (code, cache = '1') => `<w:fldSimple w:instr="${code}">${text(cache)}</w:fldSimple>`;
const headerRef = (id, variant = 'default') => `<w:headerReference w:type="${variant}" r:id="${id}"/>`;
const footerRef = (id) => `<w:footerReference w:type="default" r:id="${id}"/>`;

async function docx({ sections, parts, even = false, main = 'Body text', protect = false }) {
    const zip = new JSZip();
    const sectionXml = sections.map((refs, index) => {
        const props = `<w:sectPr>${refs}</w:sectPr>`;
        return index === sections.length - 1 ? props : paragraph(`<w:pPr>${props}</w:pPr>`);
    }).join('');
    zip.file('word/document.xml', `<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>${paragraph(text(main))}${sectionXml}</w:body></w:document>`);
    zip.file('word/settings.xml', `<w:settings xmlns:w="${W}">${even ? '<w:evenAndOddHeaders/>' : ''}${protect ? '<w:documentProtection w:enforcement="1"/>' : ''}</w:settings>`);
    let rel = '';
    for (const [id, value] of Object.entries(parts)) {
        const kind = id.startsWith('h') ? 'header' : 'footer';
        const root = kind === 'header' ? 'hdr' : 'ftr';
        zip.file(`word/${id}.xml`, `<w:${root} xmlns:w="${W}">${value}</w:${root}>`);
        rel += `<Relationship Id="${id}" Type="${R}/${kind}" Target="${id}.xml"/>`;
    }
    zip.file('word/_rels/document.xml.rels', `<Relationships xmlns="${RELS}">${rel}</Relationships>`);
    return zip.generateAsync({ type: 'uint8array' });
}

const officeResult = (value) => ({ status: 'succeeded', value });
const collection = (items) => ({ items, load: jest.fn() });
function world(bytes, { selected = 1, selection } = {}) {
    let current = bytes;
    const install = (changed) => { current = changed; };
    const log = jest.fn();
    const sendActionRequest = jest.fn();
    const refs = [1, 2, 3].map((id) => ({ id, compareLocationWith: ({ id: other }) => ({ value: id === other ? 'Equal' : 'Unrelated' }) }));
    const sections = refs.map((ref) => ({
        body: { getRange: () => ref }, getHeader: jest.fn(), getFooter: jest.fn(),
    }));
    const getFileAsync = jest.fn((_type, _options, callback) => {
        const captured = current;
        callback(officeResult({
            size: captured.length, sliceCount: 1,
            getSliceAsync: (index, complete) => complete(officeResult({ index, data: Array.from(captured) })),
            closeAsync: (complete) => complete(officeResult()),
        }));
    });
    global.Office = { AsyncResultStatus: { Succeeded: 'succeeded' }, FileType: { Compressed: 'compressed' },
        context: { platform: 'Mac', document: { getFileAsync }, requirements: { isSetSupported: jest.fn(() => true) } } };
    global.Word = { run: (fn) => fn({
        document: { sections: collection(sections.slice(0, world.count)),
            getSelection: () => ({ sections: collection([sections[selected - 1]]), ...selection }) },
        sync: jest.fn(async () => {}),
    }) };
    return { deps: { appState: { config: { backend: 'test', providers: { test: { model: 'test' } } } }, sendActionRequest, log }, sections, install, getFileAsync };
}
const oneSection = { sections: [headerRef('h1') + footerRef('f1')], parts: {
    h1: paragraph(text('Draft')),
    f1: paragraph(text('Confidential, page ') + field('PAGE')),
} };
const op = (action, extra, target = 's1:footer:default') => ({ target, action, ...extra });
const proposalArgs = (instruction, scope = 'document') => ({ instruction, scope });

beforeEach(() => { world.count = 1; });
afterEach(() => { delete global.Office; delete global.Word; });

test('lists first/even/default variants, native fields and actual links from the package', async () => {
    const fixtures = {
        sections: [headerRef('h1') + footerRef('f1') + headerRef('h3', 'first') + '<w:titlePg/>',
            headerRef('h2'), ''],
        parts: { ...oneSection.parts, h2: paragraph(text('Draft')), h3: paragraph(text('Cover')) },
        even: true,
    };
    world.count = 3;
    const f = world(await docx(fixtures));
    const result = await readHeaderFooter(f.deps, proposalArgs('检查所有页眉和页脚'));
    expect(result).toHaveLength(18);
    expect(result.find(({ id }) => id === 's1:header:default')).toMatchObject({ paragraphs: 'Draft', affectedSections: [1] });
    expect(result.find(({ id }) => id === 's2:header:default')).toMatchObject({ paragraphs: 'Draft', affectedSections: [2, 3] });
    expect(result.find(({ id }) => id === 's2:header:first')).toMatchObject({ active: false, paragraphs: 'Cover' });
    expect(result.find(({ id }) => id === 's3:footer:even')).toMatchObject({ active: true, paragraphs: '' });
    expect(result.find(({ id }) => id === 's3:footer:default')).toMatchObject({ paragraphs: 'Confidential, page {{PAGE}}', fields: ['PAGE'], affectedSections: [1, 2, 3] });
});

test('replaces exact text alongside a page field only after applying the staged operation', async () => {
    const before = await docx(oneSection);
    const after = await docx({ ...oneSection, parts: { ...oneSection.parts, f1: paragraph(text('Private, page ') + field('PAGE', '2')) } });
    const f = world(before);
    const insertText = jest.fn(() => f.install(after));
    const search = jest.fn(() => collection([{ text: 'Confidential', insertText }]));
    f.sections[0].getFooter.mockReturnValue({ search });
    f.deps.sendActionRequest.mockResolvedValue(JSON.stringify({
        operations: [op('replace', { match: 'Confidential', text: 'Private' })], unsupported: [],
    }));
    const proposal = await prepareHeaderFooterProposal(f.deps, proposalArgs('将页脚中的 Confidential 改为 Private'));
    expect(proposal.ops[0].label).toContain('Confidential, page {{PAGE}} → Private, page {{PAGE}}');
    expect(insertText).not.toHaveBeenCalled();
    await expect(applyHeaderFooterProposal(f.deps, proposal)).resolves.toMatchObject({ verified: true, verifiedStories: 1 });
    expect(search).toHaveBeenCalledWith('Confidential', expect.objectContaining({ matchCase: true }));
    expect(insertText).toHaveBeenCalledWith('Private', 'Replace');
    await expect(applyHeaderFooterProposal(f.deps, proposal)).rejects.toThrow(/already attempted/);
});

test('inserts native PAGE and NUMPAGES using only the reviewed footer operation', async () => {
    const before = await docx({ ...oneSection, parts: { ...oneSection.parts, f1: paragraph('') } });
    const after = await docx({ ...oneSection, parts: { ...oneSection.parts,
        f1: paragraph(text('Page ') + field('PAGE') + text(' of ') + field('NUMPAGES', '3')) } });
    const f = world(before);
    const commands = [];
    const range = { insertText: (value, position) => commands.push([value, position]),
        insertField: (position, value) => { commands.push([value, position]); f.install(after); } };
    const footerParagraph = { getRange: () => range, insertText: jest.fn(), load: jest.fn(), alignment: 'Left' };
    const footer = { paragraphs: { getFirst: () => footerParagraph } };
    f.sections[0].getFooter.mockReturnValue(footer);
    f.deps.sendActionRequest.mockResolvedValue(JSON.stringify({ operations: [
        op('set', { text: 'Internal' }, 's1:header:default'),
        op('append', { text: 'Page {{PAGE}} of {{NUMPAGES}}', alignment: 'Right' }),
    ], unsupported: [] }));
    const proposal = await prepareHeaderFooterProposal(f.deps, proposalArgs('设置页眉为 Internal，同时在页脚右侧添加页码和总页数'));
    await expect(applyHeaderFooterProposal(f.deps, { ...proposal, ops: [proposal.ops[1]] })).resolves.toMatchObject({ verified: true });
    expect(commands).toEqual([['NumPages', 'Start'], [' of ', 'Start'], ['Page', 'Start'], ['Page ', 'Start']]);
    expect(footerParagraph.alignment).toBe('Right');
    expect(f.sections[0].getHeader).not.toHaveBeenCalled();
});

test.each(['selection', 'named'])('rejects a %s edit that would change a linked second section', async (scope) => {
    world.count = 2;
    const f = world(await docx({ ...oneSection, sections: [...oneSection.sections, ''] }));
    f.deps.sendActionRequest.mockResolvedValue(JSON.stringify({ operations: [op('clear', {})], unsupported: [] }));
    await expect(prepareHeaderFooterProposal(f.deps, scope === 'selection'
        ? proposalArgs('清空本节页脚', 'selection') : proposalArgs('清空第1节页脚'))).rejects.toThrow(/linked to sections 1, 2/);
    expect(f.sections[0].getFooter).not.toHaveBeenCalled();
});

test('prevents stale or rejected proposals from writing to Word', async () => {
    const before = await docx(oneSection);
    const changed = await docx({ ...oneSection, main: 'New body' });
    const f = world(before);
    f.deps.sendActionRequest.mockResolvedValue(JSON.stringify({ operations: [op('clear', {})], unsupported: [] }));
    const proposal = await prepareHeaderFooterProposal(f.deps, proposalArgs('清空页脚'));
    const rejected = await prepareHeaderFooterProposal(f.deps, proposalArgs('清空页脚'));
    discardHeaderFooterProposal(f.deps, rejected);
    await expect(applyHeaderFooterProposal(f.deps, rejected)).rejects.toThrow(/unavailable/);
    f.install(changed);
    await expect(applyHeaderFooterProposal(f.deps, proposal)).rejects.toThrow(/changed since review/);
    expect(f.sections[0].getFooter).not.toHaveBeenCalled();
});

test('removes only the requested native page field and verifies its remaining neighbors', async () => {
    const before = await docx({ ...oneSection, parts: { ...oneSection.parts,
        f1: paragraph(text('Footer ') + field('PAGE') + field('NUMPAGES')) } });
    const after = await docx({ ...oneSection, parts: { ...oneSection.parts, f1: paragraph(text('Footer ') + field('NUMPAGES')) } });
    const f = world(before);
    const remove = jest.fn(() => f.install(after));
    const other = jest.fn();
    f.sections[0].getFooter.mockReturnValue({ fields: collection([
        { type: 'Page', locked: false, delete: remove }, { type: 'NumPages', locked: false, delete: other },
    ]) });
    f.deps.sendActionRequest.mockResolvedValue(JSON.stringify({ operations: [op('remove_fields', { fields: ['PAGE'] })], unsupported: [] }));
    const proposal = await prepareHeaderFooterProposal(f.deps, proposalArgs('删除页脚中的当前页码，保留总页数'));
    await expect(applyHeaderFooterProposal(f.deps, proposal)).resolves.toMatchObject({ verified: true });
    expect(remove).toHaveBeenCalledTimes(1);
    expect(other).not.toHaveBeenCalled();
});

test.each([
    ['missing variant', '修改首页页眉内容', op('set', { text: 'Title' }), /disabled/],
    ['whole replacement with a field', '修改页脚', op('set', { text: 'Private' }), /contains fields/],
    ['duplicate field', '添加页脚页码', op('append', { text: '{{PAGE}}' }), /already has/],
    ['missing text match', '修改页脚', op('replace', { match: 'missing', text: 'New' }), /match exactly once/],
    ['outside target', '修改页眉', op('clear', {}), /unavailable/],
])('rejects %s before any Word write', async (_name, instruction, operation, error) => {
    const f = world(await docx(oneSection));
    f.deps.sendActionRequest.mockResolvedValue(JSON.stringify({ operations: [operation], unsupported: [] }));
    await expect(prepareHeaderFooterProposal(f.deps, proposalArgs(instruction))).rejects.toThrow(error);
    expect(f.sections[0].getFooter).not.toHaveBeenCalled();
});

test('reports a failed read-back without allowing an automatic retry', async () => {
    const f = world(await docx(oneSection));
    const clear = jest.fn();
    f.sections[0].getFooter.mockReturnValue({ clear });
    f.deps.sendActionRequest.mockResolvedValue(JSON.stringify({ operations: [op('clear', {})], unsupported: [] }));
    const proposal = await prepareHeaderFooterProposal(f.deps, proposalArgs('清空页脚'));
    await expect(applyHeaderFooterProposal(f.deps, proposal)).rejects.toMatchObject({ mutationAttempted: true });
    await expect(applyHeaderFooterProposal(f.deps, proposal)).rejects.toThrow(/already attempted/);
    expect(clear).toHaveBeenCalledTimes(1);
});

test('explains protected documents, missing host support and missing field support', async () => {
    const f = world(await docx({ ...oneSection, protect: true }));
    await expect(prepareHeaderFooterProposal(f.deps, proposalArgs('清空页脚'))).rejects.toThrow(/protected/);
    Office.context.platform = 'OfficeOnline';
    await expect(prepareHeaderFooterProposal(f.deps, proposalArgs('清空页脚'))).rejects.toThrow(/Mac or Windows/);
    Office.context.platform = 'PC';
    Office.context.requirements.isSetSupported.mockImplementation((_set, version) => version === '1.3');
    f.install(await docx(oneSection));
    f.deps.sendActionRequest.mockResolvedValue(JSON.stringify({ operations: [op('remove_fields', { fields: ['PAGE'] })], unsupported: [] }));
    await expect(prepareHeaderFooterProposal(f.deps, proposalArgs('删除页脚页码'))).rejects.toThrow(/WordApi 1.5/);
});

test('limits a live header selection to its native range and exact selected text', async () => {
    const selection = {
        load: jest.fn(), text: 'Draft', parentBody: { type: 'Header', load: jest.fn() },
        compareLocationWith: jest.fn((range) => ({ value: range.id === 'header' ? 'Inside' : 'Unrelated' })),
    };
    const f = world(await docx(oneSection), { selection });
    f.sections[0].getHeader.mockReturnValue({ getRange: () => ({ id: 'header' }) });
    const args = { ...proposalArgs('改为内部使用', 'selection'), selectionOnly: true, selectionText: 'Draft' };
    f.deps.sendActionRequest.mockResolvedValueOnce(JSON.stringify({
        operations: [op('clear', {}, 's1:header:default')], unsupported: [],
    })).mockResolvedValueOnce(JSON.stringify({
        operations: [op('replace', { match: 'Draft', text: '内部使用' }, 's1:header:default')], unsupported: [],
    }));
    await expect(prepareHeaderFooterProposal(f.deps, args)).rejects.toThrow(/beyond the selected/);
    const result = await prepareHeaderFooterProposal(f.deps, args);
    expect(result.ops[0].label).toContain('Draft → 内部使用');
    expect(f.sections[0].getFooter).not.toHaveBeenCalled();
});
