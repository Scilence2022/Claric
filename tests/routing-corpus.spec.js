/**
 * Routing corpus: realistic English and Chinese phrasings with the pipeline
 * each must reach. Run every routing-regex change against this table — a
 * keyword tweak that fixes one phrasing tends to misroute another.
 *
 * Rows are grouped by the reason they exist. Add a row whenever a real
 * misroute is reported, together with its nearest correctly routed sibling.
 */

const { routeTurn, TURN_TYPE } = require('../src/taskpane/conversation.js');
const { BUILTIN_SKILLS } = require('../src/taskpane/skills.js');
const { intentText, politeRequestBody } = require('../src/lib/intent-text.js');

const NONE = { hasSelection: false };
const TEXT = { hasSelection: true, hasTextSelection: true };

/** [input, selection facts, expected turn type, expected scope (optional)] */
const CORPUS = {
  'polite requests are instructions': [
    ['Can you fix the typo in paragraph 3', NONE, TURN_TYPE.DOC_EDIT],
    ['Can you fix the typo in paragraph 3?', NONE, TURN_TYPE.DOC_EDIT],
    ['Could you please rewrite the introduction?', NONE, TURN_TYPE.DOC_EDIT],
    ['Could you make this more formal?', TEXT, TURN_TYPE.SELECTION_EDIT],
    ['Can you translate this into English, please?', TEXT, TURN_TYPE.SELECTION_EDIT],
    ['Would you bold the title?', NONE, TURN_TYPE.FORMAT, 'document'],
    ['你能帮我润色这段话吗？', TEXT, TURN_TYPE.SELECTION_EDIT],
    ['麻烦帮我把标题加粗', NONE, TURN_TYPE.FORMAT, 'document'],
    ['Please list the defined terms', NONE, TURN_TYPE.DOC_QA],
  ],
  'polite questions stay questions': [
    ['Can you explain the indemnification clause?', NONE, TURN_TYPE.DOC_QA],
    ['Can you tell me who the parties are?', NONE, TURN_TYPE.DOC_QA],
    ['Can you summarize the document?', NONE, TURN_TYPE.DOC_QA],
    ['Can you read tracked changes?', NONE, TURN_TYPE.DOC_QA],
    ['Can you see the colors in this chart?', NONE, TURN_TYPE.DOC_QA],
    ['Would you like to add a title?', NONE, TURN_TYPE.DOC_QA],
    ['你能改变选择的表格的样式吗？', NONE, TURN_TYPE.DOC_QA],
    ['可以帮我看看这份合同有什么风险吗？', NONE, TURN_TYPE.DOC_QA],
    ['请问如何插入目录？', NONE, TURN_TYPE.DOC_QA],
  ],
  'negated clauses are constraints, not actions': [
    ['不要修改文档，只告诉我字数', NONE, TURN_TYPE.DOC_QA],
    ["Don't edit anything, just list the dates", NONE, TURN_TYPE.DOC_QA],
    ['Do not change the formatting; tell me the word count', NONE, TURN_TYPE.DOC_QA],
    ['润色全文，不要改变格式', NONE, TURN_TYPE.DOC_EDIT],
    ['请润色全文但不要改变格式', NONE, TURN_TYPE.DOC_EDIT],
    ['Polish the whole document but keep the formatting', NONE, TURN_TYPE.DOC_EDIT],
    ['Shorten it without changing the formatting', TEXT, TURN_TYPE.SELECTION_EDIT],
    ["Don't delete the comments, just fix the typos", NONE, TURN_TYPE.DOC_EDIT],
    ['不要删除批注，只修改错别字', NONE, TURN_TYPE.DOC_EDIT],
    ["Don't forget to bold the title", NONE, TURN_TYPE.FORMAT, 'document'],
    ['删除不要的空段落', NONE, TURN_TYPE.CLEANUP],
    ['保持标题居中', NONE, TURN_TYPE.FORMAT, 'document'],
  ],
  'a question followed by an action step keeps both': [
    ['Tell me what changed, then fix the dates', NONE, TURN_TYPE.COMPOUND],
    ['Explain clause 3, then rewrite it in plain English', NONE, TURN_TYPE.COMPOUND],
    ['总结一下全文，然后修改错别字', NONE, TURN_TYPE.COMPOUND],
    ['How do I insert a table and then bold it', NONE, TURN_TYPE.DOC_QA],
    ['如何润色并加标题？', NONE, TURN_TYPE.DOC_QA],
    ['Explain the difference between bold and italic formatting', NONE, TURN_TYPE.DOC_QA],
    ['What are the risks, and then what should I do?', NONE, TURN_TYPE.DOC_QA],
  ],
  'review requests without a change are answered': [
    ['Check the numbering', NONE, TURN_TYPE.DOC_QA],
    ['Check the numbering', TEXT, TURN_TYPE.DOC_QA],
    ['Review the heading formatting', NONE, TURN_TYPE.DOC_QA],
    ['检查所有表格', NONE, TURN_TYPE.DOC_QA],
    ['Review this section', TEXT, TURN_TYPE.DOC_QA],
    ['检查并修改这段话', TEXT, TURN_TYPE.SELECTION_EDIT],
    ['检查格式并统一字体', NONE, TURN_TYPE.FORMAT, 'document'],
    ['Check the numbering and fix it', NONE, TURN_TYPE.COMPOUND],
    ['分析数据并制作图表', NONE, TURN_TYPE.ILLUSTRATION],
    ['Review the data and draw a diagram of the process', NONE, TURN_TYPE.ILLUSTRATION],
  ],
  'resizing an existing image is image management': [
    ['Make the chart bigger', NONE, TURN_TYPE.IMAGE_TOOL],
    ['Make the image smaller', NONE, TURN_TYPE.IMAGE_TOOL],
    ['把图片放大一点', NONE, TURN_TYPE.IMAGE_TOOL],
    ['Generate an image of a cat', NONE, TURN_TYPE.ILLUSTRATION],
    ['Make a chart of quarterly revenue', NONE, TURN_TYPE.ILLUSTRATION],
  ],
  'established routes stay put': [
    ['Please fix the typo in paragraph 3', NONE, TURN_TYPE.DOC_EDIT],
    ['Make this paragraph bold', TEXT, TURN_TYPE.FORMAT, 'selection'],
    ['Translate this into English', TEXT, TURN_TYPE.SELECTION_EDIT],
    ['Summarize this', TEXT, TURN_TYPE.DOC_QA],
    ['Explain the indemnification clause', NONE, TURN_TYPE.DOC_QA],
    ['What color is the heading?', NONE, TURN_TYPE.DOC_QA],
    ['Bold all the defined terms', NONE, TURN_TYPE.FORMAT, 'document'],
    ['把所有标题设为居中', NONE, TURN_TYPE.FORMAT, 'document'],
    ['Insert a 3x3 table', NONE, TURN_TYPE.TABLE],
    ['Can you delete all comments?', NONE, TURN_TYPE.COMMENT_MANAGEMENT],
    ['请删除所有批注', NONE, TURN_TYPE.COMMENT_MANAGEMENT],
    ['增加标题，并深度润色修改', NONE, TURN_TYPE.COMPOUND],
    ['Can you insert a discussion of XXX at an appropriate place in the article?', NONE, TURN_TYPE.DOCUMENT_EDIT],
    ['请在文章的合适位置插入XXX相关的内容和讨论，保留各级标题和格式', NONE, TURN_TYPE.DOCUMENT_EDIT],
    ['删除多余的空段落', NONE, TURN_TYPE.CLEANUP],
    ['让文章更有感染力', NONE, TURN_TYPE.COMPOUND],
  ],
};

describe.each(Object.entries(CORPUS))('routing corpus: %s', (_group, rows) => {
  // Rest parameter: jest treats a 4th declared parameter as a done callback.
  test.each(rows)('%j %j → %s', (input, facts, expectedType, ...[expectedScope]) => {
    const turn = routeTurn(input, { ...facts, skills: BUILTIN_SKILLS });
    expect(turn.type).toBe(expectedType);
    if (expectedScope) expect(turn.scope).toBe(expectedScope);
    // Pipelines always receive the user's complete wording, constraints included.
    expect(turn.instruction ?? turn.question).toBe(input);
  });
});

describe('intentText', () => {
  test.each([
    ['不要修改文档，只告诉我字数', '告诉我字数'],
    ["Don't edit anything, just list the dates", 'list the dates'],
    ['Polish it but do not change the headings', 'Polish it'],
    ['请润色全文但不要改变格式', '请润色全文'],
    ['Rewrite it, keep the formatting', 'Rewrite it'],
    ['保持标题居中', '保持标题居中'],
    ["Don't edit anything", ''],
    ["Don't forget to bold the title", "Don't forget to bold the title"],
    ['删除不要的空段落', '删除不要的空段落'],
    ['分别加粗两个标题', '分别加粗两个标题'],
  ])('%j → %j', (input, expected) => {
    expect(intentText(input)).toBe(expected);
  });
});

describe('politeRequestBody', () => {
  test.each([
    ['Can you fix the typo?', 'fix the typo'],
    ['Could you please rewrite it, please?', 'rewrite it'],
    ['你能帮我润色这段话吗？', '润色这段话'],
    ['麻烦帮我把标题加粗', '把标题加粗'],
    ['Can you tell me who the parties are?', null],
    ['可以帮我看看有什么风险吗？', null],
    ['你能改变选择的表格的样式吗？', null],
    ['Would you mind fixing it?', null],
    ['Fix the typo', null],
  ])('%j → %j', (input, expected) => {
    expect(politeRequestBody(input)).toBe(expected);
  });
});
