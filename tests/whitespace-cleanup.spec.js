const { requestsWhitespaceOnlyCleanup, hasRedundantSpaceCandidates, validateWhitespaceCleanup } = require('../src/lib/whitespace-cleanup.js');

test.each([
  '清理全文非表格正文中多余的空格；仅去除冗余空格，保留有意义的空格、词语和段落结构，不做内容改写。',
  'Remove redundant spaces throughout the body; do not rewrite content.',
  '清理多余空格，保持原有格式，不修改字体。',
  'Remove extra spaces without changing formatting; keep existing fonts.',
  'Remove extra spaces and preserve original indentation.',
])('recognizes narrow space cleanup: %s', (instruction) => {
  expect(requestsWhitespaceOnlyCleanup(instruction)).toBe(true);
});

test.each(['全文优化格式，例如不正确的字体加粗，多余的空格，表格修改为三线格',
  'Remove extra spaces and rewrite the conclusion', '请润色全文', 'Remove extra spaces and blank paragraphs',
  '清理多余空格并调整格式', 'Remove extra spaces and format headings',
  'Remove extra spaces and italicize titles', '清理多余空格，不要改写正文，并加粗标题'])('keeps mixed and general requests out of space-only mode: %s', (instruction) => {
  expect(requestsWhitespaceOnlyCleanup(instruction)).toBe(false);
});

test.each([' leading', 'trailing ', 'two  words', 'Hello , world', '中文 空格'])('detects a source candidate: %s', (text) => {
  expect(hasRedundantSpaceCandidates(text)).toBe(true);
});

test('meaningful word, mathematical and nonbreaking spacing is not an obvious redundant-space candidate', () => {
  expect(hasRedundantSpaceCandidates('Two words.\nx = y + z\nA\u00a0B\n\tIndented text')).toBe(false);
});

test.each([
  ['  Example  body  ', 'Example body'],
  ['First  paragraph\nSecond  paragraph', 'First paragraph\nSecond paragraph'],
  ['Hello , world', 'Hello, world'],
  ['中文 空格', '中文空格'],
  ['A\u00a0B  C\tD', 'A\u00a0B C\tD'],
])('permits conservative space deletion without rewriting: %p', (original, amendment) => {
  expect(validateWhitespaceCleanup(original, amendment)).toEqual({ valid: true });
});

test.each([
  ['Example  body', 'Different body'],
  ['Two words', 'Twowords'],
  ['x = y', 'x=y'],
  ['One\nTwo', 'One Two'],
  ['A\u00a0B', 'A B'],
  ['A\tB', 'A B'],
  ['AB', 'A B'],
])('refuses unsafe cleanup: %p', (original, amendment) => {
  expect(validateWhitespaceCleanup(original, amendment)).toMatchObject({ valid: false });
});
