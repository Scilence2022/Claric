/**
 * Unit tests for src/lib/amendment-guard.js
 *
 * Reproduces the offline-pilot failures of the whole-document amendment route:
 * - a chat reply ("The Summary paragraph was not included in the text provided")
 *   accepted as the chunk's amended text
 * - an unrelated paragraph rewritten under "Rewrite only X, change nothing else"
 */
const {
  classifyChunkReply,
  summarizeAmendment,
  parseEditScope,
  locateScopeTargets,
  confineAmendment,
} = require('../src/lib/amendment-guard.js');

const CHUNK = [
  'Recommendations',
  'We recommend adopting the phased rollout described above.',
  'Risk assessment',
  'The overall delivery risk is moderate and manageable.',
  'Budget',
  'The projected budget is within the approved envelope.',
];

describe('classifyChunkReply', () => {
  const REFUSAL = 'The Summary paragraph was not included in the text provided. '
    + 'The section supplied for amendment contains only Recommendations, Risk assessment, and Budget.';

  test('the pilot refusal text is a meta reply, not an amendment', () => {
    const verdict = classifyChunkReply(CHUNK, REFUSAL, { instruction: 'Rewrite only the Summary paragraph.' });
    expect(verdict.metaReply).toBe(true);
    expect(verdict.reason).toMatch(/explanation/);
  });

  test('a multi-line refusal is still a meta reply', () => {
    const verdict = classifyChunkReply(CHUNK, `I'm sorry, I cannot find the Summary section.\nPlease provide the full document.`);
    expect(verdict.metaReply).toBe(true);
  });

  test('a Chinese refusal is a meta reply', () => {
    const verdict = classifyChunkReply(['建议', '建议分阶段推进。'], '抱歉，提供的文本中没有找到摘要部分，无法修改。');
    expect(verdict.metaReply).toBe(true);
  });

  test('an echoed or edited chunk is an amendment', () => {
    expect(classifyChunkReply(CHUNK, CHUNK.join('\n')).metaReply).toBe(false);
    const edited = CHUNK.map((line) => line.replace('moderate', 'low'));
    expect(classifyChunkReply(CHUNK, edited.join('\n')).metaReply).toBe(false);
  });

  test('a translated chunk without chat phrasing is not dropped', () => {
    const translated = ['建议', '我们建议采用上述分阶段推广。', '风险评估', '总体交付风险为中等且可控。', '预算', '预计预算在批准范围内。'];
    expect(classifyChunkReply(CHUNK, translated.join('\n')).metaReply).toBe(false);
  });

  test('a document that legitimately says "the text provided" is not flagged', () => {
    const doc = ['The text provided by the vendor is binding.', 'Payment is due in 30 days.'];
    const edited = ['The text provided by the vendor is binding.', 'Payment is due within 30 days.'];
    expect(classifyChunkReply(doc, edited.join('\n')).metaReply).toBe(false);
  });

  test('text the instruction itself quotes is not treated as chat', () => {
    const verdict = classifyChunkReply(['Old clause'], 'The text provided is withdrawn.', {
      instruction: 'Replace the clause with: The text provided is withdrawn.',
    });
    expect(verdict.metaReply).toBe(false);
  });

  test('an edit that keeps the chunk but adds a note is left to the proposal diff', () => {
    const verdict = classifyChunkReply(CHUNK, `${CHUNK.join('\n')}\nNote: no changes were needed.`);
    expect(verdict.metaReply).toBe(false);
  });

  test('an empty reply is not a meta reply', () => {
    expect(classifyChunkReply(CHUNK, '  \n ').metaReply).toBe(false);
  });
});

describe('summarizeAmendment', () => {
  test('the refusal text deletes every paragraph and inserts one: wholesale', () => {
    const summary = summarizeAmendment(CHUNK, 'The Summary paragraph was not included in the text provided.');
    expect(summary.total).toBe(6);
    expect(summary.deleted).toBe(6);
    expect(summary.inserted).toBe(1);
    expect(summary.wholesale).toBe(true);
  });

  test('a one-paragraph edit is not wholesale', () => {
    const edited = [...CHUNK];
    edited[1] = 'We recommend adopting the phased rollout described earlier.';
    const summary = summarizeAmendment(CHUNK, edited.join('\n'));
    expect(summary.deleted).toBe(0);
    expect(summary.modified).toBe(1);
    expect(summary.unchanged).toBe(5);
    expect(summary.wholesale).toBe(false);
  });

  test('deleting a single paragraph is not wholesale', () => {
    const summary = summarizeAmendment(CHUNK, CHUNK.slice(0, 5).join('\n'));
    expect(summary.deleted).toBe(1);
    expect(summary.wholesale).toBe(false);
  });
});

describe('parseEditScope', () => {
  test.each([
    ['Rewrite only the Abstract paragraph to be more concise. Do not change any other paragraph.', 'Abstract'],
    ['Rewrite the Summary paragraph. Leave all other paragraphs unchanged.', 'Summary'],
    ['Fix the grammar in the Methods section, and do not change any other section', 'Methods'],
    ['Shorten the conclusion', 'conclusion'],
    ['只修改摘要部分，其他段落不要改动', '摘要'],
  ])('%s is restricted to its named target', (instruction, name) => {
    const scope = parseEditScope(instruction);
    expect(scope.restricted).toBe(true);
    expect(scope.targets.map((t) => t.name.toLowerCase())).toContain(name.toLowerCase());
  });

  test.each([
    'Polish the document',
    'Fix all typos in the document',
    'Only fix typos',
    'Translate the document into English',
    'Rewrite the following paragraph',
    'Polish the whole document so it matches the tone of the Introduction section',
    'Polish all sections, especially the Abstract section',
    'Fix the grammar in the Methods section',
    '润色全文，不要改变格式',
  ])('%s is not restricted', (instruction) => {
    expect(parseEditScope(instruction).restricted).toBe(false);
  });

  test('positional references are restricted but need a human-style count', () => {
    const scope = parseEditScope('Can you fix the typo in paragraph 3');
    expect(scope.restricted).toBe(true);
    expect(scope.positional).toBe(true);
  });
});

describe('locateScopeTargets', () => {
  const DOC = [
    { index: 0, text: '1. Abstract', headingLevel: 1 },
    { index: 1, text: 'We study the effect of X on Y.', headingLevel: 0 },
    { index: 2, text: '2. Results', headingLevel: 1 },
    { index: 3, text: 'Y rose by ten percent.', headingLevel: 0 },
    { index: 4, text: 'Discussion', headingLevel: 1 },
    { index: 5, text: 'This suggests a causal link.', headingLevel: 0 },
  ];

  test('a named section resolves to the body under its heading', () => {
    const located = locateScopeTargets(DOC, parseEditScope('Rewrite only the Abstract paragraph'));
    expect([...located.indexes]).toEqual([1]);
  });

  test('a paragraph opening with the name as a label resolves to that paragraph', () => {
    const doc = [
      { index: 0, text: 'Abstract: We study X.', headingLevel: 0 },
      { index: 1, text: 'Body text.', headingLevel: 0 },
    ];
    const located = locateScopeTargets(doc, parseEditScope('Rewrite only the Abstract paragraph'));
    expect([...located.indexes]).toEqual([0]);
  });

  test('a missing or positional target does not resolve', () => {
    expect(locateScopeTargets(DOC, parseEditScope('Rewrite only the Conclusion paragraph'))).toBeNull();
    expect(locateScopeTargets(DOC, parseEditScope('Fix the typo in paragraph 3'))).toBeNull();
  });

  test('an ambiguous target does not resolve', () => {
    const doc = [...DOC, { index: 6, text: 'Abstract', headingLevel: 1 }, { index: 7, text: 'Another.', headingLevel: 0 }];
    expect(locateScopeTargets(doc, parseEditScope('Rewrite only the Abstract paragraph'))).toBeNull();
  });
});

describe('confineAmendment', () => {
  const ORIG = [
    'We study the effect of X on Y.',
    'Y rose by ten percent in the treated group.',
    'This suggests a causal link between X and Y.',
  ];

  test('an unrelated rewritten paragraph is reverted; the target edit is kept', () => {
    const reply = [
      'We investigate how X affects Y.',
      'Y rose by ten percent in the treated group.',
      'This points to a causal relationship between X and Y.',
    ];
    const confined = confineAmendment(ORIG, [true, false, false], reply.join('\n'));
    expect(confined.text.split('\n')).toEqual([
      'We investigate how X affects Y.',
      ORIG[1],
      ORIG[2],
    ]);
    expect(confined.revertedParagraphs).toBe(1);
  });

  test('a reply that only touches out-of-scope paragraphs yields no change', () => {
    const reply = [ORIG[0], ORIG[1], 'This points to a causal relationship between X and Y.'];
    const confined = confineAmendment(ORIG, [true, false, false], reply.join('\n'));
    expect(confined.text).toBeNull();
    expect(confined.revertedParagraphs).toBe(1);
  });

  test('out-of-scope deletions and insertions are reverted', () => {
    const reply = [ORIG[0], ORIG[1], 'A stray new paragraph.'];
    const confined = confineAmendment(ORIG, [true, false, false], reply.join('\n'));
    expect(confined.text).toBeNull();
  });

  test('a paragraph inserted after the target is kept', () => {
    const reply = [ORIG[0], 'A second abstract paragraph.', ORIG[1], ORIG[2]];
    const confined = confineAmendment(ORIG, [true, false, false], reply.join('\n'));
    expect(confined.text.split('\n')).toEqual(reply);
  });
});
