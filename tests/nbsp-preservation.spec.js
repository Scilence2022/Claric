/** @jest-environment jsdom */

import { preserveNonBreakingSpaces, requestsNbspChange } from '../src/lib/nbsp-preservation.js';
import { checkChunkAmendment } from '../src/lib/amendment-guard.js';
import { processChunksParallel } from '../src/lib/orchestrator.js';

const NB = ' ';
const reference = `P.${NB}L. Antkowiak, J.${NB}Lietard, et al. Low cost DNA data storage. Nat. Commun. 11, 5345 (2020).`;
const retyped = reference.replaceAll(NB, ' ');

test('a reference retyped with plain spaces comes back unchanged', () => {
  expect(preserveNonBreakingSpaces(reference, retyped)).toBe(reference);
});

test('real edits survive while untouched non-breaking spaces are restored', () => {
  const edited = retyped.replace('Low cost', 'Low-cost').replace(' et al.', ' and colleagues');
  expect(preserveNonBreakingSpaces(reference, edited))
    .toBe(reference.replace('Low cost', 'Low-cost').replace(' et al.', ' and colleagues'));
  // Inserted spaces stay ordinary; kept non-breaking spaces stay as they are.
  expect(preserveNonBreakingSpaces(`10${NB}kb`, `about 10${NB}kb and 20 kb`)).toBe(`about 10${NB}kb and 20 kb`);
  expect(preserveNonBreakingSpaces(`Fig.${NB}3 shows`, 'Figure 3 shows')).toBe(`Figure${NB}3 shows`);
  expect(preserveNonBreakingSpaces('No special spaces.', 'No special spaces here.')).toBe('No special spaces here.');
  expect(preserveNonBreakingSpaces(`A${NB}B`, '')).toBe('');
});

test('surrogate pairs and line breaks keep their offsets', () => {
  const source = `𝑘${NB}=${NB}20\nP.${NB}L. Antkowiak`;
  expect(preserveNonBreakingSpaces(source, '𝑘 = 20\nP. L. Antkowiak et al.')).toBe(`${source} et al.`);
});

test('instructions about non-breaking spaces opt out', () => {
  for (const text of ['Replace non-breaking spaces with normal spaces', 'convert every nbsp', '把不间断空格换成普通空格', 'remove non breaking space']) {
    expect(requestsNbspChange(text)).toBe(true);
  }
  expect(requestsNbspChange('Polish the wording. Remove extra spaces.')).toBe(false);
});

test('the chunk guard restores non-breaking spaces only in kept paragraphs when asked', () => {
  const chunk = { id: 'c0', paragraphs: [{ text: reference }, { text: `Second${NB}paragraph text stays.` }] };
  const reply = `${retyped}\nSecond paragraph text stays.`;
  expect(checkChunkAmendment(chunk, reply, { preserveNbsp: true }).amendment)
    .toBe(`${reference}\nSecond${NB}paragraph text stays.`);
  expect(checkChunkAmendment(chunk, reply).amendment).toBe(reply);
});

describe('polish runs', () => {
  const run = (template, reply) => processChunksParallel([{ id: 'chunk-10', paragraphs: [{ text: reference }, { text: 'Data are available.' }] }], {
    config: {}, log: jest.fn(), formatContextPrefixFn: () => '',
    sendMessagesFn: async () => reply,
    promptManager: { getActiveMode: () => 'amendment', getActivePrompt: () => ({ template }) },
  });

  test('stage no edit for a reference whose only change is the space type', async () => {
    const [result] = await run('Revise {selection} into polished academic English.', `${retyped}\nData are publicly available.`);
    expect(result.amendment).toBe(`${reference}\nData are publicly available.`);
  });

  test('still convert spaces when the user asks for it', async () => {
    const [result] = await run('Replace non-breaking spaces with regular spaces. {selection}', `${retyped}\nData are available.`);
    expect(result.amendment).toBe(`${retyped}\nData are available.`);
  });
});
