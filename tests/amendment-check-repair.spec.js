/** @jest-environment jsdom */

import { checkChunkAmendment } from '../src/lib/amendment-guard.js';
import { processChunksParallel } from '../src/lib/orchestrator.js';

const source = ['The first paragraph explains the original method in detail.',
  'The second paragraph explains the original results in detail.'];
const chunk = (overrides = {}) => ({ id: 'c0', paragraphs: source.map((text) => ({ text })), ...overrides });
const options = (sendMessagesFn, overrides = {}) => ({
  config: {}, log: jest.fn(), sendMessagesFn,
  promptManager: { getActiveMode: () => 'amendment', getActivePrompt: () => ({ template: 'Polish the wording. {selection}' }) },
  formatContextPrefixFn: () => '', ...overrides,
});

test('checks allow normal edits and in-place revisions but detect unsupported structure', () => {
  expect(checkChunkAmendment(chunk(), source.join('\n')).valid).toBe(true);
  expect(checkChunkAmendment(chunk(), 'Short.').issues[0]).toMatch(/truncated/);
  const revised = chunk({ paragraphs: source.map((text) => ({ text, hasRevisions: true })) });
  expect(checkChunkAmendment(revised, source.join('\n').replaceAll('original', 'updated')).valid).toBe(true);
  expect(checkChunkAmendment(revised, source[0]).issues.join(' ')).toMatch(/unresolved revisions/);
  const table = chunk({ paragraphs: source.map((text) => ({ text, inTable: true })) });
  expect(checkChunkAmendment(table, source[0]).issues.join(' ')).toMatch(/table content/);
  expect(checkChunkAmendment(table, source.join('\n') + '\nNew paragraph.').issues.join(' ')).toMatch(/inside table/);
});

test('repairs a truncated candidate before returning it for review', async () => {
  const repaired = source.join('\n').replaceAll('original', 'updated');
  const send = jest.fn().mockResolvedValueOnce('Short.').mockResolvedValueOnce(repaired);
  const onProgress = jest.fn();
  const result = await processChunksParallel([chunk()], options(send, { onProgress }));
  expect(send).toHaveBeenCalledTimes(2);
  expect(send.mock.calls[1][1].at(-1).content).toContain('less than 30%');
  expect(send.mock.calls[1][1].at(-1).content).toContain('untrusted document data');
  expect(result[0]).toMatchObject({ status: 'fulfilled', amendment: repaired });
  expect(onProgress.mock.calls.at(-1)[0]).toMatchObject({ completed: 1, failed: 0 });
});

test('a repair that still fails the check is rejected, with a bounded request count', async () => {
  const send = jest.fn().mockResolvedValue('Short.');
  const result = await processChunksParallel([chunk()], options(send));
  expect(send).toHaveBeenCalledTimes(2);
  expect(result[0]).toMatchObject({ status: 'rejected', amendment: null });
  expect(result[0].error).toMatch(/check failed after repair.*truncated/);
});

test('valid candidates do not need an extra model call', async () => {
  const send = jest.fn().mockResolvedValue(source.join('\n'));
  await processChunksParallel([chunk()], options(send));
  expect(send).toHaveBeenCalledTimes(1);
});

test('cancellation after a check prevents the repair request', async () => {
  const controller = new AbortController();
  const send = jest.fn(async () => { controller.abort(); return 'Short.'; });
  const result = await processChunksParallel([chunk()], options(send, { signal: controller.signal }));
  expect(send).toHaveBeenCalledTimes(1);
  expect(result[0].status).toBe('cancelled');
});

test('a transient stream failure after partial output restarts the draft in a separate activity entry', async () => {
  jest.useFakeTimers();
  try {
    const onChunkToken = jest.fn();
    const sendStream = jest.fn()
      .mockImplementationOnce(async (_config, _messages, handlers) => {
        handlers.onContent('Incomplete draft');
        throw Object.assign(new Error('Upstream unavailable'), { retryable: true });
      })
      .mockImplementationOnce(async (_config, _messages, handlers) => {
        handlers.onContent(source.join('\n'));
        return { content: source.join('\n') };
      });
    const pending = processChunksParallel([chunk()], options(jest.fn(), { onChunkToken, sendMessagesStreamFn: sendStream }));
    await jest.runAllTimersAsync();
    const result = await pending;
    expect(sendStream).toHaveBeenCalledTimes(2);
    expect(result[0]).toMatchObject({ status: 'fulfilled', amendment: source.join('\n') });
    expect(onChunkToken.mock.calls.map(([info]) => info.id)).toEqual(['c0', 'c0-retry']);
  } finally {
    jest.useRealTimers();
  }
});
