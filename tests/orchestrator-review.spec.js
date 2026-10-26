const { processChunksParallel } = require('../src/lib/orchestrator.js');
const { extractContext } = require('../src/lib/context-extractor.js');
const { BUILTIN_SKILLS } = require('../src/taskpane/skills.js');
const reviewPolicy = BUILTIN_SKILLS.find((s) => s.name === 'check-doc').reviewPolicy;
const paragraphs = [{ index: 0, text: 'collabortive MFM results.' }, { index: 1, text: 'Maternal and Fetal Medicine (MFM).' }];
const chunks = [{ id: 'chunk-0', paragraphs: [paragraphs[0]], startIndex: 0, endIndex: 0 }];
const finding = { type: 'spelling', paragraphId: 'p-1', quote: 'collabortive', explanation: 'Misspelling.',
  suggestion: 'collaborative', subject: 'MFM' };
const draft = JSON.stringify({ findings: [finding] });
const decision = JSON.stringify({ decisions: [{ id: 'f1', verdict: 'confirmed', reason: 'The exact source contains a misspelling.' }] });
function options(sendMessagesFn, extra = {}) {
  return { config: {}, promptManager: { getActiveMode: () => 'comment', getActivePrompt: () => null },
    documentContext: extractContext({ paragraphs }), log: jest.fn(), reviewPolicy, sendMessagesFn, ...extra };
}

test('built-in check-doc verifies candidates with global source before returning any comment', async () => {
  const send = jest.fn().mockResolvedValueOnce(draft).mockResolvedValueOnce(decision);
  const [result] = await processChunksParallel(chunks, options(send, { conversationHistory: [{ role: 'user', content: 'OLD HISTORY' }] }));
  expect(result).toMatchObject({ status: 'fulfilled', amendment: null });
  expect(result.comment).toContain('Suggested correction: collaborative');
  expect(result.comment).not.toContain('"findings"');
  expect(send).toHaveBeenCalledTimes(2);
  const initial = send.mock.calls[0][1];
  expect(initial.at(-2).content).toContain('Return ONLY JSON');
  const verification = send.mock.calls[1][1];
  expect(verification).toHaveLength(2);
  expect(JSON.stringify(verification)).toContain('Maternal and Fetal Medicine');
  expect(JSON.stringify(verification)).not.toContain('OLD HISTORY');
});

test('normal custom comments retain their existing single-call behavior', async () => {
  const send = jest.fn().mockResolvedValue('Ordinary comment');
  const [result] = await processChunksParallel(chunks, options(send, { reviewPolicy: null }));
  expect(result.comment).toBe('Ordinary comment');
  expect(send).toHaveBeenCalledTimes(1);
});

test('an existing global reference is withheld without requesting model confirmation', async () => {
  const source = [{ index: 0, text: 'See guidance 1.' }, { index: 1, text: 'References' },
    { index: 2, text: 'Guidance for subspecialty training on-call commitments.', listString: '1.' }];
  const send = jest.fn().mockResolvedValue(JSON.stringify({ findings: [{ ...finding, type: 'broken_reference',
    quote: '1', target: { kind: 'bibliography', label: '1' } }] }));
  const [result] = await processChunksParallel([{ ...chunks[0], paragraphs: [source[0]] }],
    options(send, { documentContext: extractContext({ paragraphs: source }) }));
  expect(result.comment).toBeNull();
  expect(result.review).toEqual({ confirmedCount: 0, withheldCount: 1 });
  expect(send).toHaveBeenCalledTimes(1);
});

test('malformed drafting gets one repair and all verification remains mandatory', async () => {
  const send = jest.fn().mockResolvedValueOnce('Unsupported prose').mockResolvedValueOnce(draft).mockResolvedValueOnce(decision);
  const [result] = await processChunksParallel(chunks, options(send));
  expect(result.status).toBe('fulfilled');
  expect(send).toHaveBeenCalledTimes(3);
  expect(send.mock.calls[1][1].at(-1).content).toContain('previous response was invalid');
});

test.each([
  ['draft', ['Unsupported prose', 'Still unsupported']],
  ['verification', [draft, '{"decisions":[]}']],
])('invalid %s fails the chunk and publishes no findings', async (_name, responses) => {
  const send = jest.fn();
  for (const response of responses) send.mockResolvedValueOnce(response);
  const [result] = await processChunksParallel(chunks, options(send));
  expect(result.status).toBe('rejected');
  expect(result.comment).toBeNull();
  expect(send).toHaveBeenCalledTimes(2);
});

test('verification cancellation cannot return a comment even if the transport resolves', async () => {
  const controller = new AbortController();
  const send = jest.fn().mockResolvedValueOnce(draft).mockImplementationOnce(async () => { controller.abort(); return decision; });
  const [result] = await processChunksParallel(chunks, options(send, { signal: controller.signal }));
  expect(result).toMatchObject({ status: 'cancelled', comment: null });
});

test('verification transport failure preserves a retryable failure', async () => {
  const send = jest.fn().mockResolvedValueOnce(draft).mockRejectedValueOnce(new Error('Verification unavailable'));
  const [result] = await processChunksParallel(chunks, options(send));
  expect(result).toMatchObject({ status: 'rejected', error: 'Verification unavailable', comment: null });
});

test('streams verification under a separate model activity ID', async () => {
  const onChunkToken = jest.fn();
  const stream = jest.fn().mockImplementation(async (_config, _messages, handlers) => {
    const response = stream.mock.calls.length === 1 ? draft : decision;
    handlers.onReasoning('Checking source'); handlers.onContent(response);
    return { content: response };
  });
  const [result] = await processChunksParallel(chunks, options(jest.fn(), { onChunkToken, sendMessagesStreamFn: stream }));
  expect(result.status).toBe('fulfilled');
  expect(onChunkToken.mock.calls.map(([info]) => info.id)).toEqual(['chunk-0', 'chunk-0', 'chunk-0-verify', 'chunk-0-verify']);
});

test('an unsupported policy or incomplete global index cannot bypass the checks', async () => {
  const send = jest.fn();
  await expect(processChunksParallel(chunks, options(send, { reviewPolicy: 'unknown' }))).rejects.toThrow('Unsupported');
  await expect(processChunksParallel(chunks, options(send, { documentContext: {} }))).rejects.toThrow('complete source');
  expect(send).not.toHaveBeenCalled();
});
