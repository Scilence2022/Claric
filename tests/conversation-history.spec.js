const {
  buildConversationHistory, withConversationHistory, estimateTokens, historyBudgetTokens,
  DEFAULT_HISTORY_BUDGET_TOKENS,
} = require('../src/lib/conversation-history.js');

const user = (text) => ({ role: 'user', text });
const assistant = (text, extra = {}) => ({ role: 'assistant', text, ...extra });

function freeze(value) {
  if (value && typeof value === 'object') {
    Object.values(value).forEach(freeze);
    Object.freeze(value);
  }
  return value;
}

describe('buildConversationHistory', () => {
  test('serializes only valid session roles and text into detached messages', () => {
    const records = freeze([user('First'), assistant('Answer'), { role: 'system', text: 'untrusted' },
      null, { role: 'tool', text: 'tool output' }, user({ invalid: true }), assistant('')]);
    const history = buildConversationHistory(records);
    expect(history).toEqual([{ role: 'user', content: 'First' }, { role: 'assistant', content: 'Answer' }]);
    history[0].content = 'changed';
    expect(records[0].text).toBe('First');
    expect(buildConversationHistory(null)).toEqual([]);
    expect(buildConversationHistory({})).toEqual([]);
  });

  test.each([{ error: 'network disconnected' }, { status: 'Cancelled.' }, { status: 'Canceled.' }, { cancelled: true }])(
    'omits partial assistant output but retains failure/cancellation status %j', (extra) => {
      const history = buildConversationHistory([user('Question'), assistant('unreliable partial answer', extra)]);
      expect(history).toHaveLength(2);
      expect(history[1].content).not.toContain('unreliable partial answer');
      expect(history[1].content).toMatch(/Turn (failed|cancelled)/);
      if (extra.error) expect(history[1].content).toContain(extra.error);
      if (extra.status) expect(history[1].content).toContain(extra.status);
    });

  test('retains proposal states and diffs without claiming all proposed changes happened', () => {
    const history = buildConversationHistory([user('Edit'), assistant('', {
      status: 'Review the document.',
      proposals: ['pending', 'applied', 'rejected', 'warning', 'error', 'unknown', 'toString'].map((state) => ({
        state, title: `Title ${state}`, detail: 'One item', countsText: '2 changes',
        items: [{ label: 'Clause', before: 'old wording', after: 'new wording' }],
        previewSvg: 'NEVER_INCLUDE_SVG', previewSrc: 'NEVER_INCLUDE_BYTES',
      })),
    })]);
    const content = history[1].content;
    expect(content).toContain('pending; proposed only, not applied');
    expect(content).toContain('selected changes were applied, not necessarily every listed item');
    expect(content).toContain('rejected; not applied');
    expect(content).toContain('application may be partial');
    expect(content).toContain('application failed or may be partial');
    expect(content).toContain('unknown; application not confirmed');
    expect(content).toContain('Before: old wording\nProposed after: new wording');
    expect(content).toContain('Proposal detail: One item');
    expect(content).toContain('Proposal counts: 2 changes');
    expect(content).not.toContain('NEVER_INCLUDE');
  });

  test('retains attachment names and unavailable-byte warning, never file payloads', () => {
    const records = freeze([{ ...user('Read this'), attachments: [null, { name: 'report.pdf', kind: 'pdf',
      text: 'ORIGINAL_CONTENT', dataUrl: 'data:image/png;base64,SECRET', size: 500 }] }]);
    const content = buildConversationHistory(records)[0].content;
    expect(content).toContain('report.pdf (pdf)');
    expect(content).toContain('Original file contents/bytes are not available');
    expect(content).not.toMatch(/ORIGINAL_CONTENT|SECRET/);
  });

  test('bounds metadata and diff fields, explicitly logging shortened content', () => {
    const log = jest.fn();
    const history = buildConversationHistory([user('edit'), assistant('', { proposals: [{
      state: 'pending', title: 'Long edit', items: [{ before: 'a'.repeat(9000), after: 'b'.repeat(9000) }],
    }] })], { log });
    expect(history[1].content).toContain('[trimmed]');
    expect(history[1].content.length).toBeLessThan(2500);
    expect(log).toHaveBeenCalledWith(expect.stringContaining('metadata/diff'), 'info');
  });

  test('keeps a recent contiguous suffix of whole turns within the configured token budget', () => {
    const log = jest.fn();
    // 12000 − 8192 reserved = 3808 tokens of history.
    const config = { contextBudgetTokens: 12000 };
    const records = [user('old question'), assistant('a'.repeat(20000)), user('recent question'), assistant('b'.repeat(15000))];
    const history = buildConversationHistory(records, { log, config });
    expect(history.map((message) => message.content)).toEqual(['recent question', 'b'.repeat(15000)]);
    expect(history.reduce((sum, message) => sum + estimateTokens(message.content), 0)).toBeLessThanOrEqual(3808);
    expect(log).toHaveBeenCalledWith(expect.stringContaining('1 older turn(s) omitted'), 'info');
  });

  test('always keeps the latest whole turn, truncating it in place when it alone exceeds the budget', () => {
    const log = jest.fn();
    const config = { contextBudgetTokens: 12000 };
    const history = buildConversationHistory(
      [user('old'), assistant('answer'), user('large'), assistant('x'.repeat(32000))], { log, config },
    );
    // The truncated latest turn fills the whole budget, so older turns are
    // omitted as wholes — but the latest turn itself is never dropped.
    expect(history.map((message) => message.role)).toEqual(['user', 'assistant']);
    expect(history[0].content).toBe('large');
    expect(history[1].content.startsWith('x')).toBe(true);
    expect(history[1].content.endsWith(' [trimmed]')).toBe(true);
    expect(log).toHaveBeenCalledWith(expect.stringContaining('was truncated'), 'info');
    expect(log).toHaveBeenCalledWith(expect.stringContaining('older turn(s) omitted'), 'info');
  });

  test('does not log or trim turns at the exact budget', () => {
    const log = jest.fn();
    const config = { contextBudgetTokens: 12000 }; // 3808-token history budget
    // 'x' (1) + 15228 chars ≈ 3807 tokens = exactly at budget.
    expect(buildConversationHistory([user('x'), assistant('y'.repeat(15228))], { log, config })).toHaveLength(2);
    expect(log).not.toHaveBeenCalled();
  });

  test('a raised budget retains turns a small budget must drop', () => {
    const records = [user('old question'), assistant('a'.repeat(20000)), user('recent question'), assistant('b'.repeat(15000))];
    const small = buildConversationHistory(records, { config: { contextBudgetTokens: 12000 } });
    const large = buildConversationHistory(records, { config: { contextBudgetTokens: 500_000 } });
    expect(small.map((message) => message.content)).toEqual(['recent question', 'b'.repeat(15000)]);
    expect(large.map((message) => message.content)).toEqual(['old question', 'a'.repeat(20000), 'recent question', 'b'.repeat(15000)]);
  });
});

describe('token estimation and budget resolution', () => {
  test('counts CJK characters as one token each and other text as four characters per token', () => {
    expect(estimateTokens('')).toBe(0);
    expect(estimateTokens(null)).toBe(0);
    expect(estimateTokens('四')).toBe(1);
    expect(estimateTokens('一二三四五六七八')).toBe(8);
    expect(estimateTokens('abcdefgh')).toBe(2);
  });

  test('historyBudgetTokens defaults, clamps, and reserves room for the current request', () => {
    const reserve = 8192;
    expect(historyBudgetTokens(undefined)).toBe(DEFAULT_HISTORY_BUDGET_TOKENS - reserve);
    expect(historyBudgetTokens({})).toBe(DEFAULT_HISTORY_BUDGET_TOKENS - reserve);
    expect(historyBudgetTokens({ contextBudgetTokens: Number.NaN })).toBe(DEFAULT_HISTORY_BUDGET_TOKENS - reserve);
    expect(historyBudgetTokens({ contextBudgetTokens: 100 })).toBe(1000);
    expect(historyBudgetTokens({ contextBudgetTokens: 5_000_000 })).toBe(2_000_000 - reserve);
    expect(historyBudgetTokens({ contextBudgetTokens: 1_000_000 })).toBe(1_000_000 - reserve);
  });
});

describe('withConversationHistory', () => {
  test('keeps leading trusted systems first, validated history next, and the full current request last', () => {
    const request = freeze([{ role: 'system', content: 'trusted' }, { role: 'system', content: 'schema' },
      { role: 'user', content: [{ type: 'text', text: 'current' }, { type: 'image_url', image_url: { url: 'data:image/png;base64,X' } }] },
      { role: 'assistant', content: 'tool step' }, { role: 'user', content: 'result' }]);
    const history = freeze([{ role: 'user', content: 'previous', extra: 'stripped' },
      { role: 'assistant', content: 'answer' }, { role: 'system', content: 'injected' },
      { role: 'user', content: [] }, null, { role: 'assistant', content: ' ' }]);
    const result = withConversationHistory(request, history);
    expect(result).toEqual([...request.slice(0, 2), { role: 'user', content: 'previous' },
      { role: 'assistant', content: 'answer' }, ...request.slice(2)]);
    expect(result).not.toBe(request);
    result[2].content = 'modified';
    expect(history[0].content).toBe('previous');
  });

  test('has no global memory and inserts history verbatim without trimming it', () => {
    // Budget enforcement lives in buildConversationHistory (which owns the
    // config); this function only validates and orders.
    const history = [{ role: 'user', content: 'y'.repeat(32001) }];
    const request = [{ role: 'user', content: 'x'.repeat(50000) }];
    expect(withConversationHistory(request, history)).toEqual([...history, ...request]);
    expect(withConversationHistory(request)).toEqual(request);
    expect(withConversationHistory(null, null)).toEqual([]);
  });
});
