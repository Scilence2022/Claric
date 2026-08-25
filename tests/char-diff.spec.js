/**
 * Char-diff strategy tests (CJK-aware minimal tracked edits).
 *
 * Pure part: hasCjk / computeCharEdits.
 * Application part: applyCharDiffStrategy against a mocked Word object model
 * where ranges are plain string cursors, so we can assert that a one-comma
 * edit produces exactly one insertion and no deletions.
 */

const { hasCjk, computeCharEdits, applyCharDiffStrategy } = require('../src/lib/word-diff/char-diff.js');

describe('hasCjk', () => {
  test('detects Chinese, Japanese kana, Korean hangul', () => {
    expect(hasCjk('但有一个女孩')).toBe(true);
    expect(hasCjk('こんにちは')).toBe(true);
    expect(hasCjk('한국어')).toBe(true);
  });

  test('returns false for Latin text and empty input', () => {
    expect(hasCjk('hello world')).toBe(false);
    expect(hasCjk('')).toBe(false);
    expect(hasCjk(null)).toBe(false);
  });
});

describe('computeCharEdits', () => {
  test('a comma insertion is a single minimal INSERT op', () => {
    const ops = computeCharEdits(
      '但有一个女孩选择不备份自己与母亲的回忆。',
      '但有一个女孩，选择不备份自己与母亲的回忆。'
    );
    expect(ops).toEqual([
      [0, '但有一个女孩'],
      [1, '，'],
      [0, '选择不备份自己与母亲的回忆。'],
    ]);
  });

  test('identical texts produce a single EQUAL op', () => {
    const ops = computeCharEdits('same text', 'same text');
    expect(ops).toEqual([[0, 'same text']]);
  });

  test('handles empty inputs', () => {
    expect(computeCharEdits('', 'abc')).toEqual([[1, 'abc']]);
    expect(computeCharEdits('abc', '')).toEqual([[-1, 'abc']]);
    expect(computeCharEdits('', '')).toEqual([]);
  });
});

/**
 * Builds a minimal Word API mock: every range is a span over a mutable
 * document text carrying absolute [start, end) offsets; every mutation is
 * recorded in `applied`. The strategy locates all spans before editing
 * (phase 1) and executes edits in reverse document order (phase 2), so
 * stored offsets stay valid for the physical mock too.
 *
 * Mirrors the real Word semantics the strategy relies on:
 *  - getRange('End'|'Start') yields a zero-width span;
 *  - expandTo(other) merges two spans;
 *  - search(needle) only matches fully inside the span's [start, end);
 *  - search rejects needles longer than `searchLimit` (Word's 255 cap).
 */
function makeWordWorld(originalText, { searchLimit } = {}) {
  let docText = originalText;
  const applied = [];

  function makeSpan(start, end) {
    return {
      _start: start,
      _end: end,
      get text() { return docText.slice(start, end); },
      load() {},
      search(needle) {
        if (searchLimit && needle.length > searchLimit) {
          throw new Error('SearchStringInvalid');
        }
        const idx = docText.indexOf(needle, start);
        const hit = idx !== -1 && idx + needle.length <= end;
        return { items: hit ? [makeSpan(idx, idx + needle.length)] : [], load() {} };
      },
      delete() {
        applied.push({ type: 'delete', text: docText.slice(start, end) });
        docText = docText.slice(0, start) + docText.slice(end);
      },
      insertText(text, location) {
        if (location === 'Before') {
          docText = docText.slice(0, start) + text + docText.slice(start);
          applied.push({ type: 'insert', text });
        } else if (location === 'After') {
          docText = docText.slice(0, end) + text + docText.slice(end);
          applied.push({ type: 'insert', text });
        } else if (location === 'Replace') {
          applied.push({ type: 'replace', from: docText.slice(start, end), to: text });
          docText = docText.slice(0, start) + text + docText.slice(end);
        } else {
          throw new Error(`unexpected insert location: ${location}`);
        }
      },
      getRange(location) {
        if (location === 'End') return makeSpan(end, end);
        if (location === 'Start') return makeSpan(start, start);
        throw new Error(`unexpected getRange location: ${location}`);
      },
      expandTo(other) {
        return makeSpan(Math.min(start, other._start), Math.max(end, other._end));
      },
    };
  }

  return {
    applied,
    get docText() { return docText; },
    context: { sync: async () => {}, document: {} },
    range: makeSpan(0, originalText.length),
  };
}

describe('applyCharDiffStrategy', () => {
  beforeEach(() => {
    global.Word = {
      ChangeTrackingMode: { trackAll: 'trackAll', off: 'off' },
      InsertLocation: { before: 'Before', after: 'After', replace: 'Replace' },
      RangeLocation: { start: 'Start', end: 'End' },
    };
  });

  afterEach(() => {
    delete global.Word;
  });

  test('comma insertion in a Chinese sentence applies exactly one insertion', async () => {
    const original = '但有一个女孩选择不备份自己与母亲的回忆。';
    const amended = '但有一个女孩，选择不备份自己与母亲的回忆。';
    const world = makeWordWorld(original);

    const result = await applyCharDiffStrategy(world.context, world.range, original, amended, jest.fn());

    expect(result.strategy).toBe('char');
    expect(result.insertions).toBe(1);
    expect(result.deletions).toBe(0);
    // The whole point: only the comma was inserted, not a sentence replacement.
    expect(world.applied).toEqual([{ type: 'insert', text: '，' }]);
    expect(world.docText).toBe(amended);
  });

  test('deletion applies minimal deletes (per hunk, reverse order)', async () => {
    const original = '今天天气真不错啊。';
    const amended = '今天天气不错。';
    const world = makeWordWorld(original);

    const result = await applyCharDiffStrategy(world.context, world.range, original, amended, jest.fn());

    // dmp aligns on 不错: the minimal edit is deleting 真 and 啊 separately
    expect(result.deletions).toBe(2);
    expect(result.insertions).toBe(0);
    expect(world.applied).toEqual([
      { type: 'delete', text: '啊' },
      { type: 'delete', text: '真' },
    ]);
    expect(world.docText).toBe(amended);
  });

  test('replacement merges into a minimal replace op on the changed span', async () => {
    const original = '他站起来。';
    const amended = '他坐下来。';
    const world = makeWordWorld(original);

    const result = await applyCharDiffStrategy(world.context, world.range, original, amended, jest.fn());

    expect(world.docText).toBe(amended);
    expect(result.replacements).toBe(1);
    expect(world.applied).toEqual([{ type: 'replace', from: '站起', to: '坐下' }]);
  });

  test('multi-hunk edit applies each hunk independently', async () => {
    const original = '第一段保持不变。第二段需要修改。第三段也不变。';
    const amended = '第一段保持不变。第二段已经改好。第三段也不变！';
    const world = makeWordWorld(original);

    await applyCharDiffStrategy(world.context, world.range, original, amended, jest.fn());

    expect(world.docText).toBe(amended);
    // Untouched spans were never deleted or retyped
    const touched = world.applied.map((a) => `${a.text || a.from || ''}${a.to || ''}`).join('');
    expect(touched).not.toContain('第一段');
    expect(touched).not.toContain('第三段');
  });

  test('throws when the document text diverges from originalText', async () => {
    const world = makeWordWorld('文档里的实际内容');
    await expect(
      applyCharDiffStrategy(world.context, world.range, '完全不同的原文', '新文本', jest.fn())
    ).rejects.toThrow(/char-diff/);
  });

  test('long paragraphs with runs beyond Word\'s 255-char search limit still apply minimal edits', async () => {
    // 600-char paragraph: every equal run around the edit exceeds the search limit.
    const left = '左'.repeat(300);
    const right = '右'.repeat(299);
    const original = `${left}，${right}。`;
    const amended = `${left}、${right}。`;
    const world = makeWordWorld(original, { searchLimit: 255 });

    const result = await applyCharDiffStrategy(world.context, world.range, original, amended, jest.fn());

    expect(result.replacements).toBe(1);
    expect(world.applied).toEqual([{ type: 'replace', from: '，', to: '、' }]);
    expect(world.docText).toBe(amended);
  });

  test('default: enables tracking for the edits and always restores off', async () => {
    const original = '但有一个女孩选择不备份自己与母亲的回忆。';
    const amended = '但有一个女孩，选择不备份自己与母亲的回忆。';
    const world = makeWordWorld(original);

    await applyCharDiffStrategy(world.context, world.range, original, amended, jest.fn());

    // trackAll was set for the edit phase and restored to off afterwards.
    expect(world.context.document.changeTrackingMode).toBe('off');
  });

  test('trackChanges:false leaves the document tracking mode untouched', async () => {
    const original = '但有一个女孩选择不备份自己与母亲的回忆。';
    const amended = '但有一个女孩，选择不备份自己与母亲的回忆。';
    const world = makeWordWorld(original);

    const result = await applyCharDiffStrategy(world.context, world.range, original, amended, jest.fn(), { trackChanges: false });

    expect(result.insertions).toBe(1);
    expect(world.docText).toBe(amended);
    // The strategy never touched the caller-owned tracking mode.
    expect('changeTrackingMode' in world.context.document).toBe(false);
  });
});
