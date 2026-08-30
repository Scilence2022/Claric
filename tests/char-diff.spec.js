/**
 * Char-diff strategy tests (CJK-aware minimal tracked edits).
 *
 * Pure part: hasCjk / computeCharEdits.
 * Application part: applyCharDiffStrategy against a mocked Word object model
 * where ranges are plain string cursors, so we can assert that a one-comma
 * edit produces exactly one insertion and no deletions.
 */

const { hasCjk, computeCharEdits, applyCharDiffStrategy, _occurrenceIndex } = require('../src/lib/word-diff/char-diff.js');

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
 *  - search(needle) returns ALL greedy, left-to-right, non-overlapping
 *    matches fully inside the span's [start, end), in document order (the
 *    real Find semantics the occurrence-indexed batch locate maps against);
 *  - search rejects needles longer than `searchLimit` (Word's 255 cap).
 */
function makeWordWorld(originalText, { searchLimit } = {}) {
  let docText = originalText;
  const applied = [];
  const syncFn = jest.fn(async () => {});

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
        const items = [];
        let pos = start;
        while (pos + needle.length <= end) {
          const idx = docText.indexOf(needle, pos);
          if (idx === -1 || idx + needle.length > end) break;
          items.push(makeSpan(idx, idx + needle.length));
          pos = idx + needle.length;
        }
        return { items, load() {} };
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
    get syncCount() { return syncFn.mock.calls.length; },
    context: { sync: syncFn, document: {} },
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

  test('a delete of a REPEATED span maps to the correct (Nth) occurrence', async () => {
    // '甲乙' occurs twice; the diff deletes the SECOND one. A naive
    // first-match locate would delete the wrong instance.
    const original = '甲乙甲乙。';
    const amended = '甲乙。';
    const world = makeWordWorld(original);

    const result = await applyCharDiffStrategy(world.context, world.range, original, amended, jest.fn());

    expect(result.deletions).toBe(1);
    expect(world.applied).toEqual([{ type: 'delete', text: '甲乙' }]);
    expect(world.docText).toBe(amended);
  });

  test('a delete span longer than 400 chars locates via first+last pieces (middle gap covered)', async () => {
    const middle = '中'.repeat(500);
    const original = `头${middle}尾`;
    const amended = '头尾';
    const world = makeWordWorld(original, { searchLimit: 255 });

    const result = await applyCharDiffStrategy(world.context, world.range, original, amended, jest.fn());

    expect(result.deletions).toBe(1);
    expect(world.applied).toEqual([{ type: 'delete', text: middle }]);
    expect(world.docText).toBe(amended);
  });

  test('an insertion after a long equal run anchors via the run tail (within the 255 search cap)', async () => {
    const left = '左'.repeat(300);
    const original = `${left}右`;
    const amended = `${left}插右`;
    const world = makeWordWorld(original, { searchLimit: 255 });

    const result = await applyCharDiffStrategy(world.context, world.range, original, amended, jest.fn());

    expect(result.insertions).toBe(1);
    expect(world.applied).toEqual([{ type: 'insert', text: '插' }]);
    expect(world.docText).toBe(amended);
  });

  test('greedy-overlap positions (e.g. "aa" at offset 1 of "aaaa") fall back to the cursor walk', async () => {
    // The delete span 'aa' starts at offset 1, which Word's greedy whole-range
    // scan cannot produce as a match boundary (the match at 0 consumes it).
    // The occurrence check must detect that and the cursor-walk fallback must
    // still land the edit exactly.
    const original = 'aaaa';
    const amended = 'aXa';
    const world = makeWordWorld(original);

    const result = await applyCharDiffStrategy(world.context, world.range, original, amended, jest.fn());

    expect(result.replacements).toBe(1);
    expect(world.docText).toBe(amended);
  });

  test('two identical EQUAL runs anchoring inserts map to their own occurrences', async () => {
    // Both insert anchors share the token '甲乙' (positions 0 and 8); the
    // batch queues ONE search and maps each anchor to its own match.
    const original = '甲乙。中间。甲乙。';
    const amended = '甲乙P。中间。甲乙Q。';
    const world = makeWordWorld(original);

    const result = await applyCharDiffStrategy(world.context, world.range, original, amended, jest.fn());

    expect(result.insertions).toBe(2);
    expect(world.docText).toBe(amended);
    expect(world.applied).toEqual([
      { type: 'insert', text: 'Q' },
      { type: 'insert', text: 'P' },
    ]);
  });

  test('multi-hunk locate costs a single sync regardless of op count (batched path)', async () => {
    // The whole point of the batch: locate phase = 1 sync, then edit + restore.
    // (The per-op cursor walk would spend one sync per EQUAL/DELETE op.)
    const original = '第一句保持原样。第二句需要改动。第三句也不变。第四句再来一处。';
    const amended = '第一句保持原样。第二句已经改好。第三句也不变。第四句同样换了说法。';
    const world = makeWordWorld(original);

    await applyCharDiffStrategy(world.context, world.range, original, amended, jest.fn());

    expect(world.docText).toBe(amended);
    expect(world.syncCount).toBe(3); // batch locate + edit commit + tracking restore
  });
});

describe('_occurrenceIndex', () => {
  test('unique token resolves to occurrence 0', () => {
    expect(_occurrenceIndex('今天天气不错', 0, '今天')).toBe(0);
    expect(_occurrenceIndex('今天天气不错', 5, '错')).toBe(0);
  });

  test('repeated tokens resolve to their greedy match order', () => {
    expect(_occurrenceIndex('甲乙甲乙。', 0, '甲乙')).toBe(0);
    expect(_occurrenceIndex('甲乙甲乙。', 2, '甲乙')).toBe(1);
    expect(_occurrenceIndex('a.b.c.d', 3, '.')).toBe(1);
    expect(_occurrenceIndex('a.b.c.d', 5, '.')).toBe(2);
  });

  test('single-char tokens always resolve (a char cannot overlap itself)', () => {
    expect(_occurrenceIndex('，的，的，', 4, '，')).toBe(2);
  });

  test('a token consumed by an earlier overlapping match is unresolvable', () => {
    // Greedy scan matches 'aa' at 0 (chars 0-1); the span at offset 1 can
    // never be a whole-range match boundary.
    expect(_occurrenceIndex('aaaa', 1, 'aa')).toBeNull();
    expect(_occurrenceIndex('aaaa', 3, 'aa')).toBeNull();
  });

  test('a match ending exactly at the target position does not consume it', () => {
    expect(_occurrenceIndex('aaaa', 2, 'aa')).toBe(1);
  });
});
