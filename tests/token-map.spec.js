/**
 * Token-map strategy tests: applyTokenMapStrategy against a mocked Word
 * object model where ranges are plain spans over a mutable document text.
 *
 * Key regression covered here: adjacent deleted tokens must be coalesced into
 * ONE spanning range deletion (a single Word revision + undo entry per
 * contiguous run) instead of one delete() call per token. Per-token deletes
 * inflate Word's session undo/revision bookkeeping and cause post-apply
 * editing lag that only closing the document clears.
 */

const { applyTokenMapStrategy } = require('../src/lib/word-diff/token-map.js');

/**
 * Minimal Word API mock. Every range is a span over `docText` with absolute
 * [start, end) offsets; mutations apply immediately and are recorded in
 * `applied`. Mirrors the semantics the strategy relies on:
 *  - getTextRanges([' ']) -> one coarse span over the whole text;
 *  - search(needle) -> every occurrence span inside the coarse span;
 *  - expandTo(other) merges two spans;
 *  - deletes run in reverse document order, so stored offsets stay valid.
 */
function makeWordWorld(originalText) {
  let docText = originalText;
  const applied = [];
  let searchCount = 0;
  let getTextRangesCount = 0;
  const syncFn = jest.fn(async () => flush());
  // Word queues mutations against a consistent pre-sync coordinate view and
  // applies them atomically at context.sync(). The mock mirrors that: calls
  // are recorded in `applied` (call order) and applied to docText only on
  // sync, rebuilt from original coordinates.
  let pending = [];

  function flush() {
    if (pending.length === 0) return;
    const events = pending.slice().sort((a, b) =>
      (a.pos - b.pos) || (a.type === 'insert' ? -1 : 1)); // insert before delete at a boundary
    let out = '';
    let cur = 0;
    for (const ev of events) {
      out += docText.slice(cur, ev.pos);
      if (ev.type === 'insert') {
        out += ev.text;
        cur = ev.pos;
      } else {
        cur = ev.end;
      }
    }
    docText = out + docText.slice(cur);
    pending = [];
  }

  function makeSpan(start, end) {
    return {
      _start: start,
      _end: end,
      get text() { return docText.slice(start, end); },
      load() {},
      getTextRanges() {
        getTextRangesCount++;
        return { items: [makeSpan(start, end)], load() {} };
      },
      search(needle) {
        searchCount++;
        const items = [];
        let idx = docText.indexOf(needle, start);
        while (idx !== -1 && idx + needle.length <= end) {
          items.push(makeSpan(idx, idx + needle.length));
          idx = docText.indexOf(needle, idx + 1);
        }
        return { items, load() {} };
      },
      expandTo(other) {
        return makeSpan(Math.min(start, other._start), Math.max(end, other._end));
      },
      getRange(location) {
        if (location === 'End') return makeSpan(end, end);
        if (location === 'Start') return makeSpan(start, start);
        throw new Error(`unexpected getRange location: ${location}`);
      },
      delete() {
        applied.push({ type: 'delete', text: docText.slice(start, end) });
        pending.push({ type: 'delete', pos: start, end });
      },
      insertText(text, location) {
        const pos = location === 'Before' ? start
          : location === 'After' ? end
          : location === 'Replace' ? start
          : (() => { throw new Error(`unexpected insert location: ${location}`); })();
        applied.push({ type: 'insert', text });
        pending.push({ type: 'insert', pos, text });
        if (location === 'Replace') pending.push({ type: 'delete', pos: start, end });
      },
    };
  }

  return {
    applied,
    get docText() { return docText; },
    get searchCount() { return searchCount; },
    get getTextRangesCount() { return getTextRangesCount; },
    get syncCount() { return syncFn.mock.calls.length; },
    context: { sync: syncFn, document: {} },
    range: makeSpan(0, originalText.length),
  };
}

describe('applyTokenMapStrategy — delete coalescing', () => {
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

  test('adjacent deleted words collapse into a single spanning delete', async () => {
    const original = 'the quick brown fox jumps';
    const amended = 'the fox jumps';
    const world = makeWordWorld(original);

    const result = await applyTokenMapStrategy(world.context, world.range, original, amended, jest.fn());

    // "quick brown " is 4 consecutive tokens (word, space, word, space) —
    // the whole point: ONE delete op, not four.
    expect(world.applied).toEqual([{ type: 'delete', text: 'quick brown ' }]);
    expect(world.docText).toBe(amended);
    expect(result.deletions).toBe(4); // still counts deleted tokens
  });

  test('separated delete runs stay separate; replacement pairs delete+insert', async () => {
    const original = 'aaa bbb ccc ddd eee';
    const amended = 'aaa ccc eee';
    const world = makeWordWorld(original);

    const result = await applyTokenMapStrategy(world.context, world.range, original, amended, jest.fn());

    // Two non-adjacent runs (" bbb"/" bbb " and " ddd") → two deletes,
    // executed in reverse document order.
    const deletes = world.applied.filter((a) => a.type === 'delete');
    expect(deletes.length).toBe(2);
    expect(deletes[1].text).toContain('bbb');
    expect(deletes[0].text).toContain('ddd');
    expect(world.docText).toBe(amended);
    expect(result.deletions).toBeGreaterThan(0);
  });

  test('mixed replace: kept text is never deleted, minimal ops applied', async () => {
    const original = 'the quick fox jumps over';
    const amended = 'the quick red fox jumps high';
    const world = makeWordWorld(original);

    const result = await applyTokenMapStrategy(world.context, world.range, original, amended, jest.fn());

    expect(world.docText).toBe(amended);
    expect(result.strategy).toBe('token');
    // Untouched words were never deleted or retyped
    const deletedText = world.applied.filter((a) => a.type === 'delete').map((a) => a.text).join('');
    expect(deletedText).not.toContain('quick');
    expect(deletedText).not.toContain('jumps');
  });

  test('locates only the tokens the edits touch (no coarse reads, no full-token searches)', async () => {
    // Six-word paragraph with one deleted word: the batch must search the
    // deleted word and its preceding space only — not consult getTextRanges
    // and not issue one search per document token (upstream behavior).
    const original = 'alpha beta gamma delta epsilon zeta';
    const amended = 'alpha beta delta epsilon zeta';
    const world = makeWordWorld(original);

    const result = await applyTokenMapStrategy(world.context, world.range, original, amended, jest.fn(), { trackChanges: false });

    expect(world.docText).toBe(amended);
    expect(result.deletions).toBe(2); // 'gamma' + its space, coalesced into one delete op
    expect(world.getTextRangesCount).toBe(0);
    expect(world.searchCount).toBe(2); // 'gamma' and ' ', deduped
    // Batched locate (searches) + commit = 2 syncs; tracking owned by caller.
    expect(world.syncCount).toBe(2);
  });

  test('insert anchors map to their own occurrence among repeated words', async () => {
    // 'big ' is inserted before the SECOND 'cat'; its 'the' anchor must map
    // to the second 'the' occurrence, not the first.
    const original = 'the cat sat the cat ran';
    const amended = 'the cat sat the big cat ran';
    const world = makeWordWorld(original);

    const result = await applyTokenMapStrategy(world.context, world.range, original, amended, jest.fn(), { trackChanges: false });

    expect(world.docText).toBe(amended);
    expect(result.insertions).toBe(1);
    expect(world.applied).toEqual([{ type: 'insert', text: 'big ' }]);
  });

  test('leading insertion anchors at the range start (no token located for it)', async () => {
    const original = 'middle text here';
    const amended = 'start middle text here';
    const world = makeWordWorld(original);

    const result = await applyTokenMapStrategy(world.context, world.range, original, amended, jest.fn(), { trackChanges: false });

    expect(world.docText).toBe(amended);
    expect(result.insertions).toBe(1);
    expect(world.applied).toEqual([{ type: 'insert', text: 'start ' }]);
  });
});
