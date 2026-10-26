/** @jest-environment jsdom */

jest.mock('../src/lib/word-diff/index.js', () => ({
  applyTokenMapStrategy: jest.fn(),
  applySentenceDiffStrategy: jest.fn(),
}));
jest.mock('../src/lib/word-diff/revision-diff.js', () => ({ tryRevisionDiff: jest.fn() }));

const { applyTokenMapStrategy, applySentenceDiffStrategy } = require('../src/lib/word-diff/index.js');
const { applyChunkResults } = require('../src/lib/reassembler.js');
const { RevisionSafetyError, MutationSafetyError } = require('../src/lib/word-revisions.js');
const { tryRevisionDiff } = require('../src/lib/word-diff/revision-diff.js');

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const escaped = (text) => text.replace(/&/g, '&amp;').replace(/</g, '&lt;');
const paragraphXml = (text) => `<w:document xmlns:w="${W}"><w:body><w:p><w:r><w:t>${escaped(text)}</w:t></w:r></w:p></w:body></w:document>`;

function nativeDocument(texts, contentOverrides = {}) {
  const states = texts.map((text) => ({ text }));
  const contentRanges = states.map((state, index) => ({
    get text() { return state.text; },
    load: jest.fn(),
    getOoxml: jest.fn(() => ({ value: paragraphXml(contentOverrides[index] ?? state.text) })),
    insertText: jest.fn(),
    state,
  }));
  const paragraphs = states.map((state, index) => ({
    get text() { return state.text; },
    load: jest.fn(),
    getOoxml: jest.fn(() => ({ value: paragraphXml(state.text) })),
    parentTableOrNullObject: { isNullObject: true, load: jest.fn() },
    getRange: jest.fn(() => contentRanges[index]),
    delete: jest.fn(),
    insertParagraph: jest.fn(),
  }));
  const range = {
    text: texts.join('\r'), isNullObject: false, load: jest.fn(),
    paragraphs: { items: paragraphs, load: jest.fn() },
    insertText: jest.fn(),
  };
  const context = { sync: jest.fn(async () => {}), document: {
    changeTrackingMode: 'Off',
    getBookmarkRangeOrNullObject: jest.fn(() => range),
  } };
  global.Word = {
    run: async (callback) => callback(context),
    ChangeTrackingMode: { trackAll: 'TrackAll', off: 'Off' },
    InsertLocation: { replace: 'Replace', before: 'Before', after: 'After' },
  };
  return { context, states, paragraphs, contentRanges, range };
}

const storedTexts = ['The first paragraph has clear wording.', 'The second paragraph has clear wording.'];
const amendedTexts = ['The first paragraph has precise wording.', 'The second paragraph has precise wording.'];

async function apply(log = jest.fn(), onChunkApplied = jest.fn()) {
  const chunk = { id: 'section', startIndex: 0, endIndex: 1,
    paragraphs: storedTexts.map((text) => ({ text })) };
  const result = await applyChunkResults([
    { chunkId: 'section', status: 'fulfilled', chunk, amendment: amendedTexts.join('\n') },
  ], new Map([['section', '_section']]), {
    trackChangesEnabled: true, lineDiffEnabled: false, log, onChunkApplied,
  });
  return { result, log, onChunkApplied };
}

beforeEach(() => {
  jest.resetAllMocks();
  applyTokenMapStrategy.mockImplementation(async (_context, range, _before, after) => { range.state.text = after; });
});
afterEach(() => { delete global.Word; });

test('an unverified native write stops before touching the remaining section or inserting comments', async () => {
  const document = nativeDocument(storedTexts);
  document.range.insertComment = jest.fn();
  const nativeError = Object.assign(new Error('GeneralException'), { code: 'GeneralException', debugInfo: { errorLocation: 'Range.insertText' } });
  applyTokenMapStrategy.mockRejectedValueOnce(new MutationSafetyError(nativeError));
  const onChunkApplied = jest.fn();
  const results = ['head', 'tail'].map((id, index) => ({ chunkId: id, status: 'fulfilled',
    amendment: amendedTexts.join('\n'), comment: 'Review note.', chunk: { id, endIndex: index,
      paragraphs: storedTexts.map((text) => ({ text })) } }));
  const result = await applyChunkResults(results, new Map([['head', '_head'], ['tail', '_tail']]),
    { trackChangesEnabled: true, log: jest.fn(), onChunkApplied });
  expect(result).toMatchObject({ halted: true, interrupted: false, attemptedChunkIds: ['tail'],
    uncertainChunkIds: ['tail'], skippedChunkIds: ['head'], commentsInserted: 0 });
  expect(applyTokenMapStrategy).toHaveBeenCalledTimes(1);
  expect(document.context.document.getBookmarkRangeOrNullObject).toHaveBeenCalledTimes(1);
  expect(document.range.insertComment).not.toHaveBeenCalled();
  expect(onChunkApplied).toHaveBeenCalledWith('head', { applied: false, skipped: true });
  expect(result.errors[0]).toContain('Range.insertText');
  expect(result.errors[0]).toContain('paragraph 2: native text edit');
});

test('apply-time re-anchoring and changed-content reads bound full OOXML exports to 24 per sync', async () => {
  const source = Array.from({ length: 50 }, (_, index) => `The original paragraph number ${index} has clear wording.`);
  const amended = source.map((text) => text.replace('clear', 'precise'));
  const document = nativeDocument(source);
  let pending = 0;
  let largest = 0;
  for (const range of [...document.paragraphs, ...document.contentRanges]) {
    const getXml = range.getOoxml.getMockImplementation();
    range.getOoxml.mockImplementation(() => { pending++; return getXml(); });
  }
  document.context.sync.mockImplementation(async () => { largest = Math.max(largest, pending); pending = 0; });
  const result = await applyChunkResults([{ chunkId: 'section', status: 'fulfilled', amendment: amended.join('\n'),
    chunk: { id: 'section', paragraphs: source.map((text) => ({ text })) } }], new Map([['section', '_section']]),
  { trackChangesEnabled: true, log: jest.fn() });
  expect(result.amendmentsApplied).toBe(1);
  expect(largest).toBeLessThanOrEqual(24);
  expect(largest).toBe(24);
});

test('model blank separator lines never cause structural paragraph writes', async () => {
  const document = nativeDocument(storedTexts);
  const result = await applyChunkResults([{ chunkId: 'section', status: 'fulfilled', amendment: amendedTexts.join('\n\n\n'),
    chunk: { id: 'section', paragraphs: storedTexts.map((text) => ({ text })) } }], new Map([['section', '_section']]),
  { trackChangesEnabled: true, log: jest.fn() });
  expect(result.amendmentsApplied).toBe(1);
  for (const paragraph of document.paragraphs) {
    expect(paragraph.insertParagraph).not.toHaveBeenCalled();
    expect(paragraph.delete).not.toHaveBeenCalled();
  }
});

test('native content mismatch in the first paragraph prevents writes to the tail of the section', async () => {
  const document = nativeDocument(storedTexts, { 0: 'A different native content baseline.' });
  const { result, onChunkApplied } = await apply();

  expect(result).toMatchObject({ amendmentsApplied: 0, appliedParagraphs: 0,
    appliedChunkIds: [], attemptedChunkIds: ['section'], failedChunkIds: ['section'],
    partialChunkIds: [], uncertainChunkIds: [] });
  // Content is now read during re-anchoring, so drift is caught before
  // reaching the paragraph preflight. No tail paragraph may be written.
  expect(result.errors[0]).toMatch(/original content no longer matches the staged range/);
  expect(applyTokenMapStrategy).not.toHaveBeenCalled();
  expect(document.states.map((state) => state.text)).toEqual(storedTexts);
  expect(document.contentRanges.every((range) => !range.insertText.mock.calls.length)).toBe(true);
  expect(document.range.insertText).not.toHaveBeenCalled();
  expect(onChunkApplied).toHaveBeenCalledWith('section', { applied: false, error: true, partial: false, uncertain: false });
});

test('unlocatable existing revisions fail preflight before any paragraph is written', async () => {
  const document = nativeDocument(storedTexts);
  const revisedXml = () => ({ value: paragraphXml(storedTexts[0]).replace('<w:r>', '<w:ins><w:r>').replace('</w:r>', '</w:r></w:ins>') });
  document.paragraphs[0].getOoxml.mockImplementation(revisedXml);
  document.contentRanges[0].getOoxml.mockImplementation(revisedXml);
  tryRevisionDiff.mockRejectedValueOnce(new RevisionSafetyError('Unverified visible revision target.'));
  const { result } = await apply();
  expect(tryRevisionDiff).toHaveBeenCalledWith(document.context, document.contentRanges[0], storedTexts[0], amendedTexts[0],
    expect.any(Function), { trackChanges: false, paragraph: true, preflightOnly: true, nativeTargets: true });
  expect(result).toMatchObject({ amendmentsApplied: 0, appliedParagraphs: 0, partialChunkIds: [], uncertainChunkIds: [] });
  expect(result.errors[0]).toMatch(/Unverified visible revision target/);
  expect(applyTokenMapStrategy).not.toHaveBeenCalled();
  expect(document.states.map((state) => state.text)).toEqual(storedTexts);
});

test('a later baseline failure reports already verified paragraph edits without retrying the whole range', async () => {
  const document = nativeDocument(storedTexts);
  applyTokenMapStrategy.mockImplementationOnce(async (_context, range, _before, after) => { range.state.text = after; })
    .mockRejectedValueOnce(new RevisionSafetyError('The remaining paragraph baseline changed.'));
  const { result, log, onChunkApplied } = await apply();

  expect(result).toMatchObject({ amendmentsApplied: 0, appliedParagraphs: 1,
    appliedChunkIds: [], attemptedChunkIds: ['section'], failedChunkIds: ['section'],
    partialChunkIds: ['section'], uncertainChunkIds: [] });
  expect(document.states.map((state) => state.text)).toEqual([storedTexts[0], amendedTexts[1]]);
  expect(applyTokenMapStrategy).toHaveBeenCalledTimes(2);
  expect(applySentenceDiffStrategy).not.toHaveBeenCalled();
  expect(document.range.insertText).not.toHaveBeenCalled();
  for (const range of document.contentRanges) expect(range.insertText).not.toHaveBeenCalled();
  expect(log).toHaveBeenCalledWith(expect.stringContaining('1 paragraph edit(s) already applied'), 'error');
  expect(onChunkApplied).toHaveBeenCalledWith('section', { applied: false, error: true, partial: true, uncertain: false });
  expect(document.context.document.changeTrackingMode).toBe('Off');
});

test('an unverified strategy write is reported separately from confirmed paragraph edits and never falls back', async () => {
  const document = nativeDocument(storedTexts);
  applyTokenMapStrategy.mockRejectedValueOnce(new MutationSafetyError(new Error('Host write timed out.')));
  const { result, log, onChunkApplied } = await apply();

  expect(result).toMatchObject({ amendmentsApplied: 0, appliedParagraphs: 0,
    appliedChunkIds: [], attemptedChunkIds: ['section'], failedChunkIds: ['section'],
    partialChunkIds: [], uncertainChunkIds: ['section'] });
  expect(applyTokenMapStrategy).toHaveBeenCalledTimes(1);
  expect(document.range.insertText).not.toHaveBeenCalled();
  for (const range of document.contentRanges) expect(range.insertText).not.toHaveBeenCalled();
  expect(log).toHaveBeenCalledWith(expect.stringContaining('writes could not be verified'), 'error');
  expect(onChunkApplied).toHaveBeenCalledWith('section', { applied: false, error: true, partial: false, uncertain: true });
  expect(document.context.document.changeTrackingMode).toBe('Off');
});

function failTrackingRestoration(context, previousMode = 'Off') {
  let mode = previousMode;
  let enteredTracking = false;
  let restoring = false;
  Object.defineProperty(context.document, 'changeTrackingMode', {
    configurable: true,
    get: () => mode,
    set: (value) => {
      if (value === 'TrackAll') enteredTracking = true;
      restoring = enteredTracking && value === previousMode;
      mode = value;
    },
  });
  context.sync.mockImplementation(async () => {
    if (restoring) throw new Error('Tracking restoration failed after committed edits.');
  });
}

function fallbackDocument(previousMode = 'Off') {
  const state = { text: 'Original fallback paragraph.' };
  const range = { state, text: state.text, isNullObject: false, load: jest.fn(),
    getOoxml: jest.fn(() => ({ value: paragraphXml(state.text) })), insertText: jest.fn() };
  const context = { sync: jest.fn(async () => {}), document: {
    changeTrackingMode: previousMode, getBookmarkRangeOrNullObject: jest.fn(() => range),
  } };
  global.Word = { run: async (callback) => callback(context),
    ChangeTrackingMode: { trackAll: 'TrackAll', off: 'Off' }, InsertLocation: { replace: 'Replace' } };
  return { context, range, state };
}

test('structural edits committed before tracking restoration fails remain reported as uncertain writes', async () => {
  const document = nativeDocument(storedTexts);
  failTrackingRestoration(document.context);
  const onChunkApplied = jest.fn();
  const result = await applyChunkResults([{ chunkId: 'section', status: 'fulfilled', amendment: storedTexts[0],
    chunk: { id: 'section', paragraphs: storedTexts.map((text) => ({ text })) } }],
  new Map([['section', '_section']]), { trackChangesEnabled: true, lineDiffEnabled: false, log: jest.fn(), onChunkApplied });

  expect(document.paragraphs[1].delete).toHaveBeenCalledTimes(1);
  expect(applyTokenMapStrategy).not.toHaveBeenCalled();
  expect(result).toMatchObject({ amendmentsApplied: 0, appliedParagraphs: 0, failedChunkIds: ['section'],
    appliedChunkIds: [], partialChunkIds: [], uncertainChunkIds: ['section'] });
  expect(result.errors[0]).toContain('Tracking restoration failed');
  expect(onChunkApplied).toHaveBeenCalledWith('section', { applied: false, error: true, partial: false, uncertain: true });
});

test('a successful fallback write followed by tracking restoration failure is never reported as untouched', async () => {
  const document = fallbackDocument();
  failTrackingRestoration(document.context);
  const result = await applyChunkResults([{ chunkId: 'section', status: 'fulfilled',
    amendment: 'Revised fallback paragraph.', chunk: { id: 'section' } }],
  new Map([['section', '_section']]), { trackChangesEnabled: true, lineDiffEnabled: false, log: jest.fn() });

  expect(applyTokenMapStrategy).toHaveBeenCalledWith(document.context, document.range,
    'Original fallback paragraph.', 'Revised fallback paragraph.', expect.any(Function), { trackChanges: false });
  expect(document.state.text).toBe('Revised fallback paragraph.');
  expect(result).toMatchObject({ amendmentsApplied: 0, appliedParagraphs: 0, appliedChunkIds: [],
    attemptedChunkIds: ['section'], failedChunkIds: ['section'], partialChunkIds: [], uncertainChunkIds: ['section'] });
});

test('fallback revision-read failure while Stop is pressed restores the original tracking mode before returning', async () => {
  const document = fallbackDocument('TrackMine');
  const controller = new AbortController();
  document.range.getOoxml.mockImplementation(() => {
    controller.abort();
    return { value: '<unreadable />' };
  });
  const result = await applyChunkResults([{ chunkId: 'section', status: 'fulfilled',
    amendment: 'Revised fallback paragraph.', chunk: { id: 'section' } }],
  new Map([['section', '_section']]), {
    trackChangesEnabled: true, lineDiffEnabled: false, log: jest.fn(), signal: controller.signal,
  });

  expect(controller.signal.aborted).toBe(true);
  expect(applyTokenMapStrategy).not.toHaveBeenCalled();
  expect(document.context.document.changeTrackingMode).toBe('TrackMine');
  expect(result).toMatchObject({ amendmentsApplied: 0, appliedChunkIds: [], failedChunkIds: ['section'],
    partialChunkIds: [], uncertainChunkIds: [] });
});

test('paragraph edits restore an existing caller tracking mode even when app tracking is disabled', async () => {
  const document = nativeDocument(storedTexts);
  document.context.document.changeTrackingMode = 'TrackMine';
  const result = await applyChunkResults([{ chunkId: 'section', status: 'fulfilled', amendment: amendedTexts.join('\n'),
    chunk: { id: 'section', paragraphs: storedTexts.map((text) => ({ text })) } }],
  new Map([['section', '_section']]), { trackChangesEnabled: false, lineDiffEnabled: false, log: jest.fn() });

  expect(result.amendmentsApplied).toBe(1);
  expect(document.context.document.changeTrackingMode).toBe('TrackMine');
});

test.each(['paragraph', 'fallback'])('%s application loads the native tracking property before reading it', async (kind) => {
  const document = kind === 'paragraph' ? nativeDocument(storedTexts) : fallbackDocument('TrackMine');
  let loaded = false;
  let queued = false;
  let mode = 'TrackMine';
  document.context.document.load = jest.fn((property) => { if (property === 'changeTrackingMode') queued = true; });
  document.context.sync.mockImplementation(async () => { if (queued) loaded = true; });
  Object.defineProperty(document.context.document, 'changeTrackingMode', {
    configurable: true,
    get: () => {
      if (!loaded) throw new Error('PropertyNotLoaded: changeTrackingMode');
      return mode;
    },
    set: (value) => { mode = value; },
  });
  const result = await applyChunkResults([{ chunkId: 'section', status: 'fulfilled',
    amendment: kind === 'paragraph' ? amendedTexts.join('\n') : 'Revised fallback paragraph.',
    chunk: { id: 'section', ...(kind === 'paragraph' ? { paragraphs: storedTexts.map((text) => ({ text })) } : {}) } }],
  new Map([['section', '_section']]), { trackChangesEnabled: true, lineDiffEnabled: false, log: jest.fn() });

  expect(result.errors).toEqual([]);
  expect(result.amendmentsApplied).toBe(1);
  expect(document.context.document.load).toHaveBeenCalledWith('changeTrackingMode');
  expect(document.context.document.changeTrackingMode).toBe('TrackMine');
});

test.each(['paragraph', 'fallback'])('%s application refuses writes when the host cannot enable requested tracked changes', async (kind) => {
  const document = kind === 'paragraph' ? nativeDocument(storedTexts) : fallbackDocument();
  delete global.Word.ChangeTrackingMode;
  const result = await applyChunkResults([{ chunkId: 'section', status: 'fulfilled',
    amendment: kind === 'paragraph' ? amendedTexts.join('\n') : 'Revised fallback paragraph.',
    chunk: { id: 'section', ...(kind === 'paragraph' ? { paragraphs: storedTexts.map((text) => ({ text })) } : {}) } }],
  new Map([['section', '_section']]), { trackChangesEnabled: true, lineDiffEnabled: false, log: jest.fn() });

  expect(result).toMatchObject({ amendmentsApplied: 0, appliedChunkIds: [], failedChunkIds: ['section'],
    partialChunkIds: [], uncertainChunkIds: [] });
  expect(result.errors[0]).toContain('cannot enable tracked changes');
  expect(applyTokenMapStrategy).not.toHaveBeenCalled();
  expect(document.range.insertText).not.toHaveBeenCalled();
});

test.each([false, true])('the amendment comment pass restores caller tracking mode even after comment failure=%s', async (commentFails) => {
  const document = nativeDocument(storedTexts);
  document.context.document.changeTrackingMode = 'TrackMine';
  document.range.insertComment = jest.fn(() => {
    expect(document.context.document.changeTrackingMode).toBe('Off');
    if (commentFails) throw new Error('Host comment insertion failed.');
  });
  const result = await applyChunkResults([{ chunkId: 'section', status: 'fulfilled', amendment: amendedTexts.join('\n'),
    comment: 'Review the revised wording.', chunk: { id: 'section', paragraphs: storedTexts.map((text) => ({ text })) } }],
  new Map([['section', '_section']]), { trackChangesEnabled: true, lineDiffEnabled: false, log: jest.fn() });

  expect(result.amendmentsApplied).toBe(1);
  expect(result.commentsInserted).toBe(commentFails ? 0 : 1);
  expect(document.range.insertComment).toHaveBeenCalledTimes(1);
  expect(document.context.document.changeTrackingMode).toBe('TrackMine');
});

test.each(['structural', 'fallback'])('%s application confirms tracking before queuing any text mutations', async (kind) => {
  const document = kind === 'structural' ? nativeDocument(storedTexts) : fallbackDocument();
  document.context.sync.mockImplementation(async () => {
    if (document.context.document.changeTrackingMode === 'TrackAll') throw new Error('Tracking mode rejected.');
  });
  const result = await applyChunkResults([{ chunkId: 'section', status: 'fulfilled',
    amendment: 'An entirely unrelated substitute paragraph.',
    chunk: { id: 'section', ...(kind === 'structural' ? { paragraphs: storedTexts.map((text) => ({ text })) } : {}) } }],
  new Map([['section', '_section']]), { trackChangesEnabled: true, lineDiffEnabled: false, log: jest.fn() });

  expect(result).toMatchObject({ amendmentsApplied: 0, partialChunkIds: [], uncertainChunkIds: [] });
  expect(result.errors[0]).toContain('Could not set Word\'s revision mode');
  expect(applyTokenMapStrategy).not.toHaveBeenCalled();
  expect(document.range.insertText).not.toHaveBeenCalled();
  for (const paragraph of document.paragraphs || []) {
    expect(paragraph.delete).not.toHaveBeenCalled();
    expect(paragraph.insertParagraph).not.toHaveBeenCalled();
  }
  expect(document.context.document.changeTrackingMode).toBe('Off');
});

test('a multi-paragraph section never falls back to a whole-range rewrite that would delete blank spacers', async () => {
  // Blank spacer between the two stored paragraphs: present in Word, absent
  // from the amendment (the parser skips it).
  const document = nativeDocument([storedTexts[0], '', storedTexts[1]]);
  document.contentRanges[2].getOoxml.mockImplementation(() => { throw new Error('GeneralException'); });
  const { result, log, onChunkApplied } = await apply();

  expect(result).toMatchObject({ amendmentsApplied: 0, appliedParagraphs: 0, appliedChunkIds: [],
    failedChunkIds: ['section'], partialChunkIds: [], uncertainChunkIds: [] });
  expect(result.errors[0]).toMatch(/re-anchor check failed \(GeneralException\)/);
  expect(applyTokenMapStrategy).not.toHaveBeenCalled();
  expect(applySentenceDiffStrategy).not.toHaveBeenCalled();
  expect(document.range.insertText).not.toHaveBeenCalled();
  expect(document.states.map((state) => state.text)).toEqual([storedTexts[0], '', storedTexts[1]]);
  expect(log).not.toHaveBeenCalledWith(expect.stringContaining('falling back to range-level'), 'warning');
  expect(onChunkApplied).toHaveBeenCalledWith('section', { applied: false, error: true, partial: false, uncertain: false });
  expect(document.context.document.changeTrackingMode).toBe('Off');
});

test('a range without a paragraph collection refuses the fallback when its text spans paragraphs', async () => {
  const document = fallbackDocument('TrackMine');
  document.range.getOoxml.mockImplementation(() => ({ value: `<w:document xmlns:w="${W}"><w:body>`
    + '<w:p><w:r><w:t>First paragraph.</w:t></w:r></w:p><w:p/><w:p><w:r><w:t>Second paragraph.</w:t></w:r></w:p>'
    + '</w:body></w:document>' }));
  const result = await applyChunkResults([{ chunkId: 'section', status: 'fulfilled',
    amendment: 'First paragraph.\nSecond paragraph, revised.', chunk: { id: 'section' } }],
  new Map([['section', '_section']]), { trackChangesEnabled: true, lineDiffEnabled: false, log: jest.fn() });

  expect(result).toMatchObject({ amendmentsApplied: 0, failedChunkIds: ['section'], partialChunkIds: [], uncertainChunkIds: [] });
  expect(result.errors[0]).toContain('multi-paragraph section cannot fall back');
  expect(applyTokenMapStrategy).not.toHaveBeenCalled();
  expect(document.range.insertText).not.toHaveBeenCalled();
  expect(document.context.document.changeTrackingMode).toBe('TrackMine');
});

test('a pre-1.4 host refuses tracked writes even though office.js defines the tracking enum', async () => {
  const document = nativeDocument(storedTexts);
  global.Office = { context: { requirements: { isSetSupported: (name, version) => version !== '1.4' } } };
  try {
    const { result } = await apply();
    expect(result).toMatchObject({ amendmentsApplied: 0, failedChunkIds: ['section'], uncertainChunkIds: [] });
    expect(result.errors[0]).toContain('cannot enable tracked changes');
    expect(applyTokenMapStrategy).not.toHaveBeenCalled();
    expect(document.states.map((state) => state.text)).toEqual(storedTexts);
  } finally {
    delete global.Office;
  }
});
