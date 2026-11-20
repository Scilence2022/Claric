/**
 * Tests for applyDocumentAppend — the apply half of the `append` capability
 * (generated long-form text added at the document end).
 *
 * Covers:
 * - Generated text becomes one end-of-body paragraph per non-empty line
 * - Empty or whitespace-only model output is refused and nothing is inserted
 * - Change tracking follows the setting during the write and preserves the original Word mode
 */

jest.mock('../src/lib/llm-client.js', () => ({
  sendPrompt: jest.fn(),
  sendPromptStream: jest.fn(),
  stripMarkdown: jest.fn((t) => t),
  stripChunkDelimiters: jest.fn((t) => t),
}));

const { applyDocumentAppend } = require('../src/taskpane/word-actions.js');

function makeDeps({ trackChangesEnabled = true } = {}) {
  return {
    appState: { config: { trackChangesEnabled } },
    log: jest.fn(),
  };
}

/** Word.run mock whose body records end insertions and tracking-mode writes. */
function makeAppendContext(initialMode = 'Off') {
  const inserted = [];
  const trackingModes = [];
  const context = {
    document: {
      load: jest.fn(),
      body: {
        insertParagraph: jest.fn((text, location) => inserted.push({ text, location })),
      },
    },
    sync: jest.fn().mockResolvedValue(undefined),
  };
  let mode = initialMode;
  Object.defineProperty(context.document, 'changeTrackingMode', {
    get: () => mode,
    set: (v) => { mode = v; trackingModes.push(v); },
  });
  global.Word = {
    run: jest.fn((cb) => cb(context)),
    InsertLocation: { end: 'End' },
    ChangeTrackingMode: { trackAll: 'TrackAll', off: 'Off' },
  };
  return { context, inserted, trackingModes };
}

describe('applyDocumentAppend', () => {
  afterEach(() => { delete global.Word; });

  test('appends each non-empty line as a paragraph at the document end', async () => {
    const { inserted } = makeAppendContext();
    const deps = makeDeps();
    const result = await applyDocumentAppend(deps, {
      generatedText: 'First paragraph.\n\nSecond paragraph.\n',
    });
    expect(inserted).toEqual([
      { text: 'First paragraph.', location: 'End' },
      { text: 'Second paragraph.', location: 'End' },
    ]);
    expect(result).toEqual({ paragraphsAppended: 2, chars: 'First paragraph.\n\nSecond paragraph.'.length });
    expect(deps.log).toHaveBeenCalledWith('Appended 2 paragraph(s) to the document end.', 'success');
  });

  test.each([
    ['empty text', ''],
    ['whitespace-only text', '  \n\t \n'],
    ['a missing proposal', null],
  ])('refuses %s without touching the document', async (_label, generatedText) => {
    const { inserted, trackingModes } = makeAppendContext();
    const proposal = generatedText === null ? null : { generatedText };
    await expect(applyDocumentAppend(makeDeps(), proposal)).rejects.toThrow(/Nothing to append/);
    expect(global.Word.run).not.toHaveBeenCalled();
    expect(inserted).toEqual([]);
    expect(trackingModes).toEqual([]);
  });

  test('tracks the append and restores the original Word tracking mode', async () => {
    const { trackingModes, context } = makeAppendContext('TrackMineOnly');
    await applyDocumentAppend(makeDeps({ trackChangesEnabled: true }), { generatedText: 'Text.' });
    expect(context.document.load).toHaveBeenCalledWith('changeTrackingMode');
    expect(trackingModes).toEqual(['TrackAll', 'TrackMineOnly']);
  });

  test('restores the original Word tracking mode after a failed write sync', async () => {
    const { context, trackingModes } = makeAppendContext('TrackMineOnly');
    context.sync.mockResolvedValueOnce(undefined).mockRejectedValueOnce(new Error('Append failed'));
    await expect(applyDocumentAppend(makeDeps(), { generatedText: 'Text.' })).rejects.toThrow('Append failed');
    expect(trackingModes).toEqual(['TrackAll', 'TrackMineOnly']);
    expect(context.sync).toHaveBeenCalledTimes(3);
  });

  test('writes untracked when change tracking is off', async () => {
    const { trackingModes } = makeAppendContext();
    await applyDocumentAppend(makeDeps({ trackChangesEnabled: false }), { generatedText: 'Text.' });
    expect(trackingModes).toEqual(['Off', 'Off']);
  });
});
