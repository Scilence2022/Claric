/**
 * Tests for applyDocumentAppend — the apply half of the `append` capability
 * (generated long-form text added at the document end).
 *
 * Covers:
 * - Generated text becomes one end-of-body paragraph per non-empty line
 * - Empty or whitespace-only model output is refused and nothing is inserted
 * - Change tracking follows the setting during the write and is restored off
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
function makeAppendContext() {
  const inserted = [];
  const trackingModes = [];
  const context = {
    document: {
      body: {
        insertParagraph: jest.fn((text, location) => inserted.push({ text, location })),
      },
    },
    sync: jest.fn().mockResolvedValue(undefined),
  };
  let mode = null;
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

  test('tracks the write when change tracking is on and restores it off', async () => {
    const { trackingModes } = makeAppendContext();
    await applyDocumentAppend(makeDeps({ trackChangesEnabled: true }), { generatedText: 'Text.' });
    expect(trackingModes).toEqual(['TrackAll', 'Off']);
  });

  test('writes untracked when change tracking is off', async () => {
    const { trackingModes } = makeAppendContext();
    await applyDocumentAppend(makeDeps({ trackChangesEnabled: false }), { generatedText: 'Text.' });
    expect(trackingModes).toEqual(['Off', 'Off']);
  });
});
