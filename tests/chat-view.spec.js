/** @jest-environment jsdom */

/**
 * Chat view tests: the model activity region auto-scrolls while streaming.
 * It follows the output only while the user stays near the bottom; scrolling
 * up disengages the follow until the user scrolls back down.
 */

const {
  initChatView,
  createAssistantMessage,
  addUserMessage,
  getCurrentSession,
  setCurrentSession,
  renderHistory,
  clearSessionMessages,
} = require('../src/taskpane/ui/chat-view.js');

function setupDom() {
  document.body.innerHTML = '<div id="chatMessages"></div><div id="welcome"></div>';
  initChatView();
}

/** jsdom reports zero scroll metrics; define real ones for the region. */
function mockScrollMetrics(el, { scrollHeight = 1000, clientHeight = 200 } = {}) {
  Object.defineProperty(el, 'scrollHeight', { configurable: true, value: scrollHeight });
  Object.defineProperty(el, 'clientHeight', { configurable: true, value: clientHeight });
}

describe('model activity auto-scroll', () => {
  test('follows the stream by default (scrollTop pinned to bottom)', () => {
    setupDom();
    const msg = createAssistantMessage();
    const modelBody = msg.el.querySelector('.msg-model-body');
    mockScrollMetrics(modelBody);

    msg.appendModelToken({ id: 's' }, 'content', 'hello');
    expect(modelBody.scrollTop).toBe(1000);

    msg.appendModelToken({ id: 's' }, 'content', ' world');
    expect(modelBody.scrollTop).toBe(1000);
  });

  test('stops following when the user scrolls up, resumes near the bottom', () => {
    setupDom();
    const msg = createAssistantMessage();
    const modelBody = msg.el.querySelector('.msg-model-body');
    mockScrollMetrics(modelBody);

    msg.appendModelToken({ id: 's' }, 'content', 'one');
    expect(modelBody.scrollTop).toBe(1000);

    // User scrolls up to read earlier output -> follow disengages.
    modelBody.scrollTop = 100;
    modelBody.dispatchEvent(new Event('scroll'));
    msg.appendModelToken({ id: 's' }, 'content', 'two');
    expect(modelBody.scrollTop).toBe(100);

    // User scrolls back near the bottom -> follow re-engages.
    modelBody.scrollTop = 980;
    modelBody.dispatchEvent(new Event('scroll'));
    msg.appendModelToken({ id: 's' }, 'content', 'three');
    expect(modelBody.scrollTop).toBe(1000);
  });

  test('re-expanding the region jumps to the latest output', () => {
    setupDom();
    const msg = createAssistantMessage();
    const modelBody = msg.el.querySelector('.msg-model-body');
    mockScrollMetrics(modelBody);

    msg.appendModelToken({ id: 's' }, 'content', 'hello');
    msg.collapseModelOutput();
    expect(modelBody.style.display).toBe('none');

    modelBody.scrollTop = 0;
    msg.el.querySelector('.msg-model-toggle').click();
    expect(modelBody.style.display).not.toBe('none');
    expect(modelBody.scrollTop).toBe(1000);
  });

  test('no auto-scroll while the region is collapsed', () => {
    setupDom();
    const msg = createAssistantMessage();
    const modelBody = msg.el.querySelector('.msg-model-body');
    mockScrollMetrics(modelBody);

    msg.appendModelToken({ id: 's' }, 'content', 'hello');
    msg.collapseModelOutput();
    modelBody.scrollTop = 0;

    msg.appendModelToken({ id: 's' }, 'content', 'more');
    expect(modelBody.scrollTop).toBe(0);
  });
});

describe('current-session tracking', () => {
  beforeEach(() => {
    setupDom();
    clearSessionMessages();
  });

  test('addUserMessage appends to the current session', () => {
    addUserMessage('Hello');
    const session = getCurrentSession();
    expect(session.messages).toHaveLength(1);
    expect(session.messages[0]).toMatchObject({ role: 'user', text: 'Hello' });
    expect(session.messages[0].id).toBeDefined();
    expect(session.messages[0].ts).toBeDefined();
  });

  test('createAssistantMessage finalizes the assistant record on finalizeForHistory', () => {
    addUserMessage('What is X?');
    const msg = createAssistantMessage();
    msg.setStatus('Reading the document...');
    msg.setText('X is a placeholder.');
    msg.collapseLog();
    msg.collapseModelOutput();
    msg.finalizeForHistory();

    const session = getCurrentSession();
    expect(session.messages).toHaveLength(2);
    expect(session.messages[1]).toMatchObject({
      role: 'assistant',
      text: 'X is a placeholder.',
      status: 'Reading the document...',
      error: null,
    });
  });

  test('markError captures the error string in the assistant record', () => {
    const msg = createAssistantMessage();
    msg.markError('boom');
    msg.collapseLog();
    msg.collapseModelOutput();
    msg.finalizeForHistory();

    const session = getCurrentSession();
    expect(session.messages[0].error).toBe('boom');
    expect(session.messages[0].text).toContain('Error: boom');
  });

  test('finalizeForHistory captures worklog and model collapse summaries', () => {
    const msg = createAssistantMessage();
    msg.appendLogLine('first step');
    msg.appendLogLine('second step');
    msg.appendModelToken({ id: 's1' }, 'content', 'a');
    msg.appendModelToken({ id: 's2' }, 'content', 'b');
    msg.collapseLog();
    msg.collapseModelOutput();
    msg.finalizeForHistory();

    const session = getCurrentSession();
    const m = session.messages[0];
    expect(m.worklog).toEqual(expect.objectContaining({ count: 2 }));
    expect(m.worklog.durationMs).toBeGreaterThanOrEqual(0);
    expect(m.model).toEqual({ sections: 2 });
  });

  test('attachProposal with meta records terminal state from the proposal card', () => {
    const { createProposalCard } = require('../src/taskpane/ui/proposal-card.js');
    const msg = createAssistantMessage();
    const card = createProposalCard({
      title: 'Proposed edits',
      onApply: async () => {},
      onReject: () => {},
    });
    const meta = {
      title: 'Proposed edits',
      state: 'pending',
      countsText: '',
      items: [{ id: 's1', label: 'Section one', before: 'old', after: 'new' }],
    };
    msg.attachProposal(card, meta);
    card.markApplied();

    msg.collapseLog();
    msg.collapseModelOutput();
    msg.finalizeForHistory();

    const session = getCurrentSession();
    expect(session.messages[0].proposals).toHaveLength(1);
    expect(session.messages[0].proposals[0].state).toBe('applied');
  });

  test('clearSessionMessages wipes the array but keeps the welcome state', () => {
    addUserMessage('hi');
    expect(getCurrentSession().messages).toHaveLength(1);
    clearSessionMessages();
    expect(getCurrentSession().messages).toEqual([]);
  });
});

describe('renderHistory / setCurrentSession', () => {
  beforeEach(() => {
    setupDom();
    clearSessionMessages();
  });

  test('renderHistory rebuilds user + assistant message DOM from a saved array', () => {
    const messages = [
      { id: 'm-0', role: 'user', text: 'old question', ts: '2026-01-01T00:00:00.000Z' },
      {
        id: 'm-1', role: 'assistant',
        text: 'old answer',
        status: 'Done',
        error: null,
        worklog: null,
        model: null,
        citations: [],
        proposals: [],
        ts: '2026-01-01T00:00:01.000Z',
      },
    ];
    renderHistory(messages);

    const bubbles = document.querySelectorAll('#chatMessages .chat-message');
    expect(bubbles).toHaveLength(2);
    expect(bubbles[0].classList.contains('chat-message-user')).toBe(true);
    expect(bubbles[0].textContent).toBe('old question');
    expect(bubbles[1].classList.contains('chat-message-assistant')).toBe(true);
    expect(bubbles[1].textContent).toContain('old answer');
  });

  test('renderHistory hides the welcome state', () => {
    const messages = [{ id: 'm-0', role: 'user', text: 'q', ts: '2026-01-01T00:00:00.000Z' }];
    renderHistory(messages);
    expect(document.getElementById('welcome').style.display).toBe('none');
  });

  test('renderHistory renders assistant error inline', () => {
    const messages = [
      { id: 'm-0', role: 'user', text: 'q', ts: '2026-01-01T00:00:00.000Z' },
      { id: 'm-1', role: 'assistant', text: 'Error: bad', status: '', error: 'bad', worklog: null, model: null, citations: [], proposals: [], ts: '2026-01-01T00:00:01.000Z' },
    ];
    renderHistory(messages);
    const a = document.querySelectorAll('#chatMessages .chat-message')[1];
    expect(a.classList.contains('chat-message-error')).toBe(true);
    expect(a.textContent).toContain('bad');
  });

  test('setCurrentSession replaces DOM and seeds the live messages array', () => {
    const session = {
      id: 's-1',
      title: 'Demo',
      createdAt: '2026-01-01T00:00:00.000Z',
      updatedAt: '2026-01-01T00:00:05.000Z',
      messages: [
        { id: 'm-0', role: 'user', text: 'first', ts: '2026-01-01T00:00:00.000Z' },
        { id: 'm-1', role: 'assistant', text: 'answer', status: '', error: null, worklog: null, model: null, citations: [], proposals: [], ts: '2026-01-01T00:00:01.000Z' },
      ],
    };
    setCurrentSession(session);

    // Live array mirrors the session
    expect(getCurrentSession().id).toBe('s-1');
    expect(getCurrentSession().title).toBe('Demo');
    expect(getCurrentSession().messages.map((m) => m.text)).toEqual(['first', 'answer']);

    // DOM rebuilt
    const bubbles = document.querySelectorAll('#chatMessages .chat-message');
    expect(bubbles).toHaveLength(2);
  });

  test('new addUserMessage after setCurrentSession appends to the loaded session', () => {
    const session = {
      id: 's-2', title: 't', createdAt: 'x', updatedAt: 'y',
      messages: [{ id: 'm-0', role: 'user', text: 'old', ts: '2026-01-01T00:00:00.000Z' }],
    };
    setCurrentSession(session);
    addUserMessage('new follow-up');

    const messages = getCurrentSession().messages;
    expect(messages.map((m) => m.text)).toEqual(['old', 'new follow-up']);
  });
});

describe('proposal state persistence', () => {
  const { createProposalCard } = require('../src/taskpane/ui/proposal-card.js');
  const { setProposalStateChangeHandler, renderStaticProposalCard } = require('../src/taskpane/ui/chat-view.js');

  function attachCard(msg, meta) {
    const card = createProposalCard({
      title: meta.title,
      onApply: jest.fn(async () => {}),
      onReject: jest.fn(),
    });
    msg.attachProposal(card, meta);
    return card;
  }

  beforeEach(() => {
    setupDom();
    clearSessionMessages();
    setProposalStateChangeHandler(null);
  });
  afterEach(() => setProposalStateChangeHandler(null));

  test('reject click flows through markRejected so meta leaves "pending"', () => {
    const msg = createAssistantMessage();
    const meta = { title: 'Proposed table', state: 'pending', countsText: '', items: [] };
    const card = attachCard(msg, meta);
    card.el.querySelector('.btn-secondary').click();
    expect(meta.state).toBe('rejected');
  });

  test('late settle after finalizeForHistory re-syncs the record and notifies', () => {
    const msg = createAssistantMessage();
    const meta = { title: 'Proposed table', state: 'pending', countsText: '', items: [] };
    const card = attachCard(msg, meta);
    msg.finalizeForHistory();

    // The pre-settle snapshot is pending; the click arrives afterwards.
    expect(getCurrentSession().messages[0].proposals[0].state).toBe('pending');

    const onStateChange = jest.fn();
    setProposalStateChangeHandler(onStateChange);
    card.markApplied();

    const record = getCurrentSession().messages[0];
    expect(record.proposals[0].state).toBe('applied');
    expect(onStateChange).toHaveBeenCalledTimes(1);
  });

  test('markError also syncs and can be superseded by a later apply', () => {
    const msg = createAssistantMessage();
    const meta = { title: 't', state: 'pending', countsText: '', items: [] };
    const card = attachCard(msg, meta);
    msg.finalizeForHistory();

    card.markError('boom');
    expect(getCurrentSession().messages[0].proposals[0].state).toBe('error');
    card.markWarning('nothing applied');
    expect(getCurrentSession().messages[0].proposals[0].state).toBe('warning');
    expect(getCurrentSession().messages[0].proposals[0].detail).toBe('nothing applied');
  });

  test('tablePreview is serialized into history and rendered by the static card', () => {
    const msg = createAssistantMessage();
    const meta = {
      title: 'Proposed table',
      state: 'pending',
      countsText: '3 × 3 table',
      items: [],
      tablePreview: { rows: [['<b>x</b>', '']], headerRowCount: 1, style: 'tableGrid' },
    };
    attachCard(msg, meta);
    msg.finalizeForHistory();

    const saved = getCurrentSession().messages[0].proposals[0];
    expect(saved.tablePreview.rows).toEqual([['<b>x</b>', '']]);
    expect(saved.tablePreview.headerRowCount).toBe(1);

    const staticCard = renderStaticProposalCard(saved);
    const cell = staticCard.querySelector('.proposal-card-table th');
    expect(cell.textContent).toBe('<b>x</b>');
    expect(staticCard.querySelector('.proposal-card-table b')).toBeNull();
  });
});

