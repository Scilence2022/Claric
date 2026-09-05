const http = require('node:http');
const { JSDOM } = require('jsdom');

jest.mock('../src/lib/comment-extractor.js', () => ({
  ...jest.requireActual('../src/lib/comment-extractor.js'),
  extractDocumentStructured: jest.fn(),
}));

const dom = new JSDOM('<div id="chatMessages"></div><div id="welcome"></div>', {
  url: 'http://localhost',
});
global.window = dom.window;
global.document = dom.window.document;
global.localStorage = dom.window.localStorage;

const { createConversation } = require('../src/taskpane/conversation.js');
const chatView = require('../src/taskpane/ui/chat-view.js');
const { saveSession, loadSession } = require('../src/taskpane/sessions.js');
const { extractDocumentStructured } = require('../src/lib/comment-extractor.js');

function textOf(message) {
  return typeof message.content === 'string' ? message.content
    : message.content.filter((part) => part.type === 'text').map((part) => part.text).join('\n');
}

describe('conversation submit to HTTP request continuity', () => {
  let server;
  let url;
  let requests;
  let streamResponses;
  let delayedReceived;
  let delayedClosed;

  beforeAll(async () => {
    server = http.createServer(async (request, response) => {
      let raw = '';
      for await (const chunk of request) raw += chunk;
      const body = JSON.parse(raw);
      requests.push(body);
      if (textOf(body.messages[body.messages.length - 1]).includes('DELAYED-OLD-TURN')) {
        response.on('close', () => delayedClosed?.());
        delayedReceived?.(response);
        return;
      }
      const isPlanner = textOf(body.messages[body.messages.length - 1]).includes('You are the task planner');
      const answer = isPlanner
        ? JSON.stringify([{ type: 'qa', instruction: 'Explain the second ALPHA-27 recommendation.' }])
        : `Answer ${requests.length}: proposal reference ALPHA-27.`;
      const anthropic = request.url.endsWith('/messages');
      if (!streamResponses) {
        response.writeHead(200, { 'content-type': 'application/json' });
        response.end(JSON.stringify(anthropic
          ? { content: [{ type: 'text', text: answer }], stop_reason: 'end_turn' }
          : { choices: [{ message: { content: answer }, finish_reason: 'stop' }] }));
        return;
      }
      response.writeHead(200, { 'content-type': 'text/event-stream' });
      if (anthropic) {
        response.write(`data: ${JSON.stringify({ type: 'content_block_delta', delta: { type: 'text_delta', text: answer } })}\n\n`);
        response.end('data: {"type":"message_stop"}\n\n');
      } else {
        response.write(`data: ${JSON.stringify({ choices: [{ delta: { content: answer } }] })}\n\n`);
        response.end('data: [DONE]\n\n');
      }
    });
    await new Promise((resolve) => server.listen(0, '127.0.0.1', resolve));
    url = `http://127.0.0.1:${server.address().port}`;
  });

  afterAll(async () => {
    await new Promise((resolve) => server.close(resolve));
    dom.window.close();
    delete global.window;
    delete global.document;
    delete global.localStorage;
    delete global.Word;
  });

  beforeEach(() => {
    requests = [];
    localStorage.clear();
    document.body.innerHTML = '<div id="chatMessages"></div><div id="welcome"></div>';
    chatView.initChatView();
    chatView.clearChat();
    extractDocumentStructured.mockImplementation(async () => `FRESH-DOCUMENT-${requests.length + 1}`);
    global.Word = {
      run: async (callback) => callback({
        sync: async () => {},
        document: { getSelection: () => ({ text: '', load: () => {}, paragraphs: { items: [], load: () => {} } }) },
      }),
    };
  });

  test.each([
    ['openai', true], ['openai', false], ['claude', true], ['claude', false],
  ])('%s retains three turns and restored history (SSE: %s)', async (provider, streaming) => {
    streamResponses = streaming;
    const appState = {
      config: { backend: provider, providers: { [provider]: { url, model: 'test-model', apiPath: '/v1' } } },
      promptManager: { getPrompts: () => [], getActivePrompt: () => null },
    };
    const log = jest.fn();
    const conversation = createConversation({
      appState,
      view: { ...chatView, renderWelcome: chatView.showWelcome },
      input: { setProcessing: () => {}, setValue: () => {}, focus: () => {} },
      log,
      getSelectionText: async () => '',
      onTurnCommitted: (session) => {
        if (session?.messages.length) saveSession(session.messages, { id: session.id });
      },
    });
    const questions = [
      'What are your two recommendations for project ALPHA-27?',
      'What was your second recommendation?',
      'How does that relate to my first question?',
    ];
    for (const question of questions) await conversation.submit(question);

    expect(log.mock.calls.filter(([, level]) => level === 'error')).toEqual([]);
    expect(requests).toHaveLength(3);
    expect(requests[0].messages).toHaveLength(1);
    expect(requests[1].messages.map((message) => message.role)).toEqual(['user', 'assistant', 'user']);
    const third = requests[2].messages;
    expect(third.map((message) => message.role)).toEqual(['user', 'assistant', 'user', 'assistant', 'user']);
    expect(textOf(third[0])).toBe(questions[0]);
    expect(textOf(third[1])).toContain('Answer 1: proposal reference ALPHA-27.');
    expect(textOf(third[2])).toBe(questions[1]);
    expect(textOf(third[3])).toContain('Answer 2: proposal reference ALPHA-27.');
    expect(textOf(third[4])).toContain(questions[2]);
    expect(third.filter((message) => textOf(message).includes(questions[2]))).toHaveLength(1);
    expect(JSON.stringify(third)).not.toContain('FRESH-DOCUMENT-1');
    expect(JSON.stringify(third)).not.toContain('FRESH-DOCUMENT-2');
    expect(textOf(third[4])).toContain('FRESH-DOCUMENT-3');

    const storedId = chatView.getCurrentSession().id;
    expect(loadSession(storedId).messages).toHaveLength(6);
    conversation.newChat();
    await conversation.submit('What is the topic of this new conversation?');
    expect(requests[3].messages).toHaveLength(1);
    expect(JSON.stringify(requests[3])).not.toContain('ALPHA-27');

    conversation.newChat();
    chatView.setCurrentSession(loadSession(storedId));
    await conversation.submit('What did we decide about ALPHA-27 earlier?');
    expect(requests[4].messages.map((message) => message.role))
      .toEqual(['user', 'assistant', 'user', 'assistant', 'user', 'assistant', 'user']);
    expect(textOf(requests[4].messages[0])).toBe(questions[0]);
    expect(textOf(requests[4].messages[5])).toContain('Answer 3:');
    expect(JSON.stringify(requests[4])).not.toContain('topic of this new conversation');
    expect(chatView.getCurrentSession().id).toBe(storedId);
  });

  test('switching during HTTP cancels the old transport and keeps restored history isolated', async () => {
    streamResponses = true;
    const conversation = createConversation({
      appState: {
        config: { backend: 'openai', providers: { openai: { url, model: 'test-model', apiPath: '/v1' } } },
        promptManager: { getPrompts: () => [], getActivePrompt: () => null },
      },
      view: { ...chatView, renderWelcome: chatView.showWelcome },
      input: { setProcessing: () => {}, setValue: () => {}, focus: () => {} },
      log: jest.fn(),
      getSelectionText: async () => '',
      onTurnCommitted: (session) => {
        if (session?.messages.length) saveSession(session.messages, { id: session.id });
      },
    });
    const restored = saveSession([
      { role: 'user', text: 'Earlier RESTORED-TOPIC question' },
      { role: 'assistant', text: 'Earlier RESTORED-TOPIC answer' },
    ], { id: 'restored-target' });
    const received = new Promise((resolve) => { delayedReceived = resolve; });
    const closed = new Promise((resolve) => { delayedClosed = resolve; });
    const pending = conversation.submit('What about DELAYED-OLD-TURN?');
    const oldResponse = await received;
    const oldId = chatView.getCurrentSession().id;
    conversation.newChat();
    chatView.setCurrentSession(restored);
    await pending;
    await closed;
    expect(oldResponse.destroyed).toBe(true);
    expect(chatView.getCurrentSession().messages).toHaveLength(2);
    expect(loadSession(oldId).messages).toHaveLength(1);
    expect(loadSession(restored.id).messages).toHaveLength(2);
    await conversation.submit('What was my RESTORED-TOPIC question?');
    expect(requests).toHaveLength(2);
    expect(requests[1].messages.map((message) => message.role)).toEqual(['user', 'assistant', 'user']);
    expect(textOf(requests[1].messages[0])).toBe('Earlier RESTORED-TOPIC question');
    expect(textOf(requests[1].messages[1])).toBe('Earlier RESTORED-TOPIC answer');
    expect(JSON.stringify(requests[1])).not.toContain('DELAYED-OLD-TURN');
    expect(loadSession(restored.id).messages).toHaveLength(4);
    delayedReceived = null;
    delayedClosed = null;
  });

  test('a referential follow-up reaches both the planner and its executor with prior turns', async () => {
    streamResponses = true;
    const conversation = createConversation({
      appState: {
        config: { backend: 'openai', providers: { openai: { url, model: 'test-model', apiPath: '/v1' } } },
        promptManager: { getPrompts: () => [], getActivePrompt: () => null },
      },
      view: { ...chatView, renderWelcome: chatView.showWelcome },
      input: { setProcessing: () => {}, setValue: () => {}, focus: () => {} },
      log: jest.fn(),
      getSelectionText: async () => '',
    });
    await conversation.submit('What are your recommendations for ALPHA-27?');
    await conversation.submit('The second one, please');
    expect(requests).toHaveLength(3);
    for (const request of requests.slice(1)) {
      expect(request.messages.map((message) => message.role)).toEqual(['user', 'assistant', 'user']);
      expect(textOf(request.messages[0])).toBe('What are your recommendations for ALPHA-27?');
      expect(textOf(request.messages[1])).toContain('Answer 1: proposal reference ALPHA-27.');
    }
    expect(textOf(requests[1].messages[2])).toContain('You are the task planner');
    expect(textOf(requests[1].messages[2])).toContain('The second one, please');
    expect(textOf(requests[2].messages[2])).toContain('Explain the second ALPHA-27 recommendation.');
    expect(chatView.getCurrentSession().messages).toHaveLength(4);
  });

  test.each(['openai', 'claude'])('%s preserves role order on non-streaming transport', async (provider) => {
    const { sendMessages } = require('../src/lib/llm-client.js');
    streamResponses = false;
    const messages = [
      { role: 'system', content: 'Current trusted persona' },
      { role: 'user', content: 'Remember project ALPHA-27' },
      { role: 'assistant', content: 'I will refer to ALPHA-27.' },
      { role: 'user', content: 'Which project did I mention?' },
    ];
    const reply = await sendMessages({ url, provider, apiPath: '/v1', model: 'test-model' }, messages);
    expect(reply).toContain('ALPHA-27');
    const wire = requests[0];
    const conversation = provider === 'claude' ? wire.messages : wire.messages.slice(1);
    expect(conversation.map((message) => message.role)).toEqual(['user', 'assistant', 'user']);
    expect(conversation.map(textOf)).toEqual(messages.slice(1).map(textOf));
    if (provider === 'claude') expect(wire.system).toBe('Current trusted persona');
    else expect(wire.messages[0]).toEqual(messages[0]);
  });
});
