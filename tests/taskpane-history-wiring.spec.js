/** @jest-environment jsdom */
const fs = require('node:fs');
const path = require('node:path');

jest.mock('../src/taskpane/conversation.js', () => ({ createConversation: jest.fn() }));
jest.mock('../src/taskpane/word-actions.js', () => ({ watchSelection: jest.fn(), revealTextSnippet: jest.fn() }));
jest.mock('../src/lib/reassembler.js', () => ({ reapOrphanChunkBookmarks: jest.fn() }));
jest.mock('../src/taskpane/settings-loader.js', () => ({
  createSettingsLoader: () => ({ open: jest.fn(), testConnection: jest.fn() }),
}));
jest.mock('../src/taskpane/ui/input-bar.js', () => ({ initInputBar: jest.fn() }));
jest.mock('../src/taskpane/ui/history-view.js', () => ({ initHistoryView: jest.fn(), openHistory: jest.fn() }));

const { createConversation } = require('../src/taskpane/conversation.js');
const { initInputBar } = require('../src/taskpane/ui/input-bar.js');
const { initHistoryView } = require('../src/taskpane/ui/history-view.js');
const chatView = require('../src/taskpane/ui/chat-view.js');

it('invalidates the outgoing conversation before restoring a selected history session', () => {
  const html = fs.readFileSync(path.join(__dirname, '../src/taskpane/taskpane.html'), 'utf8');
  document.body.innerHTML = html.slice(html.indexOf('<body>') + 6, html.indexOf('</body>'));
  localStorage.clear();
  const focus = jest.fn();
  initInputBar.mockReturnValue({ focus });
  const events = [];
  createConversation.mockReturnValue({
    newChat: jest.fn(() => {
      events.push('invalidate');
      chatView.clearChat();
    }),
  });
  const setCurrent = jest.spyOn(chatView, 'setCurrentSession');
  const originalSetCurrent = setCurrent.getMockImplementation();
  setCurrent.mockImplementation((session) => {
    events.push('restore');
    originalSetCurrent(session);
  });
  global.Office = {
    HostType: { Word: 'Word' },
    onReady: (callback) => callback({ host: 'Word' }),
    context: { requirements: { isSetSupported: () => true } },
  };
  global.fetch = jest.fn(async () => ({ ok: false }));
  try {
    require('../src/taskpane/taskpane.js');
    const session = { id: 'restored-session', messages: [{ role: 'user', text: 'Earlier question' }] };
    initHistoryView.mock.calls[0][0].onLoadSession(session);
    expect(events).toEqual(['invalidate', 'restore']);
    expect(chatView.getCurrentSession().id).toBe(session.id);
    expect(focus).toHaveBeenCalledTimes(2);
  } finally {
    setCurrent.mockRestore();
    delete global.Office;
    delete global.fetch;
  }
});
