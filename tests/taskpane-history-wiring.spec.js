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
const { recordNativeTextMismatch } = require('../src/lib/native-text-diagnostics.js');

it('copies only the active conversation when opening history or starting a new chat', async () => {
  const html = fs.readFileSync(path.join(__dirname, '../src/taskpane/taskpane.html'), 'utf8');
  document.body.innerHTML = html.slice(html.indexOf('<body>') + 6, html.indexOf('</body>'));
  localStorage.clear();
  const focus = jest.fn();
  initInputBar.mockReturnValue({ focus });
  const events = [];
  let contextScope = 0;
  createConversation.mockReturnValue({
    getContextScope: () => contextScope,
    newChat: jest.fn(() => {
      events.push('invalidate');
      contextScope++;
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
  const writeText = jest.fn(async () => {});
  Object.defineProperty(navigator, 'clipboard', { configurable: true, value: { writeText } });
  const copyContext = async () => {
    document.getElementById('copyContextBtn').click();
    await new Promise((resolve) => setTimeout(resolve, 0));
    return writeText.mock.lastCall[0];
  };
  try {
    require('../src/taskpane/taskpane.js');
    await new Promise((resolve) => setTimeout(resolve, 0));
    expect(global.fetch.mock.calls.some(([url]) => String(url).includes('/coordination'))).toBe(true);
    const { getSessionLogging } = createConversation.mock.calls[0][0];
    const oldLog = getSessionLogging(contextScope);
    chatView.addUserMessage('Previous chat marker');
    oldLog.log('Previous activity marker', 'info');
    const mismatch = { value: '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body/></w:document>' };
    recordNativeTextMismatch(mismatch, { text: 'Previous diagnostic marker' }, { contextScope });
    const previous = await copyContext();
    expect(previous).toContain('Previous chat marker');
    expect(previous).toContain('Previous activity marker');
    expect(previous).toContain('Previous diagnostic marker');

    const session = { id: 'restored-session', messages: [{ role: 'user', text: 'Earlier question' }] };
    initHistoryView.mock.calls[0][0].onLoadSession(session);
    expect(events).toEqual(['invalidate', 'restore']);
    expect(chatView.getCurrentSession().id).toBe(session.id);
    expect(focus).toHaveBeenCalledTimes(2);
    oldLog.logWithRetry('Late previous activity marker', 'error', jest.fn());
    recordNativeTextMismatch(mismatch, { text: 'Late previous diagnostic marker' }, { contextScope: 0 });
    getSessionLogging(contextScope).log('Restored chat activity marker');
    const restored = await copyContext();
    expect(restored).toContain('Earlier question');
    expect(restored).toContain('Restored chat activity marker');
    expect(restored).not.toMatch(/previous (chat|activity|diagnostic) marker/i);

    document.getElementById('newChatBtn').click();
    const empty = await copyContext();
    expect(empty).toContain('## Current chat (including model activity and proposals)\n(empty)');
    expect(empty).toContain('## Activity log (current chat)\n(empty)');
    expect(empty).toContain('"wordApi": "1.8"');
    expect(empty).toContain('"writingSettings"');
    expect(empty).not.toMatch(/Earlier question|activity marker|diagnostic marker|Get started with/);
  } finally {
    setCurrent.mockRestore();
    delete global.Office;
    delete global.fetch;
  }
});
