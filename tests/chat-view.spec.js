/** @jest-environment jsdom */

/**
 * Chat view tests: the model activity region auto-scrolls while streaming.
 * It follows the output only while the user stays near the bottom; scrolling
 * up disengages the follow until the user scrolls back down.
 */

const { initChatView, createAssistantMessage } = require('../src/taskpane/ui/chat-view.js');

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
