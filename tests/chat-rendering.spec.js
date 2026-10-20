/** @jest-environment jsdom */

const { renderMarkdown } = require('../src/taskpane/ui/markdown.js');
const { initChatView, createAssistantMessage, addUserMessage, clearSessionMessages,
    getCurrentSession, setCurrentSession } = require('../src/taskpane/ui/chat-view.js');
const { createProposalCard } = require('../src/taskpane/ui/proposal-card.js');
const { normalizeMessage } = require('../src/taskpane/message-shape.js');

const answer = '## 总体评价\n\n本文是一篇**框架较完整、技术覆盖较广、对测量边界具有较好辨识度的综述初稿**。\n\n'
    + '### 主要优点\n\n- **科学表述较为审慎**\n- 对读数的区分\n\n'
    + '1. 统一分类逻辑\n2. 补强横向比较\n\n> 需要独立验证\n\n'
    + '`RNA/DNA` 比率\n\n```js\nconst x = "<script>";\n```\n\n'
    + '| 方法 | 读数 |\n| --- | --- |\n| MPRA | RNA/DNA |\n\n'
    + '[研究来源](https://example.com/paper)';

beforeEach(() => {
    document.body.innerHTML = '<div id="chatMessages"></div><div id="welcome"></div>';
    initChatView();
    clearSessionMessages();
});

test('renders the Chinese review, lists, code, tables, and links as formatted content', () => {
    const msg = createAssistantMessage();
    msg.setText(answer);
    const body = msg.el.querySelector('.msg-body');
    expect(body.querySelector('h2').textContent).toBe('总体评价');
    expect(body.querySelector('strong').textContent).toContain('框架较完整');
    expect(body.querySelectorAll('ul > li')).toHaveLength(2);
    expect(body.querySelectorAll('ol > li')).toHaveLength(2);
    expect(body.querySelector('blockquote').textContent).toContain('需要独立验证');
    expect(body.querySelector('pre code').textContent).toBe('const x = "<script>";\n');
    expect(body.querySelector('.msg-markdown-table[tabindex="0"] td').textContent).toBe('MPRA');
    expect(body.querySelector('a').rel).toBe('noopener noreferrer');
    expect(body.querySelector('a').target).toBe('_blank');
    expect(body.querySelector('script')).toBeNull();
});

test('strips active HTML, styles, remote images, and unsafe links from model output', () => {
    const element = document.createElement('div');
    renderMarkdown(element, '<script>alert(1)</script>\n\n'
        + '<style>body{display:none}</style><form><input autofocus></form>'
        + '<svg onload="alert(1)"></svg><iframe src="https://example.com"></iframe>'
        + '<p id="chatMessages" style="position:fixed" onclick="alert(1)">Safe text</p>\n\n'
        + '![track](https://example.com/tracker.png)\n\n'
        + '<a href="javascript:alert(1)">js</a> '
        + '<a href="data:text/html,unsafe">data</a> '
        + '<a href="//example.com">relative</a> '
        + '<a href="/settings">local</a> '
        + '[mail](mailto:reader@example.com) [fragment](#section)');
    expect(element.querySelector('script, style, form, input, svg, iframe, img, [onclick], [style], [id]')).toBeNull();
    expect([...element.querySelectorAll('a[href]')].map((a) => a.getAttribute('href')))
        .toEqual(['mailto:reader@example.com', '#section']);
    expect(element.textContent).toContain('Safe text');
    expect(element.querySelector('a[href="#section"]').hasAttribute('target')).toBe(false);
});

test('coalesces split Markdown tokens, flushes at completion, and preserves raw Markdown in history', async () => {
    const msg = createAssistantMessage();
    msg.appendText('## 总体');
    msg.appendText('评价\n\n**优');
    msg.appendText('点**');
    await new Promise((resolve) => requestAnimationFrame(resolve));
    expect(msg.el.querySelector('h2').textContent).toBe('总体评价');
    expect(msg.el.querySelector('strong').textContent).toBe('优点');
    msg.appendText('\n\n- 质量控制');
    msg.finalizeForHistory();
    expect(msg.el.querySelector('li').textContent).toBe('质量控制');
    const saved = getCurrentSession();
    expect(saved.messages[0].text).toBe('## 总体评价\n\n**优点**\n\n- 质量控制');
    expect(saved.messages[0].text).not.toContain('<h2>');
    setCurrentSession(saved);
    expect(document.querySelector('.msg-body h2').textContent).toBe('总体评价');
    expect(document.querySelector('.msg-body li').textContent).toBe('质量控制');
});

test('a final answer or error supersedes an outstanding streaming render', async () => {
    const msg = createAssistantMessage();
    msg.appendText('**unfinished');
    msg.setText('## Final');
    await new Promise((resolve) => requestAnimationFrame(resolve));
    expect(msg.el.querySelector('.msg-body h2').textContent).toBe('Final');
    msg.appendText('more output');
    msg.markError('<img src=x onerror=alert(1)> **failed**');
    await new Promise((resolve) => requestAnimationFrame(resolve));
    expect(msg.el.querySelector('.msg-body').textContent).toBe('Error: <img src=x onerror=alert(1)> **failed**');
    expect(msg.el.querySelector('.msg-body img, .msg-body strong')).toBeNull();
});

test('user input stays literal and answer-only turns have no duplicate task summary', () => {
    addUserMessage('## Heading <img src=x>');
    const user = document.querySelector('.chat-message-user');
    expect(user.textContent).toBe('## Heading <img src=x>');
    expect(user.querySelector('h2, img')).toBeNull();
    const msg = createAssistantMessage();
    msg.setText(answer);
    msg.finalizeForHistory();
    expect(msg.el.querySelector('.msg-completion').hidden).toBe(true);
});

function attachEdit(msg) {
    const card = createProposalCard({ title: 'Proposed edits to 5 sections', onApply: async () => {} });
    msg.attachProposal(card, { title: 'Proposed edits to 5 sections', state: 'pending',
        summary: 'Clarified terminology and tightened repetitive sentences.', countsText: '48431 → 43923 chars',
        items: Array.from({ length: 5 }, (_, i) => ({ id: i, label: `Section ${i + 1}`, before: 'old', after: 'new' })) });
    return card;
}

test('proposal tasks finish with a concise summary after the card that survives save/load', () => {
    const msg = createAssistantMessage();
    attachEdit(msg);
    expect(msg.el.querySelector('.msg-completion').hidden).toBe(true);
    msg.finalizeForHistory();
    const summary = msg.el.querySelector('.msg-completion');
    expect(summary.hidden).toBe(false);
    expect(summary.textContent).toContain('5 proposed changes');
    expect(summary.textContent).toContain('48431 → 43923 chars');
    expect(summary.textContent).toContain('Ready for review');
    expect(summary.textContent).toContain('Clarified terminology');
    expect(summary.previousElementSibling.className).toBe('msg-extras');
    expect(summary.querySelector('p').textContent).not.toContain('were applied');
    const text = summary.textContent;
    const session = getCurrentSession();
    session.messages = session.messages.map(normalizeMessage);
    setCurrentSession(session);
    expect(document.querySelector('.msg-completion').textContent).toBe(text);
});

test('model descriptions in the task summary also use sanitized Markdown', () => {
    const msg = createAssistantMessage();
    const card = createProposalCard({ title: 'Edits', onApply: async () => {} });
    msg.attachProposal(card, { state: 'pending', items: [],
        summary: 'Clarified **terminology**. <img src=x onerror=alert(1)>' });
    msg.finalizeForHistory();
    const summary = msg.el.querySelector('.msg-completion');
    expect(summary.querySelector('.msg-markdown strong').textContent).toBe('terminology');
    expect(summary.querySelector('img, [onerror]')).toBeNull();
});

test.each([
    ['markApplied', undefined, 'The selected changes were applied.'],
    ['markRejected', undefined, 'The proposals were rejected.'],
    ['markWarning', 'Application was partial.', 'requiring review'],
    ['markError', 'Write failed.', 'application error'],
    ['setPaused', 'Paused after applying one section.', 'paused during application'],
])('summary follows %s after the turn finishes and after history reload', (method, detail, expected) => {
    const msg = createAssistantMessage();
    const card = attachEdit(msg);
    msg.finalizeForHistory();
    card[method](detail);
    const summary = msg.el.querySelector('.msg-completion');
    expect(summary.textContent).toContain(expected);
    expect(summary.textContent).not.toContain('5 changes were applied');
    setCurrentSession(getCurrentSession());
    expect(document.querySelector('.msg-completion').textContent).toContain(expected);
});

test('cancelled and failed tasks keep their partial proposals without claiming completion', () => {
    const msg = createAssistantMessage();
    attachEdit(msg);
    msg.setStatus('Cancelled.');
    msg.finalizeForHistory();
    expect(msg.el.querySelector('.msg-completion').textContent).toContain('task was cancelled');
    msg.markError('A later step failed');
    expect(msg.el.querySelector('.msg-completion').textContent).toContain('task ended with an error');
    expect(getCurrentSession().messages[0].error).toBe('A later step failed');
});
