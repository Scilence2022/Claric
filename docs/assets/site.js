(() => {
  const root = document.documentElement;
  const toggle = document.getElementById('langToggle');
  const storageKey = 'claric.site.lang';
  const validLanguage = value => value === 'en' || value === 'zh';

  function initialLanguage() {
    const query = new URLSearchParams(window.location.search).get('lang');
    if (validLanguage(query)) return query;
    try {
      const saved = window.localStorage.getItem(storageKey);
      if (validLanguage(saved)) return saved;
    } catch {
      // Storage can be unavailable in a restricted browsing context.
    }
    return /^zh(?:-|$)/i.test(window.navigator.language || '') ? 'zh' : 'en';
  }

  function applyLanguage(language) {
    root.dataset.lang = language;
    root.lang = language === 'zh' ? 'zh-CN' : 'en';
    document.querySelectorAll('[data-lang]').forEach(element => {
      if (element === root) return;
      element.lang = element.dataset.lang === 'zh' ? 'zh-CN' : 'en';
      element.hidden = element.dataset.lang !== language;
    });
    document.querySelectorAll('[data-alt-en]').forEach(element => {
      element.alt = element.getAttribute(`data-alt-${language}`);
    });
    document.querySelectorAll('[data-aria-en]').forEach(element => {
      element.setAttribute('aria-label', element.getAttribute(`data-aria-${language}`));
    });
    if (toggle) {
      toggle.textContent = language === 'zh' ? 'EN' : '中文';
      toggle.lang = language === 'zh' ? 'en' : 'zh-CN';
      toggle.setAttribute('aria-label', language === 'zh' ? 'Switch to English' : '切换为中文');
      toggle.dataset.langTarget = language === 'zh' ? 'en' : 'zh';
    }
    document.title = language === 'zh'
      ? 'Claric | Microsoft Word 中先审阅、再应用的 AI 编辑'
      : 'Claric | Review-first AI editing for Microsoft Word';
  }

  applyLanguage(initialLanguage());
  if (toggle) {
    toggle.addEventListener('click', () => {
      const language = root.dataset.lang === 'zh' ? 'en' : 'zh';
      applyLanguage(language);
      try {
        window.localStorage.setItem(storageKey, language);
      } catch {
        // The visible language still changes when persistence is blocked.
      }
      try {
        const url = new URL(window.location.href);
        url.searchParams.set('lang', language);
        window.history.replaceState(null, '', url);
      } catch {
        // Switching language does not depend on history access.
      }
    });
    toggle.hidden = false;
  }

  const tablist = document.getElementById('install-tabs');
  if (!tablist) return;
  const tabs = Array.from(tablist.querySelectorAll('[data-platform]'));
  const panels = Array.from(document.querySelectorAll('[data-install-panel]'));
  if (tabs.length !== 3 || panels.length !== 3) return;

  function selectPlatform(platform, focus = false) {
    tabs.forEach(tab => {
      const selected = tab.dataset.platform === platform;
      tab.setAttribute('aria-selected', String(selected));
      tab.tabIndex = selected ? 0 : -1;
      if (selected && focus) tab.focus();
    });
    panels.forEach(panel => {
      panel.hidden = panel.dataset.installPanel !== platform;
    });
  }

  function platformFromHash() {
    const panel = panels.find(candidate => `#${candidate.id}` === window.location.hash);
    if (panel) selectPlatform(panel.dataset.installPanel);
  }

  tablist.setAttribute('role', 'tablist');
  tabs.forEach((tab, index) => {
    tab.setAttribute('role', 'tab');
    tab.addEventListener('click', () => selectPlatform(tab.dataset.platform));
    tab.addEventListener('keydown', event => {
      let next = index;
      if (event.key === 'ArrowRight') next = (index + 1) % tabs.length;
      else if (event.key === 'ArrowLeft') next = (index + tabs.length - 1) % tabs.length;
      else if (event.key === 'Home') next = 0;
      else if (event.key === 'End') next = tabs.length - 1;
      else return;
      event.preventDefault();
      selectPlatform(tabs[next].dataset.platform, true);
    });
  });
  panels.forEach(panel => {
    panel.setAttribute('role', 'tabpanel');
    panel.setAttribute('aria-labelledby', `tab-${panel.dataset.installPanel}`);
    panel.tabIndex = 0;
  });
  selectPlatform('macos');
  platformFromHash();
  window.addEventListener('hashchange', platformFromHash);
  panels[0].parentElement.dataset.enhanced = 'true';
  tablist.hidden = false;
})();
