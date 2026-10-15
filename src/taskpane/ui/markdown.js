import { Marked } from 'marked';
import { getPurifier } from '../../lib/sanitize.js';

// Keep chat options independent of the Word document renderer's defaults.
const markdown = new Marked({ gfm: true, breaks: true, async: false });

/** Render model output as formatted text without allowing active HTML. */
export function renderMarkdown(element, text) {
    element.innerHTML = getPurifier().sanitize(markdown.parse(text || ''), {
        ALLOWED_TAGS: ['p', 'br', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6',
            'strong', 'em', 'del', 's', 'ul', 'ol', 'li', 'blockquote',
            'pre', 'code', 'hr', 'a', 'table', 'thead', 'tbody', 'tr', 'th', 'td'],
        ALLOWED_ATTR: ['href', 'title', 'start', 'colspan', 'rowspan', 'align'],
        ALLOW_DATA_ATTR: false,
        ALLOW_ARIA_ATTR: false,
        ALLOWED_URI_REGEXP: /^(?:(?:https?|mailto):|#)/i,
    });
    element.classList.add('msg-markdown');
    for (const link of element.querySelectorAll('a[href]')) {
        if (link.getAttribute('href').startsWith('#')) continue;
        link.setAttribute('target', '_blank');
        link.setAttribute('rel', 'noopener noreferrer');
    }
    for (const table of element.querySelectorAll('table')) {
        const scroll = document.createElement('div');
        scroll.className = 'msg-markdown-table';
        scroll.tabIndex = 0;
        scroll.setAttribute('role', 'region');
        scroll.setAttribute('aria-label', 'Response table');
        table.replaceWith(scroll);
        scroll.appendChild(table);
    }
}
