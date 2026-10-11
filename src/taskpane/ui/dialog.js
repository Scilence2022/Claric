const focusScopes = [];

export function containFocus(root, onEscape, initialFocus) {
    const previous = document.activeElement;
    const scope = { root };
    focusScopes.push(scope);
    root.tabIndex = -1;
    const visible = (el) => {
        if (!el.isConnected || el.closest('[hidden], [inert]')) return false;
        for (let ancestor = el; ancestor; ancestor = ancestor.parentElement) {
            const style = getComputedStyle(ancestor);
            if (style.display === 'none' || style.visibility === 'hidden' || style.visibility === 'collapse') return false;
        }
        return true;
    };
    const focusable = () => [...root.querySelectorAll('button, input, select, textarea, a[href], [tabindex]')]
        .filter((el) => !el.matches(':disabled') && el.tabIndex >= 0 && visible(el));
    const entry = () => {
        const items = focusable();
        const target = items.includes(initialFocus) ? initialFocus : (items[0] || root);
        if (visible(target)) target.focus();
    };
    const keydown = (event) => {
        if (focusScopes[focusScopes.length - 1] !== scope) return;
        if (event.key === 'Escape') {
            event.preventDefault();
            event.stopImmediatePropagation();
            onEscape();
        } else if (event.key === 'Tab') {
            const items = focusable();
            const index = items.indexOf(document.activeElement);
            if (!items.length || index < 0 || (event.shiftKey ? index === 0 : index === items.length - 1)) {
                event.preventDefault();
                (items[event.shiftKey ? items.length - 1 : 0] || root).focus();
            }
        }
    };
    const focusin = (event) => {
        if (focusScopes[focusScopes.length - 1] === scope && !root.contains(event.target)) entry();
    };
    root.addEventListener('keydown', keydown);
    document.addEventListener('focusin', focusin);
    entry();
    return () => {
        const index = focusScopes.indexOf(scope);
        if (index >= 0) focusScopes.splice(index, 1);
        root.removeEventListener('keydown', keydown);
        document.removeEventListener('focusin', focusin);
        if (previous && !previous.matches(':disabled') && visible(previous)) previous.focus();
    };
}

export function confirmAutoApply() {
    return new Promise((resolve) => {
        const overlay = document.createElement('div');
        overlay.className = 'confirm-overlay';
        const dialog = document.createElement('section');
        dialog.className = 'confirm-dialog';
        dialog.setAttribute('role', 'alertdialog');
        dialog.setAttribute('aria-modal', 'true');
        dialog.setAttribute('aria-labelledby', 'autoApplyConfirmTitle');
        dialog.setAttribute('aria-describedby', 'autoApplyConfirmDescription');
        dialog.innerHTML = '<h2 id="autoApplyConfirmTitle">Enable Auto-apply?</h2>'
            + '<p id="autoApplyConfirmDescription">Proposed changes will write to your document automatically after successful turns. Track Changes must stay enabled, but some structural changes may not be tracked. Review the document after applying.</p>'
            + '<div class="modal-actions"><button type="button" class="btn btn-secondary" data-confirm="cancel">Cancel</button>'
            + '<button type="button" class="btn btn-primary" data-confirm="enable">Enable Auto-apply</button></div>';
        overlay.appendChild(dialog);
        document.body.appendChild(overlay);
        let release;
        const finish = (accepted) => {
            overlay.remove();
            release?.();
            resolve(accepted === true);
        };
        const cancel = dialog.querySelector('[data-confirm="cancel"]');
        cancel.addEventListener('click', () => finish(false));
        dialog.querySelector('[data-confirm="enable"]').addEventListener('click', () => finish(true));
        release = containFocus(dialog, () => finish(false), cancel);
    });
}

/** Longest argument preview shown in the MCP approval dialog. */
export const MCP_ARGS_PREVIEW_CHARS = 1200;

/**
 * Asks the user to approve one MCP tool call that may have side effects
 * outside the document. Server, tool and arguments come from the model and
 * the server, so they are rendered as text, never markup.
 *
 * @param {{serverName: string, tool: string, args: object}} call
 * @param {{signal?: AbortSignal}} [options] - Aborting closes the dialog and
 *   rejects with AbortError (the turn was cancelled)
 * @returns {Promise<'once'|'run'|'deny'>} Matches MCP_APPROVAL in lib/mcp-tools.js
 */
export function confirmMcpToolCall({ serverName, tool, args }, { signal } = {}) {
    return new Promise((resolve, reject) => {
        if (signal?.aborted) {
            reject(new DOMException('MCP approval cancelled.', 'AbortError'));
            return;
        }
        const overlay = document.createElement('div');
        overlay.className = 'confirm-overlay';
        const dialog = document.createElement('section');
        dialog.className = 'confirm-dialog';
        dialog.setAttribute('role', 'alertdialog');
        dialog.setAttribute('aria-modal', 'true');
        dialog.setAttribute('aria-labelledby', 'mcpConfirmTitle');
        dialog.setAttribute('aria-describedby', 'mcpConfirmDescription');

        const title = document.createElement('h2');
        title.id = 'mcpConfirmTitle';
        title.textContent = 'Allow this tool call?';
        const description = document.createElement('p');
        description.id = 'mcpConfirmDescription';
        description.append('The model wants to run ');
        const toolName = document.createElement('strong');
        toolName.textContent = tool;
        description.append(toolName, ' on MCP server ');
        const server = document.createElement('strong');
        server.textContent = serverName || 'unknown';
        description.append(server, '. The server does not mark this tool read-only, so it may change data outside '
            + 'Word. Instructions inside your document can influence what the model asks for.');
        const preview = document.createElement('pre');
        preview.className = 'confirm-args';
        let json;
        try { json = JSON.stringify(args ?? {}, null, 2); } catch { json = String(args); }
        preview.textContent = json.length > MCP_ARGS_PREVIEW_CHARS
            ? `${json.slice(0, MCP_ARGS_PREVIEW_CHARS)}\n… (${json.length - MCP_ARGS_PREVIEW_CHARS} more characters)`
            : json;

        const actions = document.createElement('div');
        actions.className = 'modal-actions';
        const button = (label, value, primary = false) => {
            const el = document.createElement('button');
            el.type = 'button';
            el.className = `btn ${primary ? 'btn-primary' : 'btn-secondary'}`;
            el.dataset.confirm = value;
            el.textContent = label;
            el.addEventListener('click', () => finish(value));
            return el;
        };
        const deny = button('Deny', 'deny');
        actions.append(deny, button('Allow for this run', 'run'), button('Allow once', 'once', true));
        dialog.append(title, description, preview, actions);
        overlay.appendChild(dialog);
        document.body.appendChild(overlay);

        let release;
        let settled = false;
        const close = () => {
            settled = true;
            overlay.remove();
            release?.();
            signal?.removeEventListener('abort', onAbort);
        };
        function finish(value) {
            if (settled) return;
            close();
            resolve(value);
        }
        function onAbort() {
            if (settled) return;
            close();
            reject(new DOMException('MCP approval cancelled.', 'AbortError'));
        }
        signal?.addEventListener('abort', onAbort, { once: true });
        release = containFocus(dialog, () => finish('deny'), deny);
    });
}
