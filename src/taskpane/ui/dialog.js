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
