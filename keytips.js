(() => {
  const CHARS = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
  const SELECTOR = 'a[href], button, input:not([type="hidden"]), select, textarea, summary, [role="button"], [role="link"], [tabindex]:not([tabindex="-1"])';

  let container = null;
  let items = [];
  let typed = '';
  let altPending = false;

  const boundWindows = new WeakSet();

  const collectElements = (doc = document, offsetX = 0, offsetY = 0) => {
    const list = [...doc.querySelectorAll(SELECTOR)].map((el) => {
      const r = el.getBoundingClientRect();
      return { el, x: r.left + offsetX, y: r.top + offsetY };
    });

    for (const frame of doc.querySelectorAll('iframe, frame')) {
      try {
        const fDoc = frame.contentDocument;
        if (!fDoc) continue;
        bindWindow(frame.contentWindow);
        const fr = frame.getBoundingClientRect();
        list.push(...collectElements(fDoc, offsetX + fr.left, offsetY + fr.top));
      } catch {
        // Cross-origin iframe: cannot access contentDocument from parent
      }
    }
    return list;
  };

  const makeCodes = (count) => {
    const len = count <= CHARS.length ? 1 : Math.ceil(Math.log(count) / Math.log(CHARS.length));
    const codes = [];
    for (let i = 0; i < count; i++) {
      let code = '', n = i;
      for (let d = 0; d < len; d++) {
        code = CHARS[n % CHARS.length] + code;
        n = Math.floor(n / CHARS.length);
      }
      codes.push(code);
    }
    return codes;
  };

  const hide = () => {
    container?.remove();
    container = null;
    items = [];
    typed = '';
  };

  const render = () => {
    items.forEach(({ code, badge }) => {
      const match = code.startsWith(typed);
      badge.style.display = match ? 'inline-block' : 'none';
      if (match) {
        const prefix = code.slice(0, typed.length);
        const rest = code.slice(typed.length);
        badge.innerHTML = `<span style="opacity:.45">${prefix}</span>${rest}`;
      }
    });
  };

  const show = () => {
    hide();
    const targets = collectElements();
    if (!targets.length) return;

    const codes = makeCodes(targets.length);
    container = document.createElement('div');
    container.id = '__alt_keytips__';
    Object.assign(container.style, {
      position: 'fixed', inset: '0', pointerEvents: 'none', zIndex: '2147483647'
    });

    items = targets.map(({ el, x, y }, i) => {
      const badge = document.createElement('span');
      Object.assign(badge.style, {
        position: 'fixed',
        left: `${Math.max(2, Math.min(x, innerWidth - 32))}px`,
        top: `${Math.max(2, Math.min(y - 6, innerHeight - 22))}px`,
        background: '#1f1f1f',
        color: '#fff',
        border: '1px solid #888',
        borderRadius: '3px',
        padding: '1px 5px',
        font: '700 11px/1.3 ui-monospace, Consolas, monospace',
        letterSpacing: '0.5px',
        boxShadow: '0 2px 5px rgba(0,0,0,.35)'
      });
      container.appendChild(badge);
      return { el, code: codes[i], badge };
    });

    document.documentElement.appendChild(container);
    render();
  };

  const onKeyDown = (e) => {
    if (e.key === 'Alt' && !e.repeat) {
      altPending = true;
      return;
    }
    altPending = false;

    if (!container) return;
    if (e.ctrlKey || e.metaKey || e.altKey) return;

    e.preventDefault();
    e.stopPropagation();

    if (e.key === 'Escape') return hide();
    if (e.key === 'Backspace') {
      typed = typed.slice(0, -1);
      return render();
    }

    const ch = e.key.toUpperCase();
    if (ch.length !== 1 || !CHARS.includes(ch)) return;

    const next = typed + ch;
    const matches = items.filter((it) => it.code.startsWith(next));
    if (!matches.length) return;

    typed = next;
    if (matches.length === 1 && matches[0].code === typed) {
      const target = matches[0].el;
      hide();
      target.focus({ preventScroll: true });
      target.click();
    } else {
      render();
    }
  };

  const onKeyUp = (e) => {
    if (e.key === 'Alt' && altPending) {
      altPending = false;
      e.preventDefault();
      container ? hide() : show();
    }
  };

  const onBlur = () => { altPending = false; hide(); };
  const onScroll = () => { if (container) hide(); };

  function bindWindow(win) {
    if (!win || boundWindows.has(win)) return;
    boundWindows.add(win);
    win.addEventListener('keydown', onKeyDown, true);
    win.addEventListener('keyup', onKeyUp, true);
    win.addEventListener('blur', onBlur);
    win.addEventListener('scroll', onScroll, true);
  }

  bindWindow(window);
})();
