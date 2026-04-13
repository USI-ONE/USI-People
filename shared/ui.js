/* USI People — Shared UI Components */

(() => {
  // ── View Switching ──
  USI.show = function(viewId) {
    document.querySelectorAll('.view').forEach(v => v.classList.toggle('active', v.id === viewId));
  };

  // ── Toast Notifications ──
  let _toastContainer;
  USI.toast = function(msg, type = '') {
    if (!_toastContainer) {
      _toastContainer = document.createElement('div');
      _toastContainer.className = 'toast-container';
      document.body.appendChild(_toastContainer);
    }
    const t = document.createElement('div');
    t.className = 'toast' + (type ? ` toast-${type}` : '');
    t.textContent = msg;
    _toastContainer.appendChild(t);
    requestAnimationFrame(() => t.classList.add('show'));
    setTimeout(() => {
      t.classList.remove('show');
      setTimeout(() => t.remove(), 300);
    }, 3500);
  };

  // ── Tab System ──
  USI.initTabs = function(containerEl) {
    const tabs = containerEl.querySelectorAll('.tab');
    const panels = containerEl.querySelectorAll('.tab-panel');
    tabs.forEach(tab => {
      tab.addEventListener('click', () => {
        const target = tab.dataset.tab;
        tabs.forEach(t => t.classList.toggle('active', t === tab));
        panels.forEach(p => p.classList.toggle('active', p.id === target));
        if (typeof tab._onChange === 'function') tab._onChange(target);
        containerEl.dispatchEvent(new CustomEvent('tabchange', { detail: target }));
      });
    });
  };

  // ── Modal ──
  USI.modal = function(title, html, buttons = []) {
    return new Promise(resolve => {
      const overlay = document.createElement('div');
      overlay.className = 'modal-overlay';
      overlay.innerHTML = `
        <div class="modal">
          <div class="modal-header">
            <div class="modal-title">${USI.escapeHtml(title)}</div>
            <button class="modal-close">&times;</button>
          </div>
          <div class="modal-body">${html}</div>
          ${buttons.length ? `<div class="modal-footer">
            ${buttons.map((b, i) => `<button class="btn ${b.class || 'btn-secondary'}" data-idx="${i}">${USI.escapeHtml(b.label)}</button>`).join('')}
          </div>` : ''}
        </div>`;
      document.body.appendChild(overlay);
      requestAnimationFrame(() => overlay.classList.add('open'));

      function close(result) {
        overlay.classList.remove('open');
        setTimeout(() => overlay.remove(), 200);
        resolve(result);
      }

      overlay.querySelector('.modal-close').onclick = () => close(null);
      overlay.addEventListener('click', e => { if (e.target === overlay) close(null); });
      overlay.querySelectorAll('.modal-footer .btn').forEach(btn => {
        btn.addEventListener('click', async () => {
          const idx = parseInt(btn.dataset.idx);
          const b = buttons[idx];
          if (b.onClick) {
            // Disable all buttons and show saving state
            const allBtns = overlay.querySelectorAll('.modal-footer .btn');
            const originalText = btn.textContent;
            allBtns.forEach(b => b.disabled = true);
            btn.textContent = 'Saving...';
            try {
              const result = await b.onClick(overlay);
              if (result !== false) close(b.value !== undefined ? b.value : b.label);
              else {
                // Validation failed — restore buttons
                allBtns.forEach(b => b.disabled = false);
                btn.textContent = originalText;
              }
            } catch (e) {
              allBtns.forEach(b => b.disabled = false);
              btn.textContent = originalText;
              throw e;
            }
          } else {
            close(b.value !== undefined ? b.value : b.label);
          }
        });
      });

      return overlay;
    });
  };

  USI.confirm = function(message) {
    return USI.modal('Confirm', `<p style="font-size:14px">${USI.escapeHtml(message)}</p>`, [
      { label: 'Cancel', class: 'btn-secondary', value: false },
      { label: 'Confirm', class: 'btn-primary', value: true }
    ]);
  };

  // ── Navigation Bar ──
  USI.renderNav = function(activePage) {
    const existing = document.querySelector('.topbar');
    if (existing) existing.remove();
    const nav = document.createElement('nav');
    nav.className = 'topbar';
    const pages = [
      { id: 'performance', label: 'Performance', href: 'performance.html' },
      { id: 'oneonone', label: '1:1s', href: 'oneonone.html' },
      { id: 'onboarding', label: 'HR / Onboarding', href: 'onboarding.html' },
      { id: 'guide', label: 'User Guide', href: 'guide.html' }
    ];
    const me = JSON.parse(sessionStorage.getItem('usi_me') || '{}');
    nav.innerHTML = `
      <button class="topbar-hamburger" onclick="this.nextElementSibling.classList.toggle('open')">
        <svg viewBox="0 0 20 20" fill="currentColor"><path fill-rule="evenodd" d="M3 5a1 1 0 011-1h12a1 1 0 110 2H4a1 1 0 01-1-1zm0 5a1 1 0 011-1h12a1 1 0 110 2H4a1 1 0 01-1-1zm0 5a1 1 0 011-1h12a1 1 0 110 2H4a1 1 0 01-1-1z" clip-rule="evenodd"/></svg>
      </button>
      <div class="topbar-links">
        ${pages.map(p => `<a href="${p.href}" class="topbar-link ${p.id === activePage ? 'active' : ''}">${p.label}</a>`).join('')}
      </div>
      <div class="topbar-brand">USI People</div>
      <div class="topbar-spacer"></div>
      <div class="topbar-user">
        <span class="topbar-user-name">${USI.escapeHtml(me.displayName || '')}</span>
        <button class="btn btn-ghost btn-sm" onclick="USI.logout()">Sign Out</button>
      </div>`;
    document.body.prepend(nav);
  };

  // ── Star Rating Component ──
  USI.renderStars = function(container, value, editable, onChange) {
    const el = typeof container === 'string' ? document.getElementById(container) : container;
    if (!el) return;
    el.className = 'star-rating' + (editable ? ' editable' : '');
    el.innerHTML = '';
    for (let i = 1; i <= 5; i++) {
      const star = document.createElement('span');
      star.className = 'star' + (i <= value ? ' filled' : '');
      star.textContent = '\u2605';
      star.dataset.value = i;
      if (editable) {
        star.addEventListener('mouseenter', () => {
          el.querySelectorAll('.star').forEach(s => {
            s.classList.toggle('hover', parseInt(s.dataset.value) <= i);
          });
        });
        star.addEventListener('mouseleave', () => {
          el.querySelectorAll('.star').forEach(s => s.classList.remove('hover'));
        });
        star.addEventListener('click', () => {
          el.querySelectorAll('.star').forEach(s => {
            s.classList.toggle('filled', parseInt(s.dataset.value) <= i);
          });
          if (onChange) onChange(i);
        });
      }
      el.appendChild(star);
    }
    return el;
  };

  USI.getStarValue = function(container) {
    const el = typeof container === 'string' ? document.getElementById(container) : container;
    if (!el) return 0;
    return el.querySelectorAll('.star.filled').length;
  };

  // ── Utility Functions ──
  USI.escapeHtml = function(str) {
    if (!str) return '';
    const d = document.createElement('div');
    d.textContent = str;
    return d.innerHTML;
  };

  USI.formatDate = function(iso) {
    if (!iso) return '';
    const d = new Date(iso);
    return d.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });
  };

  USI.formatDateTime = function(iso) {
    if (!iso) return '';
    const d = new Date(iso);
    return d.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric', hour: 'numeric', minute: '2-digit' });
  };

  USI.relativeTime = function(iso) {
    if (!iso) return '';
    const diff = Date.now() - new Date(iso).getTime();
    const mins = Math.floor(diff / 60000);
    if (mins < 1) return 'just now';
    if (mins < 60) return `${mins}m ago`;
    const hrs = Math.floor(mins / 60);
    if (hrs < 24) return `${hrs}h ago`;
    const days = Math.floor(hrs / 24);
    if (days < 30) return `${days}d ago`;
    return USI.formatDate(iso);
  };

  USI.debounce = function(fn, ms) {
    let t;
    return function(...args) { clearTimeout(t); t = setTimeout(() => fn.apply(this, args), ms); };
  };

  USI.generateId = function() {
    return Date.now().toString(36) + Math.random().toString(36).substr(2, 5);
  };

  // ── Overall Rating Label ──
  USI.ratingLabel = function(rating) {
    const labels = ['', 'Needs Improvement', 'Developing', 'Meeting Expectations', 'Exceeding Expectations', 'Outstanding'];
    return labels[Math.round(rating)] || '';
  };

  // ── Pulse Class ──
  USI.pulseClass = function(pulse) {
    const map = {
      'OnTrack': 'on-track', 'On Track': 'on-track',
      'NeedsAttention': 'needs-attention', 'Needs Attention': 'needs-attention', 'At Risk': 'at-risk',
      'Blocked': 'blocked', 'Off Track': 'off-track',
      'Not Started': 'not-started'
    };
    return map[pulse] || 'not-started';
  };

  // ── Status Pill Class ──
  USI.statusPillClass = function(status) {
    const map = {
      'Not Started': 'pill-neutral', 'Pending': 'pill-neutral',
      'In Progress': 'pill-info', 'Active': 'pill-info', 'Draft': 'pill-neutral',
      'On Track': 'pill-success', 'Complete': 'pill-success', 'Completed': 'pill-success',
      'Submitted': 'pill-info',
      'At Risk': 'pill-warning', 'Behind': 'pill-warning', 'Blocked': 'pill-error',
      'Cancelled': 'pill-neutral'
    };
    return map[status] || 'pill-neutral';
  };

  // ── Boot Helper ──
  USI.boot = async function(pageName, initFn) {
    // Check for auth redirect
    if (window.location.search.includes('code=')) {
      USI.show('loadingView');
      const ok = await USI.handleRedirect();
      if (!ok) { USI.show('loginView'); return; }
    }

    // Check for valid token
    if (!USI.getToken()) {
      USI.show('loginView');
      return;
    }

    // Authenticated — load app
    USI.show('loadingView');
    try {
      const me = await USI.getMe();
      USI.renderNav(pageName);
      await initFn(me);
      USI.show('appView');
    } catch (e) {
      console.error('Boot failed:', e);
      USI.toast('Failed to load: ' + e.message, 'error');
      USI.show('appView');
    }
  };

  // ── Microsoft Sign-In Button SVG ──
  USI.microsoftLogo = `<svg viewBox="0 0 21 21"><rect x="1" y="1" width="9" height="9" fill="#f25022"/><rect x="11" y="1" width="9" height="9" fill="#7fba00"/><rect x="1" y="11" width="9" height="9" fill="#00a4ef"/><rect x="11" y="11" width="9" height="9" fill="#ffb900"/></svg>`;
})();
