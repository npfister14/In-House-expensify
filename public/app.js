(() => {

  // Toast notifications
  const containerId = 'toast-container';
  const ensureContainer = () => {
    let c = document.getElementById(containerId);
    if (!c) {
      c = document.createElement('div');
      c.id = containerId;
      document.body.appendChild(c);
    }
    return c;
  };
  window.showToast = (msg, type = 'info') => {
    const c = ensureContainer();
    const t = document.createElement('div');
    t.className = `toast ${type}`;
    t.textContent = msg;
    c.appendChild(t);
    requestAnimationFrame(() => t.classList.add('show'));
    setTimeout(() => {
      t.classList.remove('show');
      setTimeout(() => t.remove(), 200);
    }, 3000);
  };

  // Modal (image preview)
  window.openModalImage = (src) => {
    const overlay = document.createElement('div');
    overlay.className = 'modal-overlay';
    const modal = document.createElement('div');
    modal.className = 'modal';
    const img = document.createElement('img');
    img.src = src;
    img.alt = 'Receipt';
    modal.appendChild(img);
    overlay.appendChild(modal);
    document.body.appendChild(overlay);
    overlay.addEventListener('click', () => overlay.remove());
    requestAnimationFrame(() => overlay.classList.add('open'));
  };

  // Mobile menu / sidebar toggle
  const menuToggle = document.querySelector('.menu-toggle');
  const siteNav = document.querySelector('.nav');
  const sidebar = document.querySelector('.sidebar');
  if (menuToggle) {
    menuToggle.addEventListener('click', () => {
      if (siteNav) siteNav.classList.toggle('open');
      if (sidebar) {
        sidebar.classList.toggle('open');
        document.body.classList.toggle('sidebar-open');
      }
    });
  }

  // Theme toggle (persisted)
  const themeKey = 'ihx-theme';
  const applyTheme = (t) => {
    const root = document.documentElement;
    if (t === 'light') root.setAttribute('data-theme', 'light');
    else root.removeAttribute('data-theme');
  };
  try { applyTheme(localStorage.getItem(themeKey)); } catch {}
  const themeBtn = document.getElementById('theme-toggle');
  if (themeBtn) {
    themeBtn.addEventListener('click', () => {
      const isLight = document.documentElement.getAttribute('data-theme') === 'light';
      const next = isLight ? 'dark' : 'light';
      applyTheme(next);
      try { localStorage.setItem(themeKey, next); } catch {}
      themeBtn.textContent = next === 'light' ? '☀️' : '🌙';
    });
  }

  // Dropzone helper (index page)
  const dz = document.getElementById('dropzone');
  const fileInput = document.getElementById('image');
  const preview = document.getElementById('preview');
  if (dz && fileInput) {
    const showPreview = (file) => {
      if (!file) return;
      if (preview) {
        try {
          preview.src = URL.createObjectURL(file);
        } catch {
          const r = new FileReader();
          r.onload = () => { preview.src = r.result; };
          r.readAsDataURL(file);
        }
        preview.hidden = false;
        const noPrev = document.getElementById('no-preview');
        if (noPrev) noPrev.hidden = true;
      }
      if (window.showToast) showToast('Receipt selected', 'success');
    };
    dz.addEventListener('dragover', (e) => { e.preventDefault(); dz.classList.add('hover'); });
    dz.addEventListener('dragleave', () => dz.classList.remove('hover'));
    dz.addEventListener('drop', (e) => {
      e.preventDefault(); dz.classList.remove('hover');
      const f = (e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0]) || null;
      if (f) {
        // Cross‑browser: try direct assignment first; fallback to DataTransfer if available; ignore if read‑only
        let assigned = false;
        try {
          fileInput.files = e.dataTransfer.files;
          assigned = true;
        } catch {}
        if (!assigned) {
          try {
            const dt = new DataTransfer();
            dt.items.add(f);
            fileInput.files = dt.files;
            assigned = true;
          } catch {}
        }
        showPreview(f);
      }
    });
    // File input change
    fileInput.addEventListener('change', () => {
      const f = (fileInput.files && fileInput.files[0]) || null;
      showPreview(f);
    });
  }

  // Busy helpers
  window.withBusy = async (el, fn) => {
    const old = el.textContent;
    el.disabled = true;
    el.classList.add('busy');
    try { return await fn(); }
    finally { el.disabled = false; el.classList.remove('busy'); el.textContent = old; }
  };

  // Whoami badge
  const userPill = document.getElementById('user-pill');
  if (userPill) {
    fetch('/api/whoami').then(r => r.ok ? r.json() : null).then(d => {
      if (!d) return;
      const email = (d.email || '').trim();
      if (email) userPill.textContent = email;
    }).catch(() => {});
  }

  // End
})();
