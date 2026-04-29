// Form Studio — Shared UI (Toast, Modal, Progress, Login)

// =============================================================
// TOAST HELPER
// =============================================================
function showToast(type, message) {
  const container = document.getElementById("toast-container");
  const icons = {
    success: `<svg class="toast-icon" viewBox="0 0 16 16" fill="currentColor" aria-hidden="true"><path d="M8 0a8 8 0 100 16A8 8 0 008 0zm3.78 5.78l-4.5 4.5a.75.75 0 01-1.06 0l-2-2a.75.75 0 111.06-1.06l1.47 1.47 3.97-3.97a.75.75 0 111.06 1.06z"/></svg>`,
    error:   `<svg class="toast-icon" viewBox="0 0 16 16" fill="currentColor" aria-hidden="true"><path d="M8 0a8 8 0 100 16A8 8 0 008 0zm-.75 4a.75.75 0 011.5 0v4a.75.75 0 01-1.5 0V4zm.75 7.5a1 1 0 110-2 1 1 0 010 2z"/></svg>`,
    info:    `<svg class="toast-icon" viewBox="0 0 16 16" fill="currentColor" aria-hidden="true"><path d="M8 0a8 8 0 100 16A8 8 0 008 0zm.75 4.5a.75.75 0 00-1.5 0v4a.75.75 0 001.5 0v-4zm-.75 7.5a1 1 0 110-2 1 1 0 010 2z"/></svg>`,
  };
  const toast = document.createElement("div");
  toast.className = `toast toast-${type}`;
  toast.setAttribute("role", type === "error" ? "alert" : "status");
  toast.innerHTML = `${icons[type] || icons.info}<span>${message}</span>`;
  container.appendChild(toast);
  setTimeout(() => {
    toast.classList.add("removing");
    toast.addEventListener("animationend", () => toast.remove());
  }, 4000);
}

// =============================================================
// MODAL HELPER
// =============================================================
let _modalTriggerEl = null;

function openModal(html, locked = false, wide = false) {
  window._modalLocked = locked;
  _modalTriggerEl = document.activeElement;

  document.getElementById("modal-content").innerHTML = html;
  const modal   = document.getElementById("modal");
  const overlay = document.getElementById("overlay");
  modal.classList.toggle("modal-wide", wide);

  overlay.removeAttribute("hidden");
  overlay.classList.add("open");

  const titleEl = modal.querySelector(".modal-title");
  if (titleEl) {
    if (!titleEl.id) titleEl.id = "modal-title-sr";
    overlay.setAttribute("aria-labelledby", "modal-title-sr");
  }

  requestAnimationFrame(() => {
    const focusable = modal.querySelector(
      'button:not([disabled]), [href], input:not([disabled]), select:not([disabled]), textarea:not([disabled]), [tabindex]:not([tabindex="-1"])'
    );
    if (focusable) focusable.focus();
  });

  overlay.addEventListener("keydown", _trapFocus);
}

function _trapFocus(e) {
  if (e.key !== "Tab") return;
  const modal = document.getElementById("modal");
  const focusable = Array.from(modal.querySelectorAll(
    'button:not([disabled]), [href], input:not([disabled]), select:not([disabled]), textarea:not([disabled]), [tabindex]:not([tabindex="-1"])'
  ));
  if (!focusable.length) return;
  const first = focusable[0];
  const last  = focusable[focusable.length - 1];
  if (e.shiftKey && document.activeElement === first) {
    e.preventDefault(); last.focus();
  } else if (!e.shiftKey && document.activeElement === last) {
    e.preventDefault(); first.focus();
  }
}

function closeModal() {
  window._modalLocked = false;
  const overlay = document.getElementById("overlay");
  overlay.classList.remove("open");
  overlay.setAttribute("hidden", "");
  overlay.removeEventListener("keydown", _trapFocus);
  if (_modalTriggerEl && typeof _modalTriggerEl.focus === "function") {
    _modalTriggerEl.focus();
  }
  _modalTriggerEl = null;
}

document.addEventListener("keydown", e => {
  if (e.key === "Escape" && !window._modalLocked) closeModal();
});
document.addEventListener("click", e => {
  if (e.target.id === "overlay" && !window._modalLocked) closeModal();
});

// ── Progress modal ──────────────────────────────────────────
function showProgress(title, message) {
  window._modalLocked = true;
  _modalTriggerEl = document.activeElement;
  document.getElementById("modal-content").innerHTML = `
    <div class="modal-header" style="border-bottom:none;padding-bottom:8px;">
      <span class="modal-title" id="modal-title-sr">${escHtml(title)}</span>
    </div>
    <div class="modal-body" style="align-items:center;text-align:center;padding:32px 24px;gap:20px;">
      <div class="spinner" role="status" aria-label="${escHtml(title)}" style="width:36px;height:36px;border-width:3px;"></div>
      <div>
        <div id="progress-message" aria-live="polite" style="font-size:14px;color:var(--text2);">${escHtml(message)}</div>
      </div>
    </div>
  `;
  const overlay = document.getElementById("overlay");
  overlay.removeAttribute("hidden");
  overlay.classList.add("open");
}

function updateProgress(message) {
  const el = document.getElementById("progress-message");
  if (el) el.textContent = message;
}

function hideProgress() {
  window._modalLocked = false;
  const overlay = document.getElementById("overlay");
  overlay.classList.remove("open");
  overlay.setAttribute("hidden", "");
  if (_modalTriggerEl && typeof _modalTriggerEl.focus === "function") {
    _modalTriggerEl.focus();
    _modalTriggerEl = null;
  }
}

// =============================================================
// RENDER ENGINE
// =============================================================
function render() {
  const accounts = msalInstance.getAllAccounts();
  if (!accounts.length) { renderLoginScreen(); return; }
  renderAppShell();
}

function renderLoginScreen() {
  document.getElementById("app").innerHTML = `
    <div id="login-screen">
      <div class="login-card">
        <div class="login-logo">
          <img src="https://webmail.le.ac.uk/images/logo2020.jpg" alt="University of Leicester" style="height:64px;width:auto;">
        </div>
        <h1>Form Studio</h1>
        <p>Sign in with your University of Leicester Microsoft 365 account to continue.</p>
        <button class="btn btn-primary" style="width:100%;justify-content:center;padding:12px 20px;" onclick="login()" id="login-btn">
          <svg width="16" height="16" viewBox="0 0 23 23" fill="none" aria-hidden="true">
            <path fill="#f35325" d="M1 1h10v10H1z"/>
            <path fill="#81bc06" d="M12 1h10v10H12z"/>
            <path fill="#05a6f0" d="M1 12h10v10H1z"/>
            <path fill="#ffba08" d="M12 12h10v10H12z"/>
          </svg>
          Sign in with Microsoft
        </button>
      </div>
      <p style="font-size:12px;color:var(--text3);margin-top:16px;">University of Leicester — Form Studio v1.0</p>
    </div>
  `;
}

function showLoginLoading(show) {
  const btn = document.getElementById("login-btn");
  if (!btn) return;
  btn.disabled = show;
  btn.setAttribute("aria-busy", show ? "true" : "false");
  btn.innerHTML = show
    ? `<span class="spinner" role="status" aria-label="Signing in"></span> Signing in…`
    : `<svg width="16" height="16" viewBox="0 0 23 23" fill="none" aria-hidden="true"><path fill="#f35325" d="M1 1h10v10H1z"/><path fill="#81bc06" d="M12 1h10v10H12z"/><path fill="#05a6f0" d="M1 12h10v10H1z"/><path fill="#ffba08" d="M12 12h10v10H12z"/></svg> Sign in with Microsoft`;
}

function renderAppShell() {
  const { currentUser, isAdmin, hasFormRequestAccess, currentView } = AppState;
  const initials = currentUser ? currentUser.displayName.split(" ").map(n => n[0]).join("").slice(0,2) : "?";
  const displayName = escHtml(currentUser?.displayName || "User");

  document.getElementById("app").innerHTML = `
    <div class="app-shell">
      <!-- TOP BAR -->
      <header class="topbar" role="banner">
        <a class="topbar-logo" href="#" onclick="event.preventDefault();navigateTo('home')" aria-label="Form Studio — University of Leicester home">
          <img src="https://webmail.le.ac.uk/images/logo2020.jpg" alt="University of Leicester" style="height:32px;width:auto;border-radius:3px;">
          <span style="color:rgba(255,255,255,0.4);font-weight:300;margin:0 4px;" aria-hidden="true">|</span>
          <span style="color:rgba(255,255,255,0.9);font-weight:500;letter-spacing:0;">Form Studio</span>
        </a>
        <div class="topbar-divider" aria-hidden="true"></div>
        <span class="topbar-title" aria-hidden="true">University of Leicester</span>
        <div class="topbar-right">
          <div class="theme-switcher">
            <label for="theme-select" class="theme-label">Theme</label>
            <select class="theme-select" onchange="setTheme(this.value)" id="theme-select">
              <option value="uol">🎓 UoL</option>
              <option value="google">🔵 Google</option>
              <option value="apple">🍎 Apple</option>
              <option value="dark">🌙 Dark</option>
              <option value="glass">✨ Glass</option>
              <option value="pastel">🌸 Pastel</option>
              <option value="candy">🍬 Candy</option>
              <option value="ocean">🌊 Ocean</option>
              <option value="slate">🪨 Slate</option>
	      <option value="premium">✦ Premium</option>
              <option value="hc">◑ High Contrast</option>	
            </select>
          </div>
          <button class="user-chip" onclick="showUserMenu()" aria-label="Account menu for ${displayName}" aria-haspopup="dialog">
            <div class="avatar" aria-hidden="true">${initials}</div>
            <span aria-hidden="true">${displayName}</span>
            <svg width="12" height="12" viewBox="0 0 12 12" fill="currentColor" aria-hidden="true"><path d="M6 8L1 3h10L6 8z"/></svg>
          </button>
        </div>
      </header>

      <!-- SIDEBAR -->
      <nav class="sidebar" aria-label="Main navigation">
        ${navItem("home", homeIcon(), "Home")}
        ${navItem("my-forms", myFormsIcon(), "My Forms")}
        ${(isAdmin || hasFormRequestAccess) ? `
        <div class="sidebar-section-label" aria-hidden="true">${isAdmin ? "Admin" : "My Forms"}</div>
        ${navItem("admin-review", adminIcon(), "Form Requests")}
        ${navItem("new-form", newFormIcon(), "Form Builder")}
        ` : ""}
      </nav>

      <!-- MAIN CONTENT -->
      <main class="main" id="main-content" tabindex="-1">
        <div class="loading-overlay" id="global-loader" style="position:fixed;background:rgba(10,12,16,0.5);z-index:200;display:none;" role="status" aria-label="Loading">
          <div class="spinner" aria-hidden="true" style="width:32px;height:32px;border-width:3px;"></div>
        </div>
      </main>
    </div>
  `;

  navigateTo(currentView);

  const saved = localStorage.getItem("fs-theme") || "uol";
  const sel = document.getElementById("theme-select");
  if (sel) sel.value = saved;
}

function navItem(key, icon, label) {
  const active = AppState.currentView === key;
  return `<div
    class="nav-item${active ? " active" : ""}"
    role="button"
    tabindex="0"
    data-view="${key}"
    aria-current="${active ? "page" : "false"}"
    aria-label="${label}"
    onclick="navigateTo('${key}')"
    onkeydown="if(event.key==='Enter'||event.key===' '){event.preventDefault();navigateTo('${key}')}"
  >${icon}<span>${label}</span></div>`;
}

function myFormsIcon()     { return `<svg class="nav-icon" viewBox="0 0 16 16" fill="currentColor" aria-hidden="true"><path d="M2 2a2 2 0 012-2h8a2 2 0 012 2v12a2 2 0 01-2 2H4a2 2 0 01-2-2V2zm2 0v12h8V2H4zm1 2h6v1H5V4zm0 2.5h6v1H5v-1zm0 2.5h4v1H5V9z"/></svg>`; }
function homeIcon()        { return `<svg class="nav-icon" viewBox="0 0 16 16" fill="currentColor" aria-hidden="true"><path d="M8.354 1.146a.5.5 0 00-.708 0l-6 6A.5.5 0 002 7.5v7a.5.5 0 00.5.5h4a.5.5 0 00.5-.5v-4h2v4a.5.5 0 00.5.5h4a.5.5 0 00.5-.5v-7a.5.5 0 00-.146-.354L8.354 1.146z"/></svg>`; }
function formRequestsIcon(){ return `<svg class="nav-icon" viewBox="0 0 16 16" fill="currentColor" aria-hidden="true"><path d="M4 0a2 2 0 00-2 2v12a2 2 0 002 2h8a2 2 0 002-2V2a2 2 0 00-2-2H4zm0 1.5h8a.5.5 0 01.5.5v12a.5.5 0 01-.5.5H4a.5.5 0 01-.5-.5V2a.5.5 0 01.5-.5zm1 2.5h6v1H5V4zm0 2.5h6v1H5v-1zm0 2.5h4v1H5V9z"/></svg>`; }
function newFormIcon()     { return `<svg class="nav-icon" viewBox="0 0 16 16" fill="currentColor" aria-hidden="true"><path d="M8 2a6 6 0 100 12A6 6 0 008 2zm.75 3.25v2h2a.75.75 0 010 1.5h-2v2a.75.75 0 01-1.5 0v-2h-2a.75.75 0 010-1.5h2v-2a.75.75 0 011.5 0z"/></svg>`; }
function liveFormsIcon()   { return `<svg class="nav-icon" viewBox="0 0 16 16" fill="currentColor" aria-hidden="true"><path d="M1 2.75A.75.75 0 011.75 2h12.5a.75.75 0 010 1.5H1.75A.75.75 0 011 2.75zm0 5A.75.75 0 011.75 7h12.5a.75.75 0 010 1.5H1.75A.75.75 0 011 7.75zm0 5A.75.75 0 011.75 12h7.5a.75.75 0 010 1.5h-7.5A.75.75 0 011 12.75z"/></svg>`; }
function submissionsIcon() { return `<svg class="nav-icon" viewBox="0 0 16 16" fill="currentColor" aria-hidden="true"><path d="M1.5 2.5A1.5 1.5 0 013 1h7.879a1.5 1.5 0 011.06.44l2.122 2.12A1.5 1.5 0 0114.5 4.62V13.5A1.5 1.5 0 0113 15H3a1.5 1.5 0 01-1.5-1.5v-11zM3 2.5v11h10V5H11a1.5 1.5 0 01-1.5-1.5V2.5H3zm7 0v1.5H11.5L10 2.5z"/></svg>`; }
function adminIcon()       { return `<svg class="nav-icon" viewBox="0 0 16 16" fill="currentColor" aria-hidden="true"><path d="M7.122.392a1.75 1.75 0 011.756 0l5.25 3.045c.54.313.872.89.872 1.514V8.64c0 2.048-1.19 3.914-3.05 4.856l-2.5 1.286a1.75 1.75 0 01-1.6 0l-2.5-1.286C3.19 12.554 2 10.688 2 8.64V4.951c0-.624.332-1.2.872-1.514L7.122.392z"/></svg>`; }
function liveManageIcon()  { return `<svg class="nav-icon" viewBox="0 0 16 16" fill="currentColor" aria-hidden="true"><path d="M8 9.5a1.5 1.5 0 100-3 1.5 1.5 0 000 3z"/><path d="M8 0a.75.75 0 01.695.47l1.498 3.654 3.926.573a.75.75 0 01.416 1.28l-2.841 2.768.671 3.91a.75.75 0 01-1.088.79L8 11.46l-3.277 1.985a.75.75 0 01-1.088-.79l.671-3.91-2.84-2.768a.75.75 0 01.415-1.28l3.926-.573L7.305.47A.75.75 0 018 0z"/></svg>`; }
