// Form Studio — Navigation & App Shell

// =============================================================
// SCREEN READER ANNOUNCEMENT HELPER
// =============================================================
function srAnnounce(message) {
  const live = document.getElementById("sr-live");
  if (!live) return;
  live.textContent = "";
  requestAnimationFrame(() => { live.textContent = message; });
}

// =============================================================
// MOBILE BOTTOM TAB BAR
// =============================================================
function renderMobileNav() {
  // Remove existing if present
  const existing = document.getElementById("mobile-nav");
  if (existing) existing.remove();

  const { isAdmin, hasFormRequestAccess } = AppState;
  const view = AppState.currentView;

  const tabs = [
    { key: "home", label: "Home", icon: `<svg viewBox="0 0 16 16" fill="currentColor" aria-hidden="true"><path d="M8.354 1.146a.5.5 0 00-.708 0l-6 6A.5.5 0 002 7.5v7a.5.5 0 00.5.5h4a.5.5 0 00.5-.5v-4h2v4a.5.5 0 00.5.5h4a.5.5 0 00.5-.5v-7a.5.5 0 00-.146-.354L8.354 1.146z"/></svg>` },
    { key: "my-forms", label: "My Forms", icon: `<svg viewBox="0 0 16 16" fill="currentColor" aria-hidden="true"><path d="M9 12h6M9 16h6M9 8h6M5 3h14a2 2 0 012 2v14a2 2 0 01-2 2H5a2 2 0 01-2-2V5a2 2 0 012-2z"/></svg>` },
    ...(isAdmin || hasFormRequestAccess ? [
      { key: "admin-review", label: "Requests", icon: `<svg viewBox="0 0 16 16" fill="currentColor" aria-hidden="true"><path d="M4 0a2 2 0 00-2 2v12a2 2 0 002 2h8a2 2 0 002-2V2a2 2 0 00-2-2H4zm1 3h6v1H5V4zm0 2.5h6v1H5v-1zm0 2.5h4v1H5V9z"/></svg>` },
      { key: "new-form",     label: "Builder",  icon: `<svg viewBox="0 0 16 16" fill="currentColor" aria-hidden="true"><path d="M8 2a6 6 0 100 12A6 6 0 008 2zm.75 3.25v2h2a.75.75 0 010 1.5h-2v2a.75.75 0 01-1.5 0v-2h-2a.75.75 0 010-1.5h2v-2a.75.75 0 011.5 0z"/></svg>` },
    ] : []),
  ];

  const nav = document.createElement("nav");
  nav.id = "mobile-nav";
  nav.className = "mobile-nav";
  nav.setAttribute("aria-label", "Mobile navigation");
  nav.innerHTML = tabs.map(tab => `
    <button
      class="mobile-nav-item${tab.key === view ? " active" : ""}"
      onclick="navigateTo('${tab.key}')"
      aria-label="${tab.label}"
      aria-current="${tab.key === view ? "page" : "false"}"
    >
      ${tab.icon}
      <span>${tab.label}</span>
    </button>
  `).join("");

  document.body.appendChild(nav);
}

// =============================================================
// NAVIGATION
// =============================================================
function navigateTo(view) {
  AppState.currentView = view;

  // Update sidebar active state and aria-current
  document.querySelectorAll(".nav-item").forEach(el => {
    const isActive = el.getAttribute("data-view") === view;
    el.classList.toggle("active", isActive);
    el.setAttribute("aria-current", isActive ? "page" : "false");
  });

  const main = document.getElementById("main-content");
  if (!main) return;

  const viewLabels = {
    "home":           "Home",
    "live-forms":     "Available Forms",
    "my-submissions": "My Submissions",
    "my-forms":       "My Forms",
    "admin-review":   "Form Requests",
    "new-form":       "Form Builder",
  };

  switch (view) {
    case "home":           renderHome(main); break;
    case "live-forms":     renderLiveForms(main); break;
    case "my-submissions": renderMySubmissions(main); break;
    case "my-forms":       renderMyForms(main); break;
    case "admin-review":   renderAdminReview(main); break;
    case "new-form":       startNewForm(main); break;
    default:               renderHome(main);
  }

  // Update / create mobile bottom nav
  renderMobileNav();

  // Announce navigation to screen readers
  srAnnounce(viewLabels[view] || "Page loaded");

  // Move focus to main content heading
  requestAnimationFrame(() => {
    const h1 = main.querySelector("h1");
    if (h1) {
      h1.setAttribute("tabindex", "-1");
      h1.focus({ preventScroll: true });
    } else {
      main.setAttribute("tabindex", "-1");
      main.focus({ preventScroll: true });
    }
  });
}
