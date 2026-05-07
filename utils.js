// Form Studio — Utilities, Boot & Theme

// =============================================================
// USER MENU
// =============================================================
function showUserMenu() {
  openModal(`
    <div class="modal-header">
      <span class="modal-title">Account</span>
      <button class="btn btn-ghost btn-icon btn-sm" onclick="closeModal()">
        <svg width="14" height="14" viewBox="0 0 14 14" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M1 1l12 12M13 1L1 13"/></svg>
      </button>
    </div>
    <div class="modal-body">
      <div class="flex items-center gap-3">
        <div class="avatar" style="width:40px;height:40px;font-size:16px;">${AppState.currentUser?.displayName.split(" ").map(n=>n[0]).join("").slice(0,2)||"?"}</div>
        <div>
          <div style="font-weight:500;">${escHtml(AppState.currentUser?.displayName||"")}</div>
          <div style="font-size:12.5px;color:var(--text2)">${escHtml(AppState.currentUser?.email||"")}</div>
        </div>
      </div>
      ${AppState.isAdmin ? `<div class="mt-2"><span class="badge badge-purple">Admin</span></div>` : ""}
    </div>
    <div class="modal-footer">
      <button class="btn btn-ghost" onclick="closeModal()">Close</button>
      <button class="btn btn-danger" onclick="logout()">Sign Out</button>
    </div>
  `);
}
// =============================================================
// UTILITIES
// =============================================================
function getAllFields() {
  return (AppState.builderForm.sections || []).flatMap(s => s.fields || []);
}

// Derives a SharePoint-safe column prefix from a section title.
// Strips all non-alphanumeric characters and takes the first 20 chars.
// Must produce the same value in builder.js, admin.js, and live-form.js
// so that column names are consistent across provisioning and rendering.
// Example: "Dept Review & Sign-off" → "DeptReviewSignoff"
function sectionKey(section) {
  return (section.title || "Section")
    .replace(/[^a-zA-Z0-9]/g, "")
    .slice(0, 20) || "Section";
}

function getAllFieldsFromDef(def) {
  return (def.sections || []).flatMap(s => s.fields || []);
}

function escHtml(str) {
  if (str == null) return "";
  return String(str).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;");
}

function escAttr(str) {
  if (str == null) return "";
  return String(str).replace(/'/g,"&#39;").replace(/"/g,"&quot;");
}

// =============================================================
// TAGGED TEMPLATE LITERAL — auto-escaping HTML builder
// Usage:  html`<div>${userValue}</div>`
// All interpolated values are HTML-escaped automatically.
// For deliberate raw HTML (e.g. InfoText sanitised by DOMPurify),
// wrap the value: html`<div>${safeHtml(sanitised)}</div>`
// =============================================================
function safeHtml(value) {
  // Marks a string as pre-sanitised — bypasses auto-escaping in html``
  return { __safe: true, value: value == null ? "" : String(value) };
}

function html(strings, ...values) {
  let out = strings[0];
  for (let i = 0; i < values.length; i++) {
    const val = values[i];
    if (val == null) {
      out += "";
    } else if (val.__safe === true) {
      out += val.value;                      // pre-sanitised — pass through as-is
    } else if (Array.isArray(val)) {
      out += val.join("");                   // already-rendered html`` segments
    } else {
      out += escHtml(String(val));           // escape everything else
    }
    out += strings[i + 1];
  }
  return out;
}

function formatDate(iso) {
  if (!iso) return "—";
  try {
    return new Date(iso).toLocaleDateString("en-GB", { day:"2-digit", month:"short", year:"numeric" });
  } catch (_) { return "—"; }
}

function statusBadge(status) {
  const map = {
    "Draft":     "badge-gray",
    "Created":   "badge-gray",
    "Submitted": "badge-blue",
    "Preview":   "badge-amber",
    "Approved":  "badge-green",
    "Rejected":  "badge-red",
    "Live":      "badge-green",
    "Closed":    "badge-gray",
  };
  return `<span class="badge ${map[status]||"badge-gray"}">${escHtml(status)}</span>`;
}
// =============================================================
// BOOT
// =============================================================
async function bootApp() {
  try {
    // Get current user info
    const me = await graphGet("/me?$select=displayName,mail,id");
    AppState.currentUser = {
      displayName: me.displayName || currentAccount.name,
      email: me.mail || currentAccount.username,
      id: me.id,
    };

    // Get SharePoint user ID for AuthorLookupId matching
    try {
      const siteId = await getSiteId();
      const spUser = await graphGet(`/sites/${siteId}/lists('User Information List')/items?$filter=fields/EMail eq '${me.mail}'&$expand=fields`);
      if (spUser.value?.[0]) {
        AppState._currentUserSpId = parseInt(spUser.value[0].id);
      }
    } catch (_) {}

    // Check permissions in parallel
    const [adminStatus, formReadAccess] = await Promise.all([
      checkIsAdmin(AppState.currentUser.displayName, AppState.currentUser.email),
      checkListReadAccess(CONFIG.FORMS_LIST),
    ]);

    AppState.isAdmin = adminStatus;
    AppState.hasFormRequestAccess = formReadAccess;

    // Set initial view
    AppState.currentView = "home";

    renderAppShell();

    // Handle deep links — e.g. ?view=my-forms&formId=abc123&itemId=def456
    const params = new URLSearchParams(window.location.search);
    const deepView   = params.get("view");
    const deepFormId = params.get("formId");
    const deepItemId = params.get("itemId");
    if (deepView === "my-forms" && deepFormId) {
      // Navigate to my-forms view then drill straight into the form submissions.
      // If itemId is also present, open that specific submission directly.
      AppState.currentView = "my-forms";
      renderAppShell();
      openFormSubmissions(deepFormId, deepItemId || null);
    } else if (deepView) {
      navigateTo(deepView);
    }
  } catch (e) {
    showToast("error", "Startup error: " + e.message);
    renderAppShell(); // Render anyway with defaults
  }
}

async function main() {
  await initMsal();

  const accounts = msalInstance.getAllAccounts();
  if (!accounts.length) {
    render();
    return;
  }

  currentAccount = accounts[0];
  try {
    // Silent-only on boot — never open a popup without a user gesture, or the
    // browser will block it and the user will see "popups blocked" with no
    // obvious recovery. If silent acquisition fails (refresh token expired,
    // scope change, MFA re-prompt required, etc.), drop the user to the login
    // screen so their click on "Sign in with Microsoft" provides the gesture
    // loginPopup needs.
    await getToken({ allowPopup: false });
    await bootApp();
  } catch (_) {
    currentAccount = null;
    renderLoginScreen();
  }
}

// =============================================================
// THEME SYSTEM
// =============================================================
const THEMES = ['uol','google','apple','dark','glass','pastel','candy','ocean','slate','hc','premium'];

function setTheme(theme) {
  if (!THEMES.includes(theme)) theme = 'uol';
  document.documentElement.setAttribute('data-theme', theme);
  localStorage.setItem('fs-theme', theme);
  // Sync select if it exists
  const sel = document.getElementById('theme-select');
  if (sel) sel.value = theme;
}

function loadTheme() {
  const saved = localStorage.getItem('fs-theme') || 'uol';
  setTheme(saved);
}

// Call immediately so theme applies before any render
loadTheme();

// =============================================================
// GENERIC PEOPLE SEARCH
// Replaces four near-identical debounced search implementations
// across builder.js and my-forms.js.
//
// Usage:
//   createPeopleSearch({
//     inputId:    "managers-search",    // ID of the text input
//     resultsId:  "managers-results",   // ID of the results container
//     onClickFn:  "addManagerFromEl",   // global fn name called on row click
//     extraData:  { formid: "123" },    // optional extra data-* attrs on each row
//   })
//
// The caller still defines its own onClickFn (addManagerFromEl, addPersonFromEl etc.)
// because the action on selection differs per context.
// =============================================================
const _peopleSearchTimers = {};

function createPeopleSearch({ inputId, resultsId, onClickFn, extraData = {} }) {
  // Returns { search, debounced } so the caller can wire up oninput and button onclick
  async function search(queryOverride) {
    const query = queryOverride !== undefined
      ? queryOverride
      : document.getElementById(inputId)?.value;
    if (!query || query.length < 2) return;

    const resultsEl = document.getElementById(resultsId);
    if (!resultsEl) return;
    resultsEl.innerHTML = `<span class="spinner"></span>`;

    try {
      const people = await searchPeople(query);
      if (!people.length) {
        resultsEl.innerHTML = `<p style="font-size:12.5px;color:var(--text3);">No results found.</p>`;
        return;
      }

      // Build extra data-* attributes string from extraData object
      const extraAttrs = Object.entries(extraData)
        .map(([k, v]) => `data-${k}="${escAttr(String(v))}"`).join(" ");

      resultsEl.innerHTML = html`
        <div style="background:var(--bg3);border:1px solid var(--border);border-radius:var(--radius-sm);overflow:hidden;">
          ${safeHtml(people.slice(0, 6).map(p => {
            const email = p.scoredEmailAddresses?.[0]?.address || "";
            const initials = p.displayName.split(" ").map(n => n[0]).join("").slice(0, 2);
            return html`
              <div class="flex items-center gap-2"
                style="padding:8px 12px;cursor:pointer;border-bottom:1px solid var(--border);"
                data-id="${p.id}"
                data-name="${p.displayName}"
                data-email="${email}"
                ${extraAttrs}
                onclick="${onClickFn}(this)"
                onmouseover="this.style.background='var(--surface)'"
                onmouseout="this.style.background=''">
                <div class="avatar" style="width:24px;height:24px;font-size:10px;">${initials}</div>
                <div style="flex:1;">
                  <div style="font-size:13px;">${p.displayName}</div>
                  <div style="font-size:11.5px;color:var(--text3);">${email}</div>
                </div>
              </div>
            `;
          }).join(""))}
        </div>
      `;
    } catch (e) {
      const resultsEl2 = document.getElementById(resultsId);
      if (resultsEl2) resultsEl2.innerHTML =
        html`<p style="font-size:12.5px;color:var(--red)">Search failed: ${e.message}</p>`;
    }
  }

  function debounced(val) {
    clearTimeout(_peopleSearchTimers[inputId]);
    _peopleSearchTimers[inputId] = setTimeout(() => search(val), 400);
  }

  return { search, debounced };
}

// =============================================================
// Defer main() until MSAL is confirmed loaded
function startApp() {
  console.log("[FormStudio] startApp called, msal defined:", typeof msal !== "undefined");
  if (typeof msal === "undefined") {
    document.getElementById("app").innerHTML = `
      <div style="padding:60px;text-align:center;color:var(--red);">
        <h2>Startup Error</h2>
        <pre style="margin-top:12px;font-size:12px;color:var(--text2);white-space:pre-wrap;">MSAL failed to load. The CDN script did not load — check that the server can reach unpkg.com.</pre>
      </div>
    `;
    return;
  }
  main().catch(e => {
    console.error("[FormStudio] main() error:", e);
    document.getElementById("app").innerHTML = `
      <div style="padding:60px;text-align:center;color:var(--red);">
        <h2>Startup Error</h2>
        <pre style="margin-top:12px;font-size:12px;color:var(--text2);white-space:pre-wrap;">${e.message}\n\n${e.stack||""}</pre>
      </div>
    `;
  });
}

// Fallback: if startApp is never called (MSAL script blocked), show error after 5s
setTimeout(() => {
  const app = document.getElementById("app");
  if (app && app.innerHTML.trim() === "") {
    app.innerHTML = `
      <div style="padding:60px;text-align:center;color:var(--red);">
        <h2>Startup Error</h2>
        <pre style="margin-top:12px;font-size:12px;color:#8891a8;white-space:pre-wrap;">The app did not start. Possible causes:\n- MSAL CDN script blocked by network/CSP\n- JavaScript error before startApp() fired\n\nOpen browser DevTools (F12) → Console for details.</pre>
      </div>
    `;
  }
}, 5000);