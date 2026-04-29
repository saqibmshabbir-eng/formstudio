// Form Studio — Home Screen

// =============================================================
// HOME — form portal: live forms + past submissions
// =============================================================
async function renderHome(container) {
  const { currentUser, isAdmin } = AppState;
  const displayName = currentUser?.displayName || "";
  // University accounts use "Surname, Firstname" — take the part after the comma.
  // If there's no comma (personal / non-university account), use the first word.
  const name = displayName.includes(",")
    ? (displayName.split(",")[1]?.trim() || displayName.split(",")[0]?.trim() || "there")
    : (displayName.split(" ")[0]?.trim() || "there");

  container.innerHTML = `
    <div style="max-width:900px;margin:0 auto;">
      <div style="margin-bottom:32px;">
        <h1 style="font-size:26px;font-weight:600;letter-spacing:-0.02em;">Hi, ${escHtml(name)}</h1>
        <p style="color:var(--text2);font-size:14px;margin-top:4px;">Select a form below to get started</p>
      </div>
      <div id="home-forms-section">
        <div style="text-align:center;padding:48px;"><span class="spinner" style="width:28px;height:28px;border-width:3px;"></span></div>
      </div>
    </div>
  `;

  await loadHomeForms();
}

async function loadHomeForms() {
  const el = document.getElementById("home-forms-section");
  if (!el) return;
  try {
    const items = await getListItems(CONFIG.FORMS_LIST);
    const { isAdmin, currentUser } = AppState;

    // Determine Preview visibility — use SP createdBy (authoritative) rather than JSON
    // getListItems already expands createdBy so it's available on each item
    const visible = items.filter(i => {
      const f = i.fields || {};
      const s = f[CONFIG.COL_STATUS] || f.Status;
      // Retro forms — always show if they have a ListLocation
      if (f[CONFIG.COL_RETRO]) return !!f[CONFIG.COL_LIST_LOCATION];
      if (s === "Live") return true;
      if (s === "Preview") {
        if (isAdmin) return true;
        const createdByEmail = (i.createdBy?.user?.email || "").toLowerCase();
        return createdByEmail && currentUser?.email &&
               createdByEmail === currentUser.email.toLowerCase();
      }
      return false;
    });

    if (!visible.length) {
      el.innerHTML = `
        <div class="empty-state">
          <svg width="40" height="40" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z"/><polyline points="14 2 14 8 20 8"/></svg>
          <h3 style="margin-top:12px;">No forms available yet</h3>
          <p>Check back soon — forms will appear here when they go live.</p>
        </div>`;
      return;
    }

    el.innerHTML = `
      <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:14px;">
        <span style="font-size:11px;text-transform:uppercase;letter-spacing:0.07em;color:var(--nav-icon);font-weight:600;">Available Forms (${visible.length})</span>
        ${visible.length > 6 ? `<button class="btn btn-ghost btn-sm" onclick="showFormsModal()">View all →</button>` : ""}
      </div>
      <div style="display:flex;flex-direction:column;gap:6px;">
        ${visible.slice(0,6).map(item => renderFormRow(item)).join("")}
        ${visible.length > 6 ? `
          <div style="text-align:center;padding:10px;">
            <button class="btn btn-secondary btn-sm" onclick="showFormsModal()">
              Show all ${visible.length} forms
            </button>
          </div>
        ` : ""}
      </div>
    `;
  } catch (e) {
    if (el) el.innerHTML = `<div class="empty-state"><p style="color:var(--red)">Error loading forms: ${escHtml(e.message)}</p></div>`;
  }
}

function renderFormRow(item) {
  const f = item.fields || {};
  const status = f[CONFIG.COL_STATUS] || f.Status;
  const isPreview = status === "Preview";
  const isRetro = !!f[CONFIG.COL_RETRO];
  const onclick = isRetro
    ? `window.open(${JSON.stringify(f[CONFIG.COL_LIST_LOCATION])},'_blank')`
    : `openLiveForm(this.dataset.id)`;
  return html`
    <div data-id="${item.id}" onclick="${onclick}"
      style="background:var(--input-bg);border:1px solid var(--input-border);border-radius:var(--radius-sm);padding:12px 16px;cursor:pointer;transition:var(--transition);display:flex;align-items:center;gap:12px;"
      onmouseover="this.style.borderColor='var(--input-focus)';this.style.background='#f0f4ff'"
      onmouseout="this.style.borderColor='var(--border)';this.style.background='#fff'">
      <div style="width:32px;height:32px;border-radius:8px;background:linear-gradient(135deg,var(--accent),var(--accent2));display:flex;align-items:center;justify-content:center;flex-shrink:0;">
        ${safeHtml(isRetro
          ? `<svg width="14" height="14" viewBox="0 0 16 16" fill="white"><path d="M2 2h12v12H2z" stroke="white" stroke-width="1" fill="none"/><path d="M5 6h6M5 8h6M5 10h4" stroke="white" stroke-width="1.2" stroke-linecap="round"/></svg>`
          : `<svg width="14" height="14" viewBox="0 0 16 16" fill="white"><path d="M4 0a2 2 0 00-2 2v12a2 2 0 002 2h8a2 2 0 002-2V2a2 2 0 00-2-2H4zm1 3h6v1H5V3zm0 2.5h6v1H5v-1zm0 2.5h4v1H5V8z"/></svg>`
        )}
      </div>
      <div style="flex:1;min-width:0;">
        <div style="font-size:14px;font-weight:600;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">${f.Title || "Untitled Form"}</div>
        ${safeHtml(isRetro ? `<div style="font-size:11.5px;color:var(--text3);margin-top:1px;">Opens in SharePoint</div>` : "")}
      </div>
      ${safeHtml(isPreview ? `<span class="badge badge-amber" style="flex-shrink:0;">Preview</span>` : "")}
      <svg width="14" height="14" viewBox="0 0 14 14" fill="none" stroke="var(--text3)" stroke-width="1.5" style="flex-shrink:0;"><path d="M3 7h8M7 3l4 4-4 4"/></svg>
    </div>
  `;
}

async function showFormsModal() {
  const currentEmail = (AppState.currentUser?.email || "").toLowerCase();
  const items = await getListItems(CONFIG.FORMS_LIST);
  const forms = items.filter(i => {
    const f = i.fields || {};
    const s = f[CONFIG.COL_STATUS] || f.Status;
    if (f[CONFIG.COL_RETRO]) return !!f[CONFIG.COL_LIST_LOCATION];
    if (s === "Live") return true;
    if (s === "Preview") {
      if (AppState.isAdmin) return true;
      return (i.createdBy?.user?.email || "").toLowerCase() === currentEmail;
    }
    return false;
  });
  openModal(html`
    <div class="modal-header">
      <span class="modal-title">Available Forms</span>
      <button class="btn btn-ghost btn-icon btn-sm" onclick="closeModal()">
        <svg width="14" height="14" viewBox="0 0 14 14" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M1 1l12 12M13 1L1 13"/></svg>
      </button>
    </div>
    <div style="padding:16px 24px 8px;">
      <input id="forms-modal-search" class="input" placeholder="Search forms…"
        oninput="filterFormsModal(this.value)"
        style="font-size:14px;">
    </div>
    <div class="modal-body" id="forms-modal-list" style="padding-top:8px;max-height:60vh;overflow-y:auto;">
      ${safeHtml(forms.map(item => {
        const f = item.fields || {};
        const isRetro = !!f[CONFIG.COL_RETRO];
        const onclick = isRetro
          ? `closeModal();window.open(${JSON.stringify(f[CONFIG.COL_LIST_LOCATION])},'_blank')`
          : `closeModal();openLiveForm(this.dataset.id)`;
        return html`
        <div data-id="${item.id}" onclick="${onclick}"
          style="padding:12px 0;border-bottom:1px solid var(--border);cursor:pointer;display:flex;align-items:center;gap:10px;"
          onmouseover="this.style.background='#f0f4ff';this.style.margin='0 -4px';this.style.padding='12px 4px'"
          onmouseout="this.style.background='';this.style.margin='';this.style.padding='12px 0'"
          data-title="${(f.Title||'').toLowerCase()}">
          <div style="width:28px;height:28px;border-radius:6px;background:linear-gradient(135deg,var(--accent),var(--accent2));display:flex;align-items:center;justify-content:center;flex-shrink:0;">
            <svg width="12" height="12" viewBox="0 0 16 16" fill="white"><path d="M4 0a2 2 0 00-2 2v12a2 2 0 002 2h8a2 2 0 002-2V2a2 2 0 00-2-2H4zm1 3h6v1H5V3zm0 2.5h6v1H5v-1zm0 2.5h4v1H5V8z"/></svg>
          </div>
          <span style="font-size:14px;font-weight:500;flex:1;">${f.Title||"Untitled"}</span>
          ${safeHtml((f[CONFIG.COL_STATUS]||f.Status)==="Preview" ? `<span class="badge badge-amber">Preview</span>` : "")}
          ${safeHtml(isRetro ? `<span style="font-size:11px;color:var(--text3);">SharePoint ↗</span>` : "")}
        </div>
      `}).join(""))}
    </div>
  `);
}

function filterFormsModal(query) {
  const q = query.toLowerCase().trim();
  document.querySelectorAll("#forms-modal-list [data-title]").forEach(el => {
    el.style.display = !q || el.dataset.title.includes(q) ? "" : "none";
  });
}

