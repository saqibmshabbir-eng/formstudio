// Form Studio — My Forms (Author & Manager submissions dashboard)

// =============================================================
// MY FORMS — FORM LIST
// =============================================================
async function renderMyForms(container) {
  container.innerHTML = html`
    <div class="flex items-center justify-between mb-4">
      <div>
        <h1 style="font-size:22px;font-weight:600;letter-spacing:-0.02em;">My Forms</h1>
        <p style="color:var(--text2);font-size:13.5px;margin-top:2px;">View and manage submissions for forms you own or manage</p>
      </div>
    </div>
    <div id="my-forms-list">
      <div style="padding:40px;text-align:center;"><span class="spinner"></span></div>
    </div>
  `;

  try {
    const currentEmail = (AppState.currentUser?.email || "").toLowerCase();
    const allItems     = await getListItems(CONFIG.FORMS_LIST);

    // Candidate forms — Live or Preview, plus retro forms that have a ListName.
    // Retro forms are included on the same basis as regular forms — we query
    // their data list and only show the pill if it has items.
    const candidates = allItems.filter(i => {
      const s        = i.fields?.[CONFIG.COL_STATUS] || "";
      const listName = i.fields?.[CONFIG.COL_LISTNAME] || "";
      const isRetro  = !!i.fields?.[CONFIG.COL_RETRO];

      if (isRetro) return !!listName; // retro: needs a list name to query
      return (s === "Live" || s === "Preview") && !!listName;
    });

    // For each candidate, fire a $top=1 query against its data list.
    // SharePoint's ReadSecurity=2 means each user sees only their own items —
    // no client-side filter needed. Admins and managers see all items.
    // We filter out soft-deleted items so an all-deleted list shows no pill.
    const siteId = await getSiteId();

    const hasItemsResults = await Promise.allSettled(
      candidates.map(async item => {
        const listName = item.fields[CONFIG.COL_LISTNAME];
        try {
          const listId   = await getListId(listName);
          // Probe only — we just need to know the list has at least one item.
          // No filter or expand needed; soft-deleted items are filtered properly
          // when the full submissions load runs (line 199).
          const response = await graphGet(
            `/sites/${siteId}/lists/${listId}/items?$top=1`
          );
          return (response?.value?.length || 0) > 0;
        } catch (_) {
          // List doesn't exist or user has no access — treat as no items
          return false;
        }
      })
    );

    // Keep only candidates whose list returned at least one item
    const formsWithItems = candidates.filter((_, idx) =>
      hasItemsResults[idx].status === "fulfilled" && hasItemsResults[idx].value === true
    );

    if (!formsWithItems.length) {
      document.getElementById("my-forms-list").innerHTML = html`
        <div class="empty-state">
          <svg width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M9 12h6M9 16h6M9 8h6M5 3h14a2 2 0 012 2v14a2 2 0 01-2 2H5a2 2 0 01-2-2V5a2 2 0 012-2z"/></svg>
          <h3 style="margin-top:12px;">No forms yet</h3>
          <p>Forms you create or manage will appear here once they are Live or in Preview.</p>
        </div>
      `;
      return;
    }

    // Load definitions in parallel for the forms that passed — needed only
    // for the role badge (Author / Manager). Failures are non-fatal.
    const defResults = await Promise.allSettled(
      formsWithItems.map(i => getFormDefinition(CONFIG.FORMS_LIST, i.id))
    );

    // Derive role badge per form — Author > Manager > none (submitter)
    function roleFor(item, def) {
      const createdByEmail = (item.createdBy?.user?.email || "").toLowerCase();
      if (createdByEmail === currentEmail) return "author";
      if (def?.formManagers?.some(m => (m.email || "").toLowerCase() === currentEmail)) return "manager";
      return null;
    }

    document.getElementById("my-forms-list").innerHTML = html`
      <div style="display:grid;grid-template-columns:repeat(auto-fill,minmax(300px,1fr));gap:16px;">
        ${safeHtml(formsWithItems.map((item, idx) => {
          const f       = item.fields || {};
          const isRetro = !!f[CONFIG.COL_RETRO];
          const status  = f[CONFIG.COL_STATUS] || "Live";
          const def     = defResults[idx].status === "fulfilled" ? defResults[idx].value : null;
          const role    = roleFor(item, def);

          const onclick = isRetro
            ? `window.open(${JSON.stringify(f[CONFIG.COL_VIEW_URL] || "")},'_blank')`
            : `openFormSubmissions(this.dataset.id)`;

          return html`
            <div class="card" style="cursor:pointer;transition:var(--transition);"
              data-id="${item.id}" onclick="${onclick}"
              onmouseover="this.style.borderColor='var(--border2)'"
              onmouseout="this.style.borderColor='var(--border)'">
              <div class="card-body" style="padding:20px;">
                <div class="flex items-center gap-2 mb-3" style="flex-wrap:wrap;">
                  ${safeHtml(isRetro
                    ? `<span class="badge badge-purple" style="font-size:10px;">SharePoint</span>`
                    : statusBadge(status)
                  )}
                  ${safeHtml(role === "author"
                    ? `<span class="badge badge-blue" style="font-size:10px;">Author</span>`
                    : role === "manager"
                      ? `<span class="badge badge-purple" style="font-size:10px;">Manager</span>`
                      : ""
                  )}
                </div>
                <div style="font-size:16px;font-weight:600;margin-bottom:4px;">${f.Title || "Untitled"}</div>
                <div style="font-size:12px;color:var(--text3);font-family:var(--mono);margin-bottom:12px;">
                  ${isRetro ? "External SharePoint list" : (f[CONFIG.COL_LISTNAME] || "")}
                </div>
                ${safeHtml(isRetro ? `
                  <div style="font-size:12px;color:var(--text3);display:flex;align-items:center;gap:4px;">
                    <svg width="11" height="11" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M7 3H3a1 1 0 00-1 1v9a1 1 0 001 1h9a1 1 0 001-1V9M9 1h6m0 0v6m0-6L7 9"/></svg>
                    Opens in SharePoint
                  </div>` : ""
                )}
              </div>
            </div>
          `;
        }).join(""))}
      </div>
    `;
  } catch (e) {
    const el = document.getElementById("my-forms-list");
    if (el) el.innerHTML =
      html`<div class="empty-state"><p style="color:var(--red)">Error loading forms: ${e.message}</p></div>`;
  }
}

// =============================================================
// SUBMISSIONS TABLE — drill-down per form
// =============================================================
let _currentFormItem = null;
let _currentFormDef  = null;
let _submissions     = [];
let _subFilter       = "";
let _subSortCol      = "Modified";
let _subSortAsc      = false;
let _subPage         = 1;
const _subPageSize   = 15;

// Format any SharePoint field value for display.
// Person/Group columns come back from Graph as { LookupId, LookupValue, Email }
// or as an array of those objects for multi-value columns.
function formatFieldValue(val) {
  if (val == null || val === "") return "—";
  if (Array.isArray(val)) {
    return val.map(v => formatFieldValue(v)).join(", ") || "—";
  }
  if (typeof val === "object") {
    // Person column: prefer LookupValue (display name), fallback to Email
    return val.LookupValue || val.Email || val.displayName || "—";
  }
  return String(val).replace(/<[^>]*>/g, "").trim() || "—";
}

async function openFormSubmissions(formItemId, deepItemId = null) {
  const main = document.getElementById("main-content");
  main.innerHTML = `<div style="padding:60px;text-align:center;"><span class="spinner" style="width:32px;height:32px;border-width:3px;"></span></div>`;

  try {
    const allItems = await getListItems(CONFIG.FORMS_LIST);
    _currentFormItem = allItems.find(i => i.id === formItemId) || null;
    if (!_currentFormItem) throw new Error("Form not found");

    _currentFormDef = await getFormDefinition(CONFIG.FORMS_LIST, formItemId);
    if (!_currentFormDef) throw new Error("Form definition not found");

    const listName = _currentFormItem.fields?.[CONFIG.COL_LISTNAME] || _currentFormDef.listName;
    if (!listName) throw new Error("No data list associated with this form");

    // Fetch all submissions — SharePoint ReadSecurity ensures users only see their own,
    // authors and managers with Contribute see everything
    let rawItems;
    try {
      const siteId = await getSiteId();
      const listId = await getListId(listName);
      const data = await graphGet(
        `/sites/${siteId}/lists/${listId}/items` +
        `?expand=fields($expand=${CONFIG.COL_ASSIGNED_TO}),createdBy`
      );
      rawItems = data.value || [];
    } catch (e) {
      if (e.message?.includes("404") || e.message?.includes("not found")) {
        throw new Error(`Access denied to list "${listName}". You may need to be granted explicit permissions — try re-provisioning the form via the admin Edit flow.`);
      }
      throw e;
    }
    // Filter out soft-deleted items by default
    _submissions = rawItems.filter(i => !i.fields?.IsDeleted);
    _subFilter  = "";
    _subSortCol = "Modified";
    _subSortAsc = false;
    _subPage    = 1;

    renderSubmissionsTable(main);

    // If a specific submission was linked to, open it directly after the table renders.
    // Managers get the full live form (Complete button available); others get the modal.
    if (deepItemId) {
      const currentEmail = (AppState.currentUser?.email || "").toLowerCase();
      const isManager = AppState.isAdmin ||
        (_currentFormDef?.formManagers || []).some(m => (m.email || "").toLowerCase() === currentEmail);
      viewOrOpenSubmission(deepItemId, formItemId, isManager);
    }
  } catch (e) {
    main.innerHTML = html`
      <div class="empty-state">
        <p style="color:var(--red)">Could not load submissions: ${e.message}</p>
        <button class="btn btn-secondary mt-4" onclick="navigateTo('my-forms')">Back to My Forms</button>
      </div>
    `;
  }
}

function renderSubmissionsTable(container) {
  const title = _currentFormItem?.fields?.Title || "Form";
  const allFields = (_currentFormDef?.sections || []).flatMap(s => s.fields || [])
    .filter(f => f.type !== "InfoText");

  // Determine current user role for this form
  const currentEmail = (AppState.currentUser?.email || "").toLowerCase();
  const authorEmail  = (_currentFormItem?.createdBy?.user?.email || "").toLowerCase();
  const isAuthor     = currentEmail === authorEmail;
  const isManager    = AppState.isAdmin || isAuthor ||
    (_currentFormDef?.formManagers || []).some(m => (m.email || "").toLowerCase() === currentEmail);

  // Filter
  const filterLower = _subFilter.toLowerCase();
  let filtered = _subFilter
    ? _submissions.filter(item => {
        const f = item.fields || {};
        const assignedName = String(f[CONFIG.COL_ASSIGNED_TO]?.LookupValue || "").toLowerCase();
        return allFields.some(field => {
          const val = String(f[field.internalName || field.label] || "").toLowerCase();
          return val.includes(filterLower);
        }) || formatDate(f.Modified).toLowerCase().includes(filterLower)
          || assignedName.includes(filterLower);
      })
    : [..._submissions];

  // Sort
  filtered.sort((a, b) => {
    let av = "", bv = "";
    if (_subSortCol === "Modified") {
      av = a.fields?.Modified || ""; bv = b.fields?.Modified || "";
    } else if (_subSortCol === "Created") {
      av = a.fields?.Created || ""; bv = b.fields?.Created || "";
    } else if (_subSortCol === "AssignedTo") {
      // Person column — sort alphabetically by display name; unassigned rows last
      av = a.fields?.[CONFIG.COL_ASSIGNED_TO]?.LookupValue || "";
      bv = b.fields?.[CONFIG.COL_ASSIGNED_TO]?.LookupValue || "";
      // Push empty values to the end regardless of asc/desc
      if (!av && bv) return 1;
      if (av && !bv) return -1;
    } else {
      const field = allFields.find(f => f.label === _subSortCol);
      if (field) {
        av = String(a.fields?.[field.internalName || field.label] || "");
        bv = String(b.fields?.[field.internalName || field.label] || "");
      }
    }
    const cmp = av.localeCompare(bv);
    return _subSortAsc ? cmp : -cmp;
  });

  // Paginate
  const total      = filtered.length;
  const totalPages = Math.max(1, Math.ceil(total / _subPageSize));
  _subPage         = Math.min(_subPage, totalPages);
  const start      = (_subPage - 1) * _subPageSize;
  const rows       = filtered.slice(start, start + _subPageSize);

  const arrow = col => _subSortCol === col
    ? (_subSortAsc ? " ↑" : " ↓") : "";

  // Visible columns — cap at 5 form fields, skip DateTime fields (redundant alongside "Submitted") and file uploads
  const visibleFields = allFields.filter(f => f.type !== "DateTime" && f.type !== "FileUpload").slice(0, 5);
  const hasFileUpload = allFields.some(f => f.type === "FileUpload");

  container.innerHTML = html`
    <div class="flex items-center gap-3 mb-4">
      <button class="btn btn-ghost btn-sm" onclick="navigateTo('my-forms')">
        <svg width="14" height="14" viewBox="0 0 14 14" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M9 2L4 7l5 5"/></svg>
        My Forms
      </button>
      <h1 style="font-size:20px;font-weight:600;letter-spacing:-0.02em;flex:1;">${title}</h1>
      ${safeHtml(isManager ? html`
        <button class="btn btn-ghost btn-sm" data-id="${_currentFormItem.id}" onclick="copyFormLink(this.dataset.id)" title="Copy link to this form's submissions">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M10 13a5 5 0 007.54.54l3-3a5 5 0 00-7.07-7.07l-1.72 1.71"/><path d="M14 11a5 5 0 00-7.54-.54l-3 3a5 5 0 007.07 7.07l1.71-1.71"/></svg>
          Copy Link
        </button>
        <button class="btn btn-ghost btn-sm" data-id="${_currentFormItem.id}" onclick="openManageManagers(this.dataset.id)">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M17 21v-2a4 4 0 00-4-4H5a4 4 0 00-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 00-3-3.87M16 3.13a4 4 0 010 7.75"/></svg>
          Managers
        </button>
      ` : "")}
      <button class="btn btn-secondary btn-sm" data-id="${_currentFormItem.id}" onclick="exportFormExcel(this.dataset.id)">
        <svg width="14" height="14" viewBox="0 0 24 24" fill="#1d6f42"><rect x="2" y="2" width="20" height="20" rx="3"/><path d="M7 7l3 5-3 5h2.5l1.75-3L13 17h2.5l-3-5 3-5H13l-1.75 3L9.5 7H7z" fill="white"/></svg>
        Export Excel
      </button>
    </div>

    <div class="card">
      <div style="padding:12px 16px;border-bottom:1px solid var(--border);display:flex;align-items:center;gap:10px;">
        <input class="input" style="max-width:280px;font-size:13px;padding:6px 10px;" placeholder="Search submissions…"
          value="${_subFilter}" oninput="_subFilter=this.value;_subPage=1;renderSubmissionsTable(document.getElementById('main-content'))">
        <span style="font-size:12.5px;color:var(--text2);margin-left:auto;">${total} submission${total !== 1 ? "s" : ""}</span>
      </div>

      <div class="table-wrap">
        <table>
          <thead>
            <tr>
              <th style="cursor:pointer;width:140px;" onclick="_subSortCol='AssignedTo';_subSortAsc=(_subSortCol==='AssignedTo'&&!_subSortAsc);renderSubmissionsTable(document.getElementById('main-content'))">
                Assigned To${arrow("AssignedTo")}
              </th>
              ${safeHtml(visibleFields.map(f => html`
                <th style="cursor:pointer;" data-col="${f.label}" onclick="_subSortCol=this.dataset.col;_subSortAsc=(_subSortCol===this.dataset.col&&!_subSortAsc);renderSubmissionsTable(document.getElementById('main-content'))">
                  ${f.label}${arrow(f.label)}
                </th>
              `).join(""))}
              <th style="cursor:pointer;" onclick="_subSortCol='Modified';_subSortAsc=(_subSortCol==='Modified'&&!_subSortAsc);renderSubmissionsTable(document.getElementById('main-content'))">
                Submitted${arrow("Modified")}
              </th>
              ${safeHtml(hasFileUpload ? `<th style="width:40px;text-align:center;" title="Attachments"><svg width="13" height="13" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M13.5 8.5l-5.5 5.5a4 4 0 01-5.66-5.66l6-6a2.5 2.5 0 013.54 3.54l-6.01 6a1 1 0 01-1.41-1.41l5.5-5.5"/></svg></th>` : "")}
              <th style="width:160px;">Actions</th>
            </tr>
          </thead>
          <tbody>
            ${safeHtml(!rows.length
              ? `<tr><td colspan="${visibleFields.length + (hasFileUpload ? 4 : 3)}" style="text-align:center;color:var(--text3);padding:32px;">No submissions found</td></tr>`
              : rows.map(item => {
                  const f = item.fields || {};
                  const hasAttachment = f.Attachments === true || f.Attachments === 1 || item.hasAttachments === true;
                  // AssignedToEmail is written by the app alongside the AssignedTo person
                  // column — use it for exact email comparison rather than display name.
                  const assignedEmail = (f[CONFIG.COL_ASSIGNED_TO_EMAIL] || "").toLowerCase();
                  const assignedName  = f[CONFIG.COL_ASSIGNED_TO] || "";
                  const isMine        = !!assignedEmail && assignedEmail === currentEmail;
                  const isUnassigned  = !assignedName;
                  const assignTitle   = isUnassigned ? "Assign to me" : `Take from ${assignedName}`;
                  return html`<tr style="cursor:pointer;" data-id="${item.id}"
                    onmouseover="this.style.background='var(--surface2)'" onmouseout="this.style.background=''">
                    <td onclick="viewOrOpenSubmission('${item.id}','${_currentFormItem.id}',${isManager})" style="font-size:13px;color:var(--text2);">
                      ${assignedName ? escHtml(assignedName) : safeHtml(`<span style="color:var(--text3);">—</span>`)}
                    </td>
                    ${safeHtml(visibleFields.map(field => html`
                      <td onclick="viewOrOpenSubmission('${item.id}','${_currentFormItem.id}',${isManager})">${formatFieldValue(f[field.internalName || field.label]).slice(0, 80)}</td>
                    `).join(""))}
                    <td onclick="viewOrOpenSubmission('${item.id}','${_currentFormItem.id}',${isManager})" style="color:var(--text2);font-size:12.5px;">${formatDate(f.Modified)}</td>
                    ${safeHtml(hasFileUpload ? `<td onclick="viewOrOpenSubmission('${item.id}','${_currentFormItem.id}',${isManager})" style="text-align:center;">
                      ${hasAttachment
                        ? `<svg width="13" height="13" viewBox="0 0 16 16" fill="none" stroke="var(--accent)" stroke-width="1.5" title="Has attachment"><path d="M13.5 8.5l-5.5 5.5a4 4 0 01-5.66-5.66l6-6a2.5 2.5 0 013.54 3.54l-6.01 6a1 1 0 01-1.41-1.41l5.5-5.5"/></svg>`
                        : `<span style="color:var(--text3)">—</span>`}
                    </td>` : "")}
                    <td onclick="event.stopPropagation()">
                      <div class="flex gap-1">
                        <button class="btn btn-ghost btn-sm btn-icon" title="View" data-id="${item.id}" data-formid="${_currentFormItem.id}" data-ismanager="${isManager}" onclick="viewOrOpenSubmission(this.dataset.id,this.dataset.formid,this.dataset.ismanager==='true')">
                          <svg width="13" height="13" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.5"><circle cx="8" cy="8" r="3"/><path d="M1 8s2.5-5 7-5 7 5 7 5-2.5 5-7 5-7-5-7-5z"/></svg>
                        </button>
                        ${safeHtml(isManager && isMine ? html`
                          <button class="btn btn-ghost btn-sm btn-icon" title="Edit" data-id="${item.id}" data-listname="${_currentFormItem.fields?.[CONFIG.COL_LISTNAME]||""}" data-formid="${_currentFormItem.id}" onclick="editSubmission(this.dataset.id,this.dataset.listname,this.dataset.formid)">
                            <svg width="13" height="13" viewBox="0 0 13 13" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M9.5 1.5l2 2-8 8H1.5v-2l8-8z"/></svg>
                          </button>
                          <button class="btn btn-ghost btn-sm btn-icon" title="Delete" data-id="${item.id}" onclick="deleteSubmission(this.dataset.id)" style="color:var(--red)">
                            <svg width="13" height="13" viewBox="0 0 13 13" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M2 3h9M5 3V2h3v1M4 3v7h5V3"/></svg>
                          </button>
                        ` : "")}
                        ${safeHtml(isManager && !isMine ? html`
                          <button class="btn btn-secondary btn-sm" title="${escHtml(assignTitle)}" data-id="${item.id}" onclick="assignSubmissionToMe(this.dataset.id)">
                            Assign to me
                          </button>
                        ` : "")}
                      </div>
                    </td>
                  </tr>`;
                }).join("")
            )}
          </tbody>
        </table>
      </div>

      ${safeHtml(total > _subPageSize ? html`
        <div style="padding:12px 16px;border-top:1px solid var(--border);display:flex;align-items:center;justify-content:space-between;">
          <span style="font-size:12.5px;color:var(--text2);">Showing ${start+1}–${Math.min(start+_subPageSize,total)} of ${total}</span>
          <div style="display:flex;gap:4px;">
            <button class="btn btn-ghost btn-sm" data-p="${_subPage-1}" onclick="_subPage=+this.dataset.p;renderSubmissionsTable(document.getElementById('main-content'))" ${_subPage<=1?"disabled":""}>←</button>
            ${safeHtml(Array.from({length:totalPages},(_,i)=>i+1).filter(p=>p===1||p===totalPages||Math.abs(p-_subPage)<=1).map((p,idx,arr)=>
              `${idx>0&&arr[idx-1]<p-1?`<span style="padding:0 4px;color:var(--text3)">…</span>`:""}`+
              html`<button class="btn btn-sm ${p===_subPage?"btn-primary":"btn-ghost"}" data-p="${p}" onclick="_subPage=+this.dataset.p;renderSubmissionsTable(document.getElementById('main-content'))">${p}</button>`
            ).join(""))}
            <button class="btn btn-ghost btn-sm" data-p="${_subPage+1}" onclick="_subPage=+this.dataset.p;renderSubmissionsTable(document.getElementById('main-content'))" ${_subPage>=totalPages?"disabled":""}>→</button>
          </div>
        </div>
      ` : "")}
    </div>
  `;
}

// =============================================================
// VIEW SUBMISSION
// =============================================================
// Routes a submission click based on role.
// Managers open the full live form (so the Complete button is available).
// Submitters get the read-only modal.
function viewOrOpenSubmission(submissionId, formItemId, isManager) {
  if (isManager && formItemId) {
    openLiveForm(formItemId, submissionId);
  } else {
    viewSubmission(submissionId);
  }
}

function viewSubmission(submissionId) {
  const item = _submissions.find(i => i.id === submissionId);
  if (!item) return;
  const f = item.fields || {};
  const allFields = (_currentFormDef?.sections || []).flatMap(s => s.fields || [])
    .filter(field => field.type !== "InfoText");
  const hasFileUpload = allFields.some(field => field.type === "FileUpload");
  const listName = _currentFormItem?.fields?.[CONFIG.COL_LISTNAME] || _currentFormDef?.listName || "";

  // Edit button is only shown when the row is currently assigned to the signed-in user.
  // Use AssignedToEmail (plain text column written by the app) for exact email comparison.
  const currentEmail  = (AppState.currentUser?.email || "").toLowerCase();
  const assignedEmail = (f[CONFIG.COL_ASSIGNED_TO_EMAIL] || "").toLowerCase();
  const isMine        = !!assignedEmail && assignedEmail === currentEmail;

  openModal(html`
    <div class="modal-header">
      <span class="modal-title">Submission — ${formatDate(f.Modified)}</span>
      <button class="btn btn-ghost btn-icon btn-sm" onclick="closeModal()">
        <svg width="14" height="14" viewBox="0 0 14 14" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M1 1l12 12M13 1L1 13"/></svg>
      </button>
    </div>
    <div class="modal-body" style="display:flex;flex-direction:column;gap:14px;">
      ${safeHtml(allFields.filter(field => field.type !== "FileUpload").map(field => {
        const val = f[field.internalName || field.label];
        const display = formatFieldValue(val);
        return html`
          <div>
            <div style="font-size:11px;text-transform:uppercase;letter-spacing:0.06em;color:var(--text3);margin-bottom:3px;">${field.label}</div>
            <div style="font-size:13.5px;color:var(--text);white-space:pre-wrap;">${display}</div>
          </div>
        `;
      }).join(""))}
      ${safeHtml(hasFileUpload ? html`
        <div>
          <div style="font-size:11px;text-transform:uppercase;letter-spacing:0.06em;color:var(--text3);margin-bottom:6px;">Attachment</div>
          <div id="submission-attachments-${submissionId}">
            <span class="spinner"></span>
          </div>
        </div>
      ` : "")}
    </div>
    <div class="modal-footer">
      <button class="btn btn-ghost" onclick="closeModal()">Close</button>
      ${safeHtml(isMine ? html`
        <button class="btn btn-secondary" data-id="${submissionId}"
          data-listname="${listName}"
          data-formid="${_currentFormItem.id}"
          onclick="closeModal();editSubmission(this.dataset.id,this.dataset.listname,this.dataset.formid)">
          Edit
        </button>
      ` : "")}
    </div>
  `, false, true);

  // Load attachments asynchronously if form has a FileUpload field
  if (hasFileUpload) {
    loadSubmissionAttachments(submissionId, listName);
  }
}

async function loadSubmissionAttachments(submissionId, listName) {
  const el = document.getElementById(`submission-attachments-${submissionId}`);
  if (!el) return;
  try {
    const token = await getSpToken();
    const siteUrl = CONFIG.SITE_URL.replace(/\/$/, "");
    // Origin only — ServerRelativeUrl already contains the full path from root
    const spOrigin = new URL(CONFIG.SITE_URL).origin;
    // Need SP numeric item ID
    const spId = await getSpItemId(listName, submissionId);
    const url = `${siteUrl}/_api/web/lists/GetByTitle('${encodeURIComponent(listName)}')/items(${spId})/AttachmentFiles`;
    const res = await fetch(url, {
      headers: { Authorization: "Bearer " + token, Accept: "application/json;odata=verbose" }
    });
    const data = await res.json();
    const files = data?.d?.results || [];
    if (!files.length) {
      el.innerHTML = `<span style="font-size:13px;color:var(--text3);">No files attached</span>`;
      return;
    }
    el.innerHTML = files.map(f => html`
      <div style="display:inline-flex;align-items:center;gap:6px;font-size:13px;">
        <svg width="13" height="13" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M14 10v3a1 1 0 01-1 1H3a1 1 0 01-1-1v-3M8 2v8M5 9l3 3 3-3"/></svg>
        <a href="#" style="color:var(--accent);text-decoration:none;"
          data-url="${spOrigin}${f.ServerRelativeUrl}"
          data-filename="${f.FileName}"
          onclick="event.preventDefault();openAttachment(this.dataset.url, this.dataset.filename)">
          ${f.FileName}
        </a>
      </div>
    `).join("");
  } catch (e) {
    const el2 = document.getElementById(`submission-attachments-${submissionId}`);
    if (el2) el2.innerHTML = `<span style="font-size:12px;color:var(--red);">Could not load attachment: ${e.message}</span>`;
  }
}

async function openAttachment(url, filename) {
  try {
    showToast("info", "Opening attachment…");
    const token = await getSpToken();
    // Use SharePoint REST GetFileByServerRelativeUrl to avoid redirect issues
    const spOrigin = new URL(CONFIG.SITE_URL).origin;
    const serverRelativeUrl = url.replace(spOrigin, "");
    const restUrl = `${CONFIG.SITE_URL.replace(/\/$/, "")}/_api/web/GetFileByServerRelativeUrl('${encodeURIComponent(serverRelativeUrl)}')/$value`;
    const res = await fetch(restUrl, {
      headers: { Authorization: "Bearer " + token }
    });
    if (!res.ok) throw new Error(`${res.status} ${res.statusText}`);
    const blob = await res.blob();
    const blobUrl = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = blobUrl;
    a.download = filename;
    a.click();
    setTimeout(() => URL.revokeObjectURL(blobUrl), 10000);
  } catch (e) {
    showToast("error", `Could not open attachment: ${e.message}`);
  }
}

// =============================================================
// COPY LINK
// =============================================================
function copyFormLink(formItemId) {
  const url = `${window.location.origin}${window.location.pathname}?view=my-forms&formId=${formItemId}`;
  navigator.clipboard.writeText(url).then(() => {
    showToast("success", "Link copied to clipboard");
  }).catch(() => {
    // Fallback for browsers without clipboard API
    const input = document.createElement("input");
    input.value = url;
    document.body.appendChild(input);
    input.select();
    document.execCommand("copy");
    document.body.removeChild(input);
    showToast("success", "Link copied to clipboard");
  });
}

// =============================================================
// MANAGE MANAGERS
// =============================================================
async function openManageManagers(formItemId) {
  openModal(html`
    <div class="modal-header">
      <span class="modal-title">Manage Form Managers</span>
      <button class="btn btn-ghost btn-icon btn-sm" onclick="closeModal()">
        <svg width="14" height="14" viewBox="0 0 14 14" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M1 1l12 12M13 1L1 13"/></svg>
      </button>
    </div>
    <div class="modal-body">
      <p style="font-size:13px;color:var(--text2);margin-bottom:16px;">
        Form Managers can view and manage all submissions. Add or remove access below.
      </p>

      <div id="managers-current-list">
        <div style="text-align:center;padding:16px;"><span class="spinner"></span></div>
      </div>

      <div style="margin-top:16px;border-top:1px solid var(--border);padding-top:16px;">
        <label style="font-weight:600;font-size:13px;display:block;margin-bottom:8px;">Add a Manager</label>
        <div class="flex gap-2">
          <input id="manager-add-search" class="input" placeholder="Search by name or email…"
            oninput="debouncedManagerAddSearch(this.value)">
          <button class="btn btn-secondary" onclick="searchManagerAdd()">Search</button>
        </div>
        <div id="manager-add-results" style="margin-top:8px;"></div>
      </div>
    </div>
    <div class="modal-footer">
      <button class="btn btn-ghost" onclick="closeModal()">Close</button>
    </div>
  `, false, true);

  await loadCurrentManagers(formItemId);
}

async function loadCurrentManagers(formItemId) {
  const el = document.getElementById("managers-current-list");
  if (!el) return;

  try {
    const listName = _currentFormItem?.fields?.[CONFIG.COL_LISTNAME] || _currentFormDef?.listName;
    const token = await getSpToken();
    const siteUrl = CONFIG.SITE_URL.replace(/\/$/, "");

    // Fetch the list GUID via Graph
    const siteId = await getSiteId();
    const listId = await getListId(listName);

    // Read all role assignments on the list via SP REST
    const url = `${siteUrl}/_api/web/lists(guid'${listId}')/roleassignments?$expand=Member,RoleDefinitionBindings`;
    const res = await fetch(url, {
      headers: { Authorization: "Bearer " + token, Accept: "application/json;odata=verbose" }
    });
    const data = await res.json();
    const assignments = data?.d?.results || [];

    // Get contribute role name for comparison
    const contributeNames = ["Contribute", "Edit"];

    // Find individual users (not groups) with Contribute — these are managers and the author
    const authorEmail = (_currentFormItem?.createdBy?.user?.email || "").toLowerCase();
    const managers = assignments.filter(a => {
      const isUser = a.Member?.PrincipalType === 1; // 1 = User, 8 = SharePoint Group
      const hasContribute = a.RoleDefinitionBindings?.results?.some(r => contributeNames.includes(r.Name));
      const email = (a.Member?.Email || a.Member?.LoginName || "").toLowerCase();
      const isAuthor = email.includes(authorEmail) || authorEmail.includes(email.split("|").pop());
      return isUser && hasContribute && !isAuthor;
    });

    if (!managers.length) {
      el.innerHTML = `<p style="font-size:13px;color:var(--text3);">No managers added yet.</p>`;
      return;
    }

    el.innerHTML = html`
      <label style="font-weight:600;font-size:13px;display:block;margin-bottom:8px;">Current Managers</label>
      <div style="display:flex;flex-direction:column;gap:6px;">
        ${safeHtml(managers.map(a => {
          const name = a.Member?.Title || a.Member?.LoginName || "Unknown";
          const email = a.Member?.Email || "";
          const principalId = a.Member?.Id;
          const initials = name.split(" ").map(n => n[0]).join("").slice(0, 2).toUpperCase();
          return html`
            <div class="flex items-center gap-3" style="padding:8px 12px;background:var(--bg3);border:1px solid var(--border);border-radius:var(--radius-sm);">
              <div class="avatar" style="width:28px;height:28px;font-size:11px;flex-shrink:0;">${initials}</div>
              <div style="flex:1;min-width:0;">
                <div style="font-size:13.5px;font-weight:500;">${name}</div>
                <div style="font-size:11.5px;color:var(--text3);">${email}</div>
              </div>
              <button class="btn btn-ghost btn-sm btn-icon" style="color:var(--red);"
                data-principalid="${principalId}" data-name="${name}" data-email="${email}"
                data-formid="${formItemId}"
                onclick="removeFormManager(this)">
                <svg width="13" height="13" viewBox="0 0 13 13" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M1 1l11 11M12 1L1 12"/></svg>
              </button>
            </div>
          `;
        }).join(""))}
      </div>
    `;
  } catch (e) {
    const el2 = document.getElementById("managers-current-list");
    if (el2) el2.innerHTML = html`<p style="color:var(--red);font-size:13px;">Error loading managers: ${e.message}</p>`;
  }
}

// Manager add search — uses the shared createPeopleSearch utility.
// addFormManagerFromEl reads _currentFormItem directly so no extra
// data-formid attribute is needed on the result rows.
const _managerAddSearch = createPeopleSearch({
  inputId:   "manager-add-search",
  resultsId: "manager-add-results",
  onClickFn: "addFormManagerFromEl",
});
function debouncedManagerAddSearch(val) { _managerAddSearch.debounced(val); }
async function searchManagerAdd(q)      { await _managerAddSearch.search(q); }

async function addFormManagerFromEl(el) {
  const email      = el.dataset.email;
  const name       = el.dataset.name;
  // formItemId comes from _currentFormItem — not a data attribute — because
  // createPeopleSearch no longer injects data-formid on result rows.
  const formItemId = _currentFormItem?.id;
  const listName   = _currentFormItem?.fields?.[CONFIG.COL_LISTNAME] || _currentFormDef?.listName;

  el.style.opacity = "0.5";
  el.style.pointerEvents = "none";

  try {
    const siteId = await getSiteId();
    const listId = await getListId(listName);
    const listBase = `/_api/web/lists(guid'${listId}')`;

    // Get contribute role ID
    const rolesData = await spGet(`/_api/web/roledefinitions`);
    const roles = rolesData?.d?.results || [];
    const contributeRole = roles.find(r => r.Name === "Contribute") || roles.find(r => r.Name === "Edit");
    if (!contributeRole) throw new Error("Contribute role not found");

    // Ensure user and grant Contribute
    const userData = await spPost(`/_api/web/ensureuser`, { logonName: `i:0#.f|membership|${email}` });
    const spUserId = userData?.d?.Id;
    if (!spUserId) throw new Error("Could not resolve user");
    await spPost(`${listBase}/roleassignments/addroleassignment(principalid=${spUserId},roledefid=${contributeRole.Id})`);

    // Sync to form definition JSON
    await syncManagersToDefinition(formItemId, { action: "add", email, displayName: name, id: el.dataset.id });

    showToast("success", `${name} added as Form Manager`);

    // Clear search and reload manager list
    const searchEl = document.getElementById("manager-add-search");
    const resultsEl = document.getElementById("manager-add-results");
    if (searchEl) searchEl.value = "";
    if (resultsEl) resultsEl.innerHTML = "";
    await loadCurrentManagers(formItemId);
  } catch (e) {
    showToast("error", `Could not add manager: ${e.message}`);
    el.style.opacity = "";
    el.style.pointerEvents = "";
  }
}

async function removeFormManager(btn) {
  const principalId = btn.dataset.principalid;
  const name        = btn.dataset.name;
  const email       = btn.dataset.email;
  const formItemId  = btn.dataset.formid;
  const listName    = _currentFormItem?.fields?.[CONFIG.COL_LISTNAME] || _currentFormDef?.listName;

  btn.disabled = true;

  try {
    const listId = await getListId(listName);
    const listBase = `/_api/web/lists(guid'${listId}')`;
    await spDelete(`${listBase}/roleassignments/getbyprincipalid(${principalId})`);

    // Sync removal to form definition JSON
    await syncManagersToDefinition(formItemId, { action: "remove", email });

    showToast("success", `${name} removed`);
    await loadCurrentManagers(formItemId);
  } catch (e) {
    showToast("error", `Could not remove manager: ${e.message}`);
    btn.disabled = false;
  }
}

async function syncManagersToDefinition(formItemId, change) {
  // Keep formManagers array in definition JSON in sync with list permissions
  try {
    if (!_currentFormDef) return;
    let managers = [...(_currentFormDef.formManagers || [])];
    if (change.action === "add") {
      if (!managers.find(m => (m.email || "").toLowerCase() === change.email.toLowerCase())) {
        managers.push({ id: change.id || "", displayName: change.displayName, email: change.email });
      }
    } else if (change.action === "remove") {
      managers = managers.filter(m => (m.email || "").toLowerCase() !== change.email.toLowerCase());
    }
    _currentFormDef.formManagers = managers;
    await uploadJsonAttachment(CONFIG.FORMS_LIST, formItemId, "form-definition.json", _currentFormDef);
  } catch (e) {
    console.warn("Could not sync managers to definition:", e.message);
  }
}

// =============================================================
// ASSIGN TO ME — soft check-out for a single submission row
// Writes the current user into AssignedTo, then re-fetches the row from
// SharePoint to confirm who actually owns it after the write. If two users
// click "Assign to me" within the same second, last-write-wins on the server;
// the loser's re-fetch shows the winner so they don't think they own it.
// =============================================================
async function assignSubmissionToMe(submissionId) {
  const listName = _currentFormItem?.fields?.[CONFIG.COL_LISTNAME] || _currentFormDef?.listName;
  if (!listName) { showToast("error", "Cannot determine list name"); return; }

  const email = AppState.currentUser?.email;
  if (!email) { showToast("error", "Could not identify current user"); return; }

  try {
    // SP integer user ID is cached at boot in AppState._currentUserSpId.
    // Fall back to resolveSpUserId for older sessions or if the boot fetch failed.
    let spUserId = AppState._currentUserSpId;
    if (!spUserId) spUserId = await resolveSpUserId(email);
    if (!spUserId) throw new Error("Could not resolve current user to a SharePoint ID");

    // Person columns are written via Graph as <ColumnName>LookupId with the integer ID.
    // AssignedToEmail is written alongside as a plain text column so isMine checks
    // can compare by email — Graph does not return Email on programmatic Person columns.
    await updateListItem(listName, submissionId, {
      [`${CONFIG.COL_ASSIGNED_TO}LookupId`]: spUserId,
      [CONFIG.COL_ASSIGNED_TO_EMAIL]:        email,
    });

    // Update local state optimistically — SP replication lag means a re-fetch
    // immediately after the PATCH often returns the old value. Trust the PATCH.
    const idx = _submissions.findIndex(i => i.id === submissionId);
    if (idx !== -1) {
      _submissions[idx].fields = _submissions[idx].fields || {};
      _submissions[idx].fields[CONFIG.COL_ASSIGNED_TO]       = AppState.currentUser?.displayName || email;
      _submissions[idx].fields[CONFIG.COL_ASSIGNED_TO_EMAIL] = email;
    }

    showToast("success", "Assigned to you");
    renderSubmissionsTable(document.getElementById("main-content"));
  } catch (e) {
    showToast("error", "Could not assign: " + e.message);
  }
}

function editSubmission(submissionId, listName, formItemId) {
  // Reuse the existing openLiveForm with editItemId so the live form
  // renders pre-populated with existing values and submits as an update
  openLiveForm(formItemId, submissionId);
}

// =============================================================
// SOFT DELETE SUBMISSION
// =============================================================
function deleteSubmission(submissionId) {
  openModal(html`
    <div class="modal-header">
      <span class="modal-title">Delete Submission</span>
      <button class="btn btn-ghost btn-icon btn-sm" onclick="closeModal()">
        <svg width="14" height="14" viewBox="0 0 14 14" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M1 1l12 12M13 1L1 13"/></svg>
      </button>
    </div>
    <div class="modal-body">
      <p style="color:var(--text2);">This will mark the submission as deleted. It won't be visible in the submissions list but remains in SharePoint for audit purposes.</p>
    </div>
    <div class="modal-footer">
      <button class="btn btn-ghost" onclick="closeModal()">Cancel</button>
      <button class="btn btn-danger" data-id="${submissionId}" onclick="closeModal();doDeleteSubmission(this.dataset.id)">Delete</button>
    </div>
  `);
}

async function doDeleteSubmission(submissionId) {
  const listName = _currentFormItem?.fields?.[CONFIG.COL_LISTNAME] || _currentFormDef?.listName;
  if (!listName) { showToast("error", "Cannot determine list name"); return; }
  try {
    await updateListItem(listName, submissionId, { [CONFIG.COL_IS_DELETED]: true });
    // Remove from in-memory list and re-render without reloading
    _submissions = _submissions.filter(i => i.id !== submissionId);
    showToast("success", "Submission deleted");
    renderSubmissionsTable(document.getElementById("main-content"));
  } catch (e) {
    showToast("error", "Delete failed: " + e.message);
  }
}

// =============================================================
// EXPORT — Excel bulk export for a form's submissions
// =============================================================
async function exportFormExcel(formItemId) {
  const listName = _currentFormItem?.fields?.[CONFIG.COL_LISTNAME] || _currentFormDef?.listName;
  if (!listName) { showToast("error", "No data list found"); return; }

  showToast("info", "Preparing export…");
  try {
    const allFields = (_currentFormDef?.sections || []).flatMap(s => s.fields || [])
      .filter(f => f.type !== "InfoText");

    // Build CSV (simple, no dependency)
    const headers = [...allFields.map(f => f.label), "Submitted", "Modified"];
    const csvRows = [headers.map(h => `"${h}"`).join(",")];

    for (const item of _submissions) {
      const f = item.fields || {};
      const row = [
        ...allFields.map(field => {
          const val = f[field.internalName || field.label];
          return `"${String(val == null ? "" : val).replace(/"/g, '""')}"`;
        }),
        `"${formatDate(f.Created)}"`,
        `"${formatDate(f.Modified)}"`,
      ];
      csvRows.push(row.join(","));
    }

    const csv = csvRows.join("\n");
    const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
    const url  = URL.createObjectURL(blob);
    const a    = document.createElement("a");
    a.href     = url;
    a.download = `${listName}-submissions.csv`;
    a.click();
    URL.revokeObjectURL(url);
    showToast("success", "Export downloaded");
  } catch (e) {
    showToast("error", "Export failed: " + e.message);
  }
}