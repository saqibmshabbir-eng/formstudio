// Form Studio — Admin: Review & Approval

// =============================================================
// ADMIN REVIEW
// All form lifecycle stages live in CONFIG.FORMS_LIST.
// Status flow: Draft → Submitted → Approved for Preview → Approved/Rejected
// Approved items get Status set to "Preview" or "Live" after SP list creation.
// =============================================================
// =============================================================
// COLUMN NAME SANITISATION
// SharePoint internal column name rules:
//   - ASCII alphanumeric and underscore only
//   - Cannot start with a digit or underscore
//   - Max 32 characters
//   - No reserved names (Title, ID, etc. — SharePoint rejects duplicates)
// =============================================================
function sanitiseColumnName(label) {
  return (label || "Field")
    .normalize("NFD")                          // decompose accented chars (é → e + ́)
    .replace(/[\u0300-\u036f]/g, "")          // strip accent marks
    .replace(/[^a-zA-Z0-9]/g, "")             // strip everything non-alphanumeric
    .replace(/^[0-9]+/, "")                   // strip leading digits
    || "Field";                                // fallback if entirely stripped
}

async function renderAdminReview(container) {
  // Authors with form request access can see this view too — not just admins
  if (!AppState.isAdmin && !AppState.hasFormRequestAccess) {
    container.innerHTML = `<div class="empty-state"><h3>Access Denied</h3><p>You don't have permission to view form requests.</p></div>`;
    return;
  }

  const isAdmin = AppState.isAdmin;
  const currentEmail = (AppState.currentUser?.email || "").toLowerCase();

  container.innerHTML = html`
    <div class="flex items-center justify-between mb-4">
      <div>
        <h1 style="font-size:22px;font-weight:600;letter-spacing:-0.02em;">Form Requests</h1>
        <p style="color:var(--text2);font-size:13.5px;margin-top:2px;">${isAdmin ? "Approve, preview, or reject submitted form definitions" : "Manage and submit your form requests"}</p>
      </div>
    </div>
    <div class="card" id="admin-table">
      <div style="padding:40px;text-align:center;"><span class="spinner"></span></div>
    </div>
  `;

  try {
    const items = await getListItems(CONFIG.FORMS_LIST);

    // Retro forms never appear in admin/author workflows
    const nonRetro = items.filter(i => !i.fields?.[CONFIG.COL_RETRO]);

    // Admins see everything except Created drafts from other authors.
    // Authors see only their own forms (all statuses).
    const visibleItems = nonRetro.filter(i => {
      const s = i.fields?.[CONFIG.COL_STATUS] || "";
      const createdByEmail = (i.createdBy?.user?.email || "").toLowerCase();
      const isOwn = createdByEmail === currentEmail;
      if (isAdmin) return s !== "Created" || isOwn;
      return isOwn; // authors only see their own
    });

    AppState.allRequests = visibleItems;
    const card = document.getElementById("admin-table");

    if (!visibleItems.length) {
      card.innerHTML = `<div class="empty-state"><h3>${isAdmin ? "No forms yet" : "You have no form requests yet"}</h3></div>`;
      return;
    }

    card.innerHTML = html`
      <div class="table-wrap">
        <table>
          <thead><tr><th>Form</th><th>Creator</th><th>Status</th><th>Modified</th><th>Actions</th></tr></thead>
          <tbody>
            ${safeHtml(visibleItems.map(item => {
              const f = item.fields || {};
              const status = f[CONFIG.COL_STATUS] || "Created";
              const author = item.createdBy?.user?.displayName || "—";
              const createdByEmail = (item.createdBy?.user?.email || "").toLowerCase();
              const isOwn = createdByEmail === currentEmail;

              // Admin approval actions from STATUS_FLOW
              const adminActions = isAdmin ? (CONFIG.STATUS_FLOW[status] || []) : [];

              const isCreated   = status === "Created";
              const isSubmitted = status === "Submitted";
              const isPreview   = status === "Preview";
              const isLocked    = ["Approved","Live","Rejected"].includes(status);

              // Authors: Submit + Edit + Delete on Created only. Recall on Submitted only.
              const canSubmit     = isOwn && isCreated;
              const authorCanEdit = isOwn && isCreated;
              const canRecall     = isOwn && isSubmitted;
              const canDelete     = isOwn && isCreated;
              // Admins: Edit on Submitted and Preview (Preview triggers warning)
              const adminCanEdit  = isAdmin && (isSubmitted || isPreview);

              const canEdit    = authorCanEdit || adminCanEdit;
              const showLocked = isLocked && !adminActions.length;

              return html`<tr>
                <td>
                  <strong>${f.Title||"—"}</strong>
                  ${safeHtml(f[CONFIG.COL_COMMENTS] ? html`<div style="font-size:12px;color:var(--text3);margin-top:3px;font-style:italic;">${f[CONFIG.COL_COMMENTS]}</div>` : "")}
                </td>
                <td style="color:var(--text2);font-size:13px;">${author}</td>
                <td>${safeHtml(statusBadge(status))}</td>
                <td style="color:var(--text2);font-size:12.5px;">${formatDate(f.Modified)}</td>
                <td>
                  <div class="flex gap-2" style="flex-wrap:wrap;gap:6px;">
                    <button class="btn btn-sm btn-ghost" data-id="${item.id}" data-title="${f.Title||""}" onclick="adminPreviewForm(this.dataset.id, this.dataset.title)">
                      <svg width="12" height="12" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.5"><circle cx="8" cy="8" r="3"/><path d="M1 8s2.5-5 7-5 7 5 7 5-2.5 5-7 5-7-5-7-5z"/></svg>
                      View
                    </button>
                    ${safeHtml(canSubmit ? html`
                      <button class="btn btn-sm btn-primary" data-id="${item.id}" onclick="submitRequest(this.dataset.id)">
                        Submit
                      </button>
                    ` : "")}
                    ${safeHtml(canRecall ? html`
                      <button class="btn btn-sm btn-secondary" data-id="${item.id}" onclick="recallRequest(this.dataset.id)">
                        Recall
                      </button>
                    ` : "")}
                    ${safeHtml(canEdit ? html`
                      <button class="btn btn-sm btn-secondary" data-id="${item.id}" data-preview="${isPreview}" data-created="${isCreated}" onclick="handleEditRequest(this)">
                        <svg width="12" height="12" viewBox="0 0 13 13" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M9.5 1.5l2 2-8 8H1.5v-2l8-8z"/></svg>
                        Edit
                      </button>
                    ` : "")}
                    ${safeHtml(canDelete ? html`
                      <button class="btn btn-sm btn-danger" data-id="${item.id}" data-title="${f.Title||"Untitled"}" onclick="deleteFormRequest(this.dataset.id, this.dataset.title)">
                        <svg width="12" height="12" viewBox="0 0 13 13" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M2 3h9M5 3V2h3v1M4 3v7h5V3"/></svg>
                        Delete
                      </button>
                    ` : "")}
                    ${safeHtml(adminActions.map(a => html`
                      <button class="btn btn-sm ${a==="Reject"?"btn-danger":"btn-secondary"}"
                        data-id="${item.id}" data-action="${a}" onclick="adminChangeStatus(this.dataset.id, this.dataset.action)">
                        ${a}
                      </button>
                    `).join(""))}
                    ${safeHtml(showLocked ? `<span style="font-size:12px;color:var(--text3);">Locked</span>` : "")}
                  </div>
                </td>
              </tr>`;
            }).join(""))}
          </tbody>
        </table>
      </div>
    `;
  } catch (e) {
    document.getElementById("admin-table").innerHTML = html`<div class="empty-state"><p style="color:var(--red)">Error: ${e.message}</p></div>`;
  }
}

// =============================================================
// ADMIN FORM PREVIEW
// =============================================================
async function adminPreviewForm(itemId, title) {
  openModal(`
    <div class="modal-header" style="border-bottom:1px solid var(--border);">
      <div style="display:flex;align-items:center;gap:10px;flex:1;min-width:0;">
        <span class="modal-title" style="font-size:15px;">${escHtml(title || "Form Preview")}</span>
        <span class="badge badge-blue" style="flex-shrink:0;">Preview</span>
      </div>
      <button class="btn btn-ghost btn-icon btn-sm" onclick="closeModal()">
        <svg width="14" height="14" viewBox="0 0 14 14" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M1 1l12 12M13 1L1 13"/></svg>
      </button>
    </div>
    <div class="modal-body" id="admin-preview-body" style="padding:32px;">
      <div style="text-align:center;padding:40px 0;"><span class="spinner"></span><p style="margin-top:12px;color:var(--text2);font-size:13px;">Loading form definition…</p></div>
    </div>
    <div class="modal-footer" id="admin-preview-footer" style="justify-content:flex-end;gap:8px;">
      <button class="btn btn-ghost" onclick="closeModal()">Close</button>
    </div>
  `, false, true);

  try {
    const def = await getFormDefinition(CONFIG.FORMS_LIST, itemId);
    if (!def) throw new Error("Form definition not found — ensure the FormDefinition column exists on the Forms list.");

    const body = document.getElementById("admin-preview-body");
    const footer = document.getElementById("admin-preview-footer");
    if (!body) return;

    const item = AppState.allRequests?.find(r => r.id === itemId);
    const status = item?.fields?.[CONFIG.COL_STATUS] || "";
    const actions = CONFIG.STATUS_FLOW[status] || [];

    const { title: formTitle, sections = [], layout = "single" } = def;

    const previewHtml = sections.length ? `
      <div class="form-preview" style="box-shadow:none;border:none;padding:0;">
        <h2 style="font-size:20px;font-weight:700;letter-spacing:-0.02em;margin-bottom:6px;">${escHtml(formTitle || title)}</h2>
        ${layout === "multistep" ? `<p style="color:var(--text2);font-size:13px;margin-bottom:4px;">Multi-step form — ${sections.length} step${sections.length!==1?"s":""}</p>` : ""}
        <div style="margin-top:20px;">
          ${sections.map(sec => `
            <div class="preview-section">
              ${sec.title ? `<div class="preview-section-title">${escHtml(sec.title)}</div>` : ""}
              ${(sec.fields || []).map(field => renderPreviewField(field)).join("")}
            </div>
          `).join("")}
        </div>
        ${layout === "multistep" ? `
          <div class="flex gap-2" style="margin-top:16px;padding-top:16px;border-top:1px solid var(--border);">
            <button class="btn btn-secondary" disabled>← Previous</button>
            <button class="btn btn-primary" disabled>Next →</button>
          </div>
        ` : `<button class="btn btn-primary" style="margin-top:16px;" disabled>Submit</button>`}
      </div>
    ` : `<div class="empty-state" style="padding:40px 0;"><p style="color:var(--text2);">No sections defined yet.</p></div>`;

    const comment = item?.fields?.[CONFIG.COL_COMMENTS] || "";
    const allFields = sections.flatMap(s => s.fields || []).filter(f => f.type !== "InfoText");
    const summaryHtml = `
      <div style="background:var(--bg3);border:1px solid var(--border);border-radius:var(--radius-sm);padding:14px 16px;margin-bottom:20px;display:flex;flex-wrap:wrap;gap:16px 32px;">
        <div><span style="font-size:11px;text-transform:uppercase;letter-spacing:0.06em;color:var(--text3);">Sections</span><br><strong>${sections.length}</strong></div>
        <div><span style="font-size:11px;text-transform:uppercase;letter-spacing:0.06em;color:var(--text3);">Fields</span><br><strong>${allFields.length}</strong></div>
        <div><span style="font-size:11px;text-transform:uppercase;letter-spacing:0.06em;color:var(--text3);">Layout</span><br><strong>${layout === "multistep" ? "Multi-step" : "Single page"}</strong></div>
        <div><span style="font-size:11px;text-transform:uppercase;letter-spacing:0.06em;color:var(--text3);">Submission</span><br><strong>${def.submissionType === "SubmitEdit" ? "Submit & Edit" : "Submit Only"}</strong></div>
        <div><span style="font-size:11px;text-transform:uppercase;letter-spacing:0.06em;color:var(--text3);">Access</span><br><strong>${CONFIG.ACCESS_OPTIONS.find(a=>a.value===def.access)?.label || def.access || "—"}</strong></div>
      </div>
      ${comment ? `
        <div style="background:rgba(0,33,71,0.04);border:1px solid var(--border);border-left:3px solid var(--accent);border-radius:var(--radius-sm);padding:12px 16px;margin-bottom:20px;">
          <div style="font-size:11px;text-transform:uppercase;letter-spacing:0.06em;color:var(--text3);margin-bottom:6px;">Admin Comment</div>
          <div style="font-size:13.5px;color:var(--text);line-height:1.6;">${escHtml(comment)}</div>
        </div>
      ` : ""}
    `;

    body.innerHTML = summaryHtml + previewHtml;

    if (footer && actions.length) {
      footer.innerHTML = html`
        <button class="btn btn-ghost" onclick="closeModal()">Close</button>
        ${safeHtml(actions.map(a => html`
          <button class="btn btn-sm ${a==="Reject"?"btn-danger":"btn-primary"}"
            data-id="${itemId}" data-action="${a}" onclick="closeModal();adminChangeStatus(this.dataset.id, this.dataset.action)">
            ${a}
          </button>
        `).join(""))}
      `;
    }
  } catch (e) {
    const body = document.getElementById("admin-preview-body");
    if (body) body.innerHTML = `<div class="empty-state" style="padding:40px 0;"><p style="color:var(--red);">Failed to load preview: ${escHtml(e.message)}</p></div>`;
  }
}

// =============================================================
// EDIT A PREVIEW-STAGE FORM
// Deletes the item entirely, opens the definition in a fresh builder
// =============================================================
function handleEditRequest(el) {
  const id = el.dataset.id;
  const isPreview = el.dataset.preview === "true";
  const isCreated = el.dataset.created === "true";
  if (isPreview) editPreviewFormRequest(id);
  else if (isCreated) doEditFormRequest(id); // Never been submitted — no warning needed
  else editFormRequest(id);
}

// =============================================================
// DELETE FORM REQUEST (author, Created state only)
// =============================================================
function deleteFormRequest(itemId, title) {
  openModal(html`
    <div class="modal-header">
      <span class="modal-title">Delete Form Request</span>
      <button class="btn btn-ghost btn-icon btn-sm" onclick="closeModal()">
        <svg width="14" height="14" viewBox="0 0 14 14" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M1 1l12 12M13 1L1 13"/></svg>
      </button>
    </div>
    <div class="modal-body">
      <p style="color:var(--text2);">Are you sure you want to delete <strong>${escHtml(title)}</strong>?</p>
      <p style="color:var(--text2);margin-top:8px;">This will permanently remove the form request. Because it has not been approved or provisioned, no data or SharePoint list will be affected.</p>
    </div>
    <div class="modal-footer">
      <button class="btn btn-ghost" onclick="closeModal()">Cancel</button>
      <button class="btn btn-danger" data-id="${itemId}" onclick="closeModal();doDeleteFormRequest(this.dataset.id)">Delete</button>
    </div>
  `);
}

async function doDeleteFormRequest(itemId) {
  try {
    const siteId = await getSiteId();
    const listId = await getListId(CONFIG.FORMS_LIST);
    await graphDelete(`/sites/${siteId}/lists/${listId}/items/${itemId}`);
    showToast("success", "Form request deleted");
    renderAdminReview(document.getElementById("admin-table"));
  } catch (e) {
    showToast("error", "Could not delete form request: " + e.message);
  }
}

async function editPreviewFormRequest(itemId) {
  openModal(html`
    <div class="modal-header">
      <span class="modal-title">⚠️ Warning — Form is in Preview</span>
      <button class="btn btn-ghost btn-icon btn-sm" onclick="closeModal()">
        <svg width="14" height="14" viewBox="0 0 14 14" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M1 1l12 12M13 1L1 13"/></svg>
      </button>
    </div>
    <div class="modal-body">
      <p style="font-size:14px;line-height:1.7;">This form is currently in Preview and may already have users testing it. Editing will:</p>
      <ul style="margin:12px 0 0 20px;font-size:13.5px;line-height:2.2;color:var(--text2);">
        <li><strong style="color:var(--red);">Delete the SharePoint data list and all submissions within it</strong></li>
        <li>Reset the form back to Created — the author will need to re-submit for review</li>
        <li>Require re-approval before the form goes live again</li>
      </ul>
      <div style="margin-top:16px;padding:12px 14px;background:rgba(220,38,38,0.06);border:1px solid rgba(220,38,38,0.2);border-radius:var(--radius-sm);">
        <strong style="font-size:13px;color:var(--red);">This cannot be undone.</strong>
        <span style="font-size:13px;color:var(--text2);"> Export any submitted data before continuing.</span>
      </div>
    </div>
    <div class="modal-footer">
      <button class="btn btn-ghost" onclick="closeModal()">Cancel</button>
      <button class="btn btn-danger" data-id="${itemId}" onclick="closeModal();doEditFormRequest(this.dataset.id)">I understand — Edit anyway</button>
    </div>
  `);
}

async function editFormRequest(itemId) {
  openModal(html`
    <div class="modal-header">
      <span class="modal-title">⚠️ Warning — Data Will Be Lost</span>
      <button class="btn btn-ghost btn-icon btn-sm" onclick="closeModal()">
        <svg width="14" height="14" viewBox="0 0 14 14" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M1 1l12 12M13 1L1 13"/></svg>
      </button>
    </div>
    <div class="modal-body">
      <p style="font-size:14px;line-height:1.7;">Editing this form definition will:</p>
      <ul style="margin:12px 0 0 20px;font-size:13.5px;line-height:2.2;color:var(--text2);">
        <li><strong style="color:var(--red);">Delete the SharePoint data list and all submissions within it</strong></li>
        <li>Reset the form back to Created — the author will need to re-submit for review</li>
        <li>Require re-approval before the form goes live again</li>
      </ul>
      <div style="margin-top:16px;padding:12px 14px;background:rgba(220,38,38,0.06);border:1px solid rgba(220,38,38,0.2);border-radius:var(--radius-sm);">
        <strong style="font-size:13px;color:var(--red);">This cannot be undone.</strong>
        <span style="font-size:13px;color:var(--text2);"> Export any submitted data before continuing.</span>
      </div>
    </div>
    <div class="modal-footer">
      <button class="btn btn-ghost" onclick="closeModal()">Cancel</button>
      <button class="btn btn-danger" data-id="${itemId}" onclick="closeModal();doEditFormRequest(this.dataset.id)">I understand — Edit anyway</button>
    </div>
  `);
}

async function doEditFormRequest(itemId) {
  const main = document.getElementById("main-content");
  main.innerHTML = `<div style="padding:60px;text-align:center;"><span class="spinner" style="width:32px;height:32px;border-width:3px;"></span><p style="margin-top:16px;color:var(--text2)">Preparing form for editing…</p></div>`;

  try {
    const def = await getFormDefinition(CONFIG.FORMS_LIST, itemId);
    if (!def) throw new Error("Form definition not found");

    showProgress("Cleaning up", "Deleting data list…");

    // Delete the provisioned SP data list if it exists
    const listName = def.listName;
    if (listName) {
      try {
        const siteId = await getSiteId();
        const lists  = await graphGet(`/sites/${siteId}/lists`);
        const existing = (lists.value || []).find(l => l.displayName === listName);
        if (existing) await graphDelete(`/sites/${siteId}/lists/${existing.id}`);
      } catch (e) {
        console.warn("Could not delete SP data list:", e.message);
      }
    }

    // Reset the Forms item back to Created rather than deleting it —
    // the author edits in place and re-submits. Clear ListName so there
    // is no stale reference to the deleted data list.
    updateProgress("Resetting form entry…");
    try {
      await updateListItem(CONFIG.FORMS_LIST, itemId, {
        [CONFIG.COL_STATUS]:   "Created",
        [CONFIG.COL_LISTNAME]: "",
        [CONFIG.COL_COMMENTS]: "",
      });
    } catch (e) {
      console.warn("Could not reset Forms item:", e.message);
    }

    hideProgress();

    // Load definition into the builder for editing — item already exists, update on re-submit
    resetBuilderForm();
    AppState.builderMode                    = "edit";
    AppState.builderItemId                  = itemId;
    AppState.builderForm.title              = def.title              || "";
    AppState.builderForm.listName           = def.listName           || generateListName(def.title);
    AppState.builderForm.access             = def.access             || "StaffStudents";
    AppState.builderForm.submissionType     = def.submissionType     || "Submit";
    AppState.builderForm.layout             = def.layout             || "single";
    AppState.builderForm.sections           = def.sections           || [];
    AppState.builderForm.conditions         = def.conditions         || [];
    AppState.builderForm.dependentDropdowns = def.dependentDropdowns || [];
    AppState.builderForm.specificPeople     = def.specificPeople     || [];
    AppState.builderForm.formManagers       = def.formManagers       || [];

    renderBuilder(main);
    showToast("info", "Data list removed — edit and re-submit for review when ready");

  } catch (e) {
    hideProgress();
    showToast("error", "Could not prepare form for editing: " + e.message);
    renderAdminReview(main);
  }
}

// =============================================================
// STATUS CHANGES
// =============================================================
async function adminChangeStatus(itemId, actionLabel) {
  const newStatus = CONFIG.STATUS_ACTION_MAP[actionLabel] || actionLabel;
  const isRejection = newStatus === "Rejected";
  const existingItem = AppState.allRequests?.find(r => r.id === itemId);
  const existingComment = existingItem?.fields?.[CONFIG.COL_COMMENTS] || "";

  openModal(html`
    <div class="modal-header">
      <span class="modal-title">${actionLabel}</span>
      <button class="btn btn-ghost btn-icon btn-sm" onclick="closeModal()">
        <svg width="14" height="14" viewBox="0 0 14 14" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M1 1l12 12M13 1L1 13"/></svg>
      </button>
    </div>
    <div class="modal-body">
      <div class="form-group">
        <label for="admin-comment" style="font-weight:600;">
          ${isRejection ? "Reason for rejection" : "Comment (optional)"}
        </label>
        <textarea id="admin-comment" class="textarea" style="min-height:100px;"
          placeholder="${isRejection ? "Please explain why this form has been rejected…" : "Add a note for the form author…"}"
        >${existingComment}</textarea>
        ${safeHtml(isRejection ? `<span class="input-hint">This will be visible to the form author.</span>` : `<span class="input-hint">Optional — visible to the form author.</span>`)}
      </div>
    </div>
    <div class="modal-footer">
      <button class="btn btn-ghost" onclick="closeModal()">Cancel</button>
      <button class="btn ${isRejection ? "btn-danger" : "btn-primary"}"
        data-id="${itemId}" data-action="${actionLabel}"
        onclick="confirmAdminAction(this)">
        ${actionLabel}
      </button>
    </div>
  `);
}

function confirmAdminAction(el) {
  const id = el.dataset.id;
  const action = el.dataset.action;
  const comment = document.getElementById("admin-comment")?.value || "";
  closeModal();
  doAdminChangeStatus(id, action, comment);
}

async function doAdminChangeStatus(itemId, actionLabel, comment) {
  const newStatus = CONFIG.STATUS_ACTION_MAP[actionLabel] || actionLabel;
  const needsProvisioning = newStatus === "Preview" || newStatus === "Approved";

  if (needsProvisioning) showProgress("Processing", "Updating status…");
  try {
    await updateListItem(CONFIG.FORMS_LIST, itemId, {
      [CONFIG.COL_STATUS]:   newStatus,
      [CONFIG.COL_COMMENTS]: comment || "",
    });

    if (needsProvisioning) {
      updateProgress("Provisioning SharePoint list…");
      await provisionDataList(itemId, newStatus);
    }

    if (needsProvisioning) hideProgress();
    showToast("success", `Status updated to "${newStatus}"`);
    renderAdminReview(document.getElementById("main-content"));
  } catch (e) {
    hideProgress();
    showToast("error", "Status update failed: " + e.message);
  }
}

// =============================================================
// PROVISION DATA LIST
// Called when a form is approved (for preview or live).
// Creates/recreates the SP data list and writes resolved column
// names back into the JSON so the renderer knows what to submit to.
// =============================================================
async function provisionDataList(itemId, newStatus) {
  const def = await getFormDefinition(CONFIG.FORMS_LIST, itemId);
  if (!def) throw new Error("No form definition found");

  const listName = def.listName;
  if (!listName) throw new Error("List name is missing from form definition");

  const liveStatus = newStatus === "Approved" ? "Live" : "Preview";
  const siteId = await getSiteId();

  if (liveStatus === "Live") {
    updateProgress("Promoting to Live…");
    await updateListItem(CONFIG.FORMS_LIST, itemId, { [CONFIG.COL_STATUS]: "Live" });
    return;
  }

  // Fetch the Forms item to get the author's email from createdBy
  let authorEmail = null;
  try {
    const listId = await getListId(CONFIG.FORMS_LIST);
    const formsItem = await graphGet(`/sites/${siteId}/lists/${listId}/items/${itemId}?expand=createdBy`);
    authorEmail = formsItem?.createdBy?.user?.email || null;
  } catch (e) {
    console.warn("Could not fetch author email:", e.message);
  }

  // Provisioning for Preview: delete any existing list and recreate fresh.
  try {
    const lists = await graphGet(`/sites/${siteId}/lists`);
    const existing = (lists.value || []).find(l => l.displayName === listName);
    if (existing) {
      updateProgress(`Deleting existing list "${listName}"…`);
      await graphDelete(`/sites/${siteId}/lists/${existing.id}`);
    }
  } catch (_) {}

  // Create the data list and write resolved internalNames back into def
  await doCreateSharePointList(listName, def, def.access || "StaffStudents", authorEmail);

  // Persist the updated definition (with resolved internalNames) back to the Forms item
  updateProgress("Saving resolved column names…");
  await uploadJsonAttachment(CONFIG.FORMS_LIST, itemId, "form-definition.json", def);

  // Set status to Preview
  await updateListItem(CONFIG.FORMS_LIST, itemId, { [CONFIG.COL_STATUS]: "Preview" });
}

// =============================================================
// LIVE FORM MANAGEMENT (admin view of deployed forms)
// =============================================================
async function renderAdminLive(container) {
  if (!AppState.isAdmin) {
    container.innerHTML = `<div class="empty-state"><h3>Access Denied</h3></div>`;
    return;
  }

  container.innerHTML = `
    <div class="flex items-center justify-between mb-4">
      <h1 style="font-size:22px;font-weight:600;letter-spacing:-0.02em;">Live Form Management</h1>
    </div>
    <div class="card" id="admin-live-table">
      <div style="padding:40px;text-align:center;"><span class="spinner"></span></div>
    </div>
  `;

  try {
    const items = await getListItems(CONFIG.FORMS_LIST);
    const deployed = items.filter(i => {
      const s = i.fields?.[CONFIG.COL_STATUS] || "";
      return (s === "Preview" || s === "Live") && !i.fields?.[CONFIG.COL_RETRO];
    });
    const card = document.getElementById("admin-live-table");

    if (!deployed.length) {
      card.innerHTML = `<div class="empty-state"><h3>No deployed forms yet</h3></div>`;
      return;
    }

    // Load JSON defs in parallel for access info
    const defs = await Promise.all(deployed.map(async item => {
      try { return await getFormDefinition(CONFIG.FORMS_LIST, item.id); }
      catch (_) { return null; }
    }));

    card.innerHTML = html`
      <div class="table-wrap">
        <table>
          <thead><tr><th>Form</th><th>List Name</th><th>Status</th><th>Access</th><th>Actions</th></tr></thead>
          <tbody>
            ${safeHtml(deployed.map((item, idx) => {
              const f = item.fields || {};
              const def = defs[idx];
              const accessLabel = CONFIG.ACCESS_OPTIONS.find(a => a.value === def?.access)?.label || def?.access || "—";
              return html`<tr>
                <td><strong>${f.Title||"—"}</strong></td>
                <td><span style="font-family:var(--mono);font-size:12px;color:var(--text2)">${f[CONFIG.COL_LISTNAME]||"—"}</span></td>
                <td>${safeHtml(statusBadge(f[CONFIG.COL_STATUS]||"—"))}</td>
                <td style="color:var(--text2);font-size:13px;">${accessLabel}</td>
                <td>
                  ${safeHtml(f[CONFIG.COL_STATUS] === "Preview" ? html`
                    <button class="btn btn-sm btn-primary" data-id="${item.id}" onclick="promoteToLive(this.dataset.id)">Promote to Live</button>
                  ` : html`
                    <div class="flex gap-2" style="gap:6px;">
                      <span style="font-size:12px;color:var(--text3);line-height:28px;">Live</span>
                      <button class="btn btn-sm btn-secondary" data-id="${item.id}" onclick="openSafeEditModal(this.dataset.id)">Edit Details</button>
                    </div>
                  `)}
                </td>
              </tr>`;
            }).join(""))}
          </tbody>
        </table>
      </div>
    `;
  } catch (e) {
    document.getElementById("admin-live-table").innerHTML =
      `<div class="empty-state"><p style="color:var(--red)">Error: ${escHtml(e.message)}</p></div>`;
  }
}

async function promoteToLive(itemId) {
  try {
    await updateListItem(CONFIG.FORMS_LIST, itemId, { [CONFIG.COL_STATUS]: "Live" });
    showToast("success", "Form promoted to Live");
    renderAdminLive(document.getElementById("main-content"));
  } catch (e) {
    showToast("error", "Failed: " + e.message);
  }
}


// =============================================================
// SP DATA LIST CREATION
// =============================================================
async function doCreateSharePointList(listName, def, access, authorEmail = null) {
  updateProgress(`Creating SharePoint list "${listName}"…`);
  const siteId = await getSiteId();

  const newList = await graphPost(`/sites/${siteId}/lists`, {
    displayName: listName,
    list: { template: "genericList" },
    columns: [],
  });
  const newListId = newList.id;
  updateProgress(`Creating columns…`);

  // First pass: resolve all column names with deduplication
  const allFields = (def.sections || []).flatMap(s => s.fields || []);
  const usedNames = new Set();
  const nameClashes = [];
  for (const field of allFields) {
    if (field.type === "InfoText") continue;
    const baseName = sanitiseColumnName(field.label || "Field");
    // SharePoint internal column names are capped at 32 characters
    const baseNameTruncated = baseName.slice(0, 32);
    let uniqueName = baseNameTruncated;
    let counter = 2;
    while (usedNames.has(uniqueName.toLowerCase())) {
      const suffix = `_${counter++}`;
      uniqueName = baseNameTruncated.slice(0, 32 - suffix.length) + suffix;
    }
    if (uniqueName !== baseName) {
      nameClashes.push(`"${field.label}" → ${uniqueName}`);
    }
    usedNames.add(uniqueName.toLowerCase());
    field.internalName = uniqueName;
  }

  // Second pass: create form definition columns — read back the SP-assigned internalName
  for (const field of allFields) {
    const colDef = buildColumnDefinition(field);
    if (!colDef) continue;
    try {
      const created = await graphPost(`/sites/${siteId}/lists/${newListId}/columns`, colDef);
      // SharePoint may silently rename columns (e.g. "Field" → "field2" to avoid reserved name clashes).
      // Always use the name SP actually assigned, not what we requested.
      const spName = created?.name || created?.columnGroup || null;
      if (spName && field.internalName !== spName) {
        if (CONFIG.DEBUG_LOGGING) console.log(`[Columns] SP renamed "${field.internalName}" → "${spName}"`);
        field.internalName = spName;
      }
    } catch (e) {
      console.warn(`Column "${field.label}" failed:`, e.message);
    }
  }

  // Add IsDeleted Yes/No column for soft delete
  try {
    await graphPost(`/sites/${siteId}/lists/${newListId}/columns`, {
      name: "IsDeleted",
      displayName: "IsDeleted",
      boolean: {},
    });
  } catch (e) {
    console.warn("IsDeleted column failed:", e.message);
  }

  if (nameClashes.length) {
    console.info(`[Columns] ${nameClashes.length} field name(s) auto-renamed: ${nameClashes.join(", ")}`);
  }

  // Set ReadSecurity=2 (users see own items only) and WriteSecurity=2 (users edit own items only)
  // Form Managers get Contribute with override so they bypass this and see all items
  try {
    updateProgress("Configuring list security…");
    await spMerge(`/_api/web/lists(guid'${newListId}')`, {
      "__metadata": { "type": "SP.List" },
      "ReadSecurity": 2,
      "WriteSecurity": 2,
    });
  } catch (e) {
    console.warn("List security settings failed:", e.message);
  }

  try {
    updateProgress("Setting permissions…");
    await setListPermissions(siteId, newListId, access, def.specificPeople || [], def.formManagers || [], def.submissionType || "Submit", authorEmail);
  } catch (e) {
    console.warn("Permission assignment failed:", e.message);
    showToast("info", `List created but permissions could not be set: ${e.message}`);
  }
}

function buildColumnDefinition(field) {
  const base = {
    name: sanitiseColumnName(field.internalName || field.label).slice(0, 32),
    displayName: field.label,
    required: !!field.required,
    description: field.description || "",
  };
  switch (field.type) {
    case "InfoText":    return null;
    case "FileUpload":  return null; // stored as native SP list item attachment
    case "Text":        return { ...base, text: {} };
    case "Note":
    case "RichText":    return { ...base, text: { allowMultipleLines: true, linesForEditing: 6 } };
    case "Number":      return { ...base, number: {} };
    case "Currency":    return { ...base, currency: {} };
    case "DateTime":    return { ...base, dateTime: { displayAs: "default", format: "dateOnly" } };
    case "Boolean":     return { ...base, boolean: {} };
    case "URL":         return { ...base, hyperlinkOrPicture: { isPicture: false } };
    case "User":        return { ...base, personOrGroup: { allowMultipleSelection: true } };
    case "Choice":      return { ...base, text: {} }; // choices are renderer-only; SP stores the selected value as plain text
    case "MultiChoice": return {
      ...base,
      choice: {
        choices: field.choices || [],
        allowTextEntry: false,
        displayAs: "checkBoxes",
      }
    };
    default:
      console.warn(`[buildColumnDefinition] Unknown field type "${field.type}" — skipping column creation`);
      return null;
  }
}

// =============================================================
// LIST PERMISSIONS
// =============================================================
async function ensureFormSubmitterRole() {
  // Returns the Id of the "Form Submitter" role definition, creating it if needed.
  // Permissions: ViewListItems (0x1) + AddListItems (0x2) = Low: 3, High: 0
  const rolesData = await spGet(`/_api/web/roledefinitions`);
  const roles = rolesData?.d?.results || [];
  const existing = roles.find(r => r.Name === CONFIG.SUBMITTER_ROLE);
  if (existing) return existing.Id;

  // Create the custom role definition
  const created = await spPost(`/_api/web/roledefinitions`, {
    "__metadata": { "type": "SP.RoleDefinition" },
    "Name": CONFIG.SUBMITTER_ROLE,
    "Description": "Can add and view own items only. Used for form submissions.",
    "BasePermissions": {
      "__metadata": { "type": "SP.BasePermissions" },
      "Low": "3",   // ViewListItems + AddListItems
      "High": "0",
    },
  });
  return created?.d?.Id;
}

async function ensureContributeRole(roles) {
  return roles.find(r => r.Name === "Contribute") ||
         roles.find(r => r.Name === "Edit") ||
         roles.find(r => r.BasePermissions?.Low === "1011028719");
}

async function setListPermissions(siteId, graphListId, access, specificPeople = [], formManagers = [], submissionType = "Submit", authorEmail = null) {
  updateProgress("Configuring list permissions…");
  const listBase = `/_api/web/lists(guid'${graphListId}')`;
  await spPost(`${listBase}/breakroleinheritance(copyRoleAssignments=false,clearSubscopes=true)`);

  // Get role definitions
  const rolesData = await spGet(`/_api/web/roledefinitions`);
  const roles = rolesData?.d?.results || [];

  const contributeRole = await ensureContributeRole(roles);
  if (!contributeRole) throw new Error("Could not find Contribute role definition on site");
  const contributeId = contributeRole.Id;

  // Ensure "Form Submitter" role exists (add-only, view own items — no edit)
  updateProgress("Ensuring Form Submitter role…");
  const submitterRoleId = await ensureFormSubmitterRole();

  // SubmitEdit → Contribute (can add + edit own items — WriteSecurity=2 restricts to own)
  // Submit only → Form Submitter (ViewListItems + AddListItems only — no edit)
  // Note: if Graph returns 404 on submission, the Entra group grant likely failed —
  // check the browser console for permission errors during provisioning.
  const submitterRoleToApply = submissionType === "SubmitEdit" ? contributeId : submitterRoleId;

  // ── Grant submitters the appropriate role ──────────────────
  if (access === "StaffStudents" || access === "StaffOnly") {
    const groupGuid = access === "StaffStudents" ? CONFIG.STUDENT_GROUP : CONFIG.STAFF_GROUP;
    const claimName = `c:0t.c|tenant|${groupGuid}`;
    updateProgress(`Granting submit access to group ${groupGuid}…`);
    try {
      const userData = await spPost(`/_api/web/ensureuser`, { logonName: claimName });
      const spPrincipalId = userData?.d?.Id;
      if (CONFIG.DEBUG_LOGGING) console.log("[Permissions] Submitter group principal ID:", spPrincipalId, "role:", submitterRoleToApply);
      if (!spPrincipalId) throw new Error(`ensureUser returned no ID for group ${groupGuid}`);
      await spPost(`${listBase}/roleassignments/addroleassignment(principalid=${spPrincipalId},roledefid=${submitterRoleToApply})`);
      if (CONFIG.DEBUG_LOGGING) console.log("[Permissions] Submitter group grant succeeded");
    } catch (e) {
      console.error("[Permissions] Submitter group grant failed:", e.message);
      throw new Error(`Could not resolve Entra group "${groupGuid}": ${e.message}`);
    }
  } else if (access === "Specific") {
    for (const person of specificPeople) {
      updateProgress(`Granting submit access to ${person.displayName}…`);
      try {
        const userEmail = person.email || person.mail || "";
        if (!userEmail) continue;
        const userData = await spPost(`/_api/web/ensureuser`, { logonName: `i:0#.f|membership|${userEmail}` });
        const spUserId = userData?.d?.Id;
        if (!spUserId) continue;
        await spPost(`${listBase}/roleassignments/addroleassignment(principalid=${spUserId},roledefid=${submitterRoleToApply})`);
        if (CONFIG.DEBUG_LOGGING) console.log("[Permissions] Specific person grant succeeded:", userEmail);
      } catch (e) {
        console.error(`[Permissions] Submit access failed for ${person.displayName}:`, e.message);
      }
    }
  }

  // ── Grant author Contribute ──────────────────────────────
  if (authorEmail) {
    updateProgress("Granting author full access…");
    try {
      const userData = await spPost(`/_api/web/ensureuser`, { logonName: `i:0#.f|membership|${authorEmail}` });
      const spUserId = userData?.d?.Id;
      if (CONFIG.DEBUG_LOGGING) console.log("[Permissions] Author principal ID:", spUserId, "email:", authorEmail);
      if (spUserId) {
        await spPost(`${listBase}/roleassignments/addroleassignment(principalid=${spUserId},roledefid=${contributeId})`);
        if (CONFIG.DEBUG_LOGGING) console.log("[Permissions] Author grant succeeded");
      }
    } catch (e) {
      console.error("[Permissions] Author access failed:", e.message);
    }
  } else {
    console.warn("[Permissions] No author email available — author will not have list access");
  }

  // ── Grant Form Managers Contribute ──────────────────────
  for (const manager of formManagers) {
    updateProgress(`Granting manager access to ${manager.displayName}…`);
    try {
      const userEmail = manager.email || manager.mail || "";
      if (!userEmail) continue;
      const userData = await spPost(`/_api/web/ensureuser`, { logonName: `i:0#.f|membership|${userEmail}` });
      const spUserId = userData?.d?.Id;
      if (!spUserId) continue;
      await spPost(`${listBase}/roleassignments/addroleassignment(principalid=${spUserId},roledefid=${contributeId})`);
      if (CONFIG.DEBUG_LOGGING) console.log("[Permissions] Manager grant succeeded:", userEmail);
    } catch (e) {
      console.error(`[Permissions] Manager access failed for ${manager.displayName}:`, e.message);
      showToast("info", `Could not grant manager access to ${manager.displayName} — add manually`);
    }
  }

  updateProgress("Permissions set successfully");
}

// =============================================================
// SAFE EDIT — Live form metadata editing
// Only touches the JSON definition and SP column displayName/description.
// Never deletes the data list, never changes status, never touches internalName.
// Only Choice fields benefit from free choice editing (they are text columns in SP).
// MultiChoice choices are shown read-only because they are SP choice columns.
// =============================================================

async function openSafeEditModal(itemId) {
  // Re-load fresh from server — never mutate a stale in-memory copy
  let def;
  try {
    def = await getFormDefinition(CONFIG.FORMS_LIST, itemId);
    if (!def) throw new Error("Form definition not found");
  } catch (e) {
    showToast("error", "Could not load form definition: " + e.message);
    return;
  }

  const allSections = def.sections || [];

  openModal(html`
    <div class="modal-header">
      <span class="modal-title">Edit Live Form Details</span>
      <button class="btn btn-ghost btn-icon btn-sm" onclick="closeModal()">
        <svg width="14" height="14" viewBox="0 0 14 14" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M1 1l12 12M13 1L1 13"/></svg>
      </button>
    </div>
    <div class="modal-body" style="display:flex;flex-direction:column;gap:16px;max-height:70vh;overflow-y:auto;">

      <div style="padding:10px 14px;background:rgba(34,197,94,0.07);border:1px solid rgba(34,197,94,0.2);border-radius:var(--radius-sm);font-size:13px;color:var(--text2);">
        The SharePoint data list and all submissions are preserved. Only display metadata is changed.
      </div>

      <div class="form-group">
        <label style="font-weight:600;">Form Title</label>
        <input id="se-title" class="input" value="${escHtml(def.title || "")}">
        <span class="input-hint">Shown in the form directory and header</span>
      </div>

      <div style="border-top:1px solid var(--border);padding-top:16px;">
        ${safeHtml(allSections.map((sec, si) => `
          <div style="border:1px solid var(--border);border-radius:var(--radius-sm);padding:14px 16px;margin-bottom:12px;">
            <div style="font-size:11px;text-transform:uppercase;letter-spacing:0.06em;color:var(--text3);margin-bottom:10px;">
              Section ${si + 1}
            </div>
            <div class="form-group" style="margin-bottom:12px;">
              <label>Section Title</label>
              <input class="input se-section-title" data-si="${si}" value="${escHtml(sec.title || "")}">
            </div>
            ${(sec.fields || []).map((field, fi) => renderSafeEditField(field, si, fi)).join("")}
          </div>
        `).join(""))}
      </div>

    </div>
    <div class="modal-footer">
      <button class="btn btn-ghost" onclick="closeModal()">Cancel</button>
      <button class="btn btn-primary" data-id="${itemId}" onclick="doSafeEdit(this.dataset.id)">Save Changes</button>
    </div>
  `, false, true);
}

function renderSafeEditField(field, si, fi) {
  // InfoText: only the HTML content is editable — no SP column involved
  if (field.type === "InfoText") {
    return html`
      <div style="padding:10px 0;border-top:1px solid var(--border);margin-top:4px;">
        <div style="font-size:11px;font-family:var(--mono);color:var(--text3);margin-bottom:6px;">InfoText</div>
        <div class="form-group">
          <label>Content (HTML)</label>
          <textarea class="textarea se-info-content" data-si="${si}" data-fi="${fi}" rows="3">${escHtml(field.infoContent || "")}</textarea>
        </div>
      </div>
    `;
  }

  // Choice: SP column is plain text — choices array is purely renderer-side, fully editable
  const choicesBlock = (field.type === "Choice") ? html`
    <div class="form-group">
      <label>Choices <span style="font-size:11px;color:var(--text3);font-weight:400;">(one per line — add or reorder freely)</span></label>
      <textarea class="textarea se-choices" data-si="${si}" data-fi="${fi}" rows="5">${escHtml((field.choices || []).join("\n"))}</textarea>
    </div>
  ` : (field.type === "MultiChoice") ? html`
    <div style="padding:8px 10px;background:var(--bg3);border-radius:var(--radius-sm);font-size:12px;color:var(--text3);">
      MultiChoice options are locked — stored as an SP choice column. Use the full edit workflow to change them.
    </div>
  ` : "";

  return html`
    <div style="padding:10px 0;border-top:1px solid var(--border);margin-top:4px;">
      <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:8px;">
        <span style="font-size:11px;font-family:var(--mono);color:var(--text3);">${field.type}</span>
        <label style="display:flex;align-items:center;gap:6px;font-size:12px;cursor:pointer;">
          <input type="checkbox" class="se-required" data-si="${si}" data-fi="${fi}" ${field.required ? "checked" : ""}> Required
        </label>
      </div>
      <div class="form-group">
        <label>Label</label>
        <input class="input se-label" data-si="${si}" data-fi="${fi}" value="${escHtml(field.label || "")}">
      </div>
      <div class="form-group">
        <label>Description / hint</label>
        <input class="input se-desc" data-si="${si}" data-fi="${fi}" value="${escHtml(field.description || "")}">
      </div>
      ${safeHtml(choicesBlock)}
    </div>
  `;
}

async function doSafeEdit(itemId) {
  showProgress("Saving", "Updating form definition…");
  try {
    // Always reload from server as the base — never trust the DOM alone
    const def = await getFormDefinition(CONFIG.FORMS_LIST, itemId);
    if (!def) throw new Error("Could not reload form definition");

    // ── Patch form title ──────────────────────────────────────
    const newTitle = document.getElementById("se-title")?.value.trim();
    if (newTitle) def.title = newTitle;

    // ── Patch section titles ──────────────────────────────────
    document.querySelectorAll(".se-section-title").forEach(el => {
      const si = parseInt(el.dataset.si);
      if (def.sections[si]) def.sections[si].title = el.value;
    });

    // ── Patch field metadata ──────────────────────────────────
    // label, description, required, infoContent — never touch id, type, internalName
    document.querySelectorAll(".se-label").forEach(el => {
      const field = def.sections[parseInt(el.dataset.si)]?.fields?.[parseInt(el.dataset.fi)];
      if (field) field.label = el.value;
    });
    document.querySelectorAll(".se-desc").forEach(el => {
      const field = def.sections[parseInt(el.dataset.si)]?.fields?.[parseInt(el.dataset.fi)];
      if (field) field.description = el.value;
    });
    document.querySelectorAll(".se-required").forEach(el => {
      const field = def.sections[parseInt(el.dataset.si)]?.fields?.[parseInt(el.dataset.fi)];
      if (field) field.required = el.checked;
    });
    document.querySelectorAll(".se-info-content").forEach(el => {
      const field = def.sections[parseInt(el.dataset.si)]?.fields?.[parseInt(el.dataset.fi)];
      if (field && field.type === "InfoText") field.infoContent = el.value;
    });

    // ── Patch Choice choices (Choice only — MultiChoice is SP-locked) ─────
    // Choices may be freely added, removed, or reordered because Choice columns
    // are stored as plain text in SP — the stored value is just a string.
    document.querySelectorAll(".se-choices").forEach(el => {
      const field = def.sections[parseInt(el.dataset.si)]?.fields?.[parseInt(el.dataset.fi)];
      if (field && field.type === "Choice") {
        field.choices = el.value.split("\n").map(c => c.trim()).filter(Boolean);
      }
    });

    // ── Write updated JSON back ───────────────────────────────
    updateProgress("Writing form definition…");
    await uploadJsonAttachment(CONFIG.FORMS_LIST, itemId, "form-definition.json", def);

    // ── Sync SP Forms list Title column ──────────────────────
    if (newTitle) {
      await updateListItem(CONFIG.FORMS_LIST, itemId, { Title: newTitle });
    }

    // ── Best-effort: patch SP column displayNames and descriptions ────────
    // Choice fields are text columns — Graph accepts displayName + description patches.
    // This keeps the SP list view column headers in sync with the form labels.
    // Failures here are non-fatal — the form still works correctly via internalName.
    try {
      updateProgress("Syncing column labels…");
      const siteId = await getSiteId();
      const lists = await graphGet(`/sites/${siteId}/lists`);
      const spList = (lists.value || []).find(l => l.displayName === def.listName);
      if (spList) {
        const allFields = (def.sections || []).flatMap(s => s.fields || []);
        for (const field of allFields) {
          if (!field.internalName) continue;
          if (field.type === "InfoText" || field.type === "FileUpload") continue;
          try {
            await graphPatch(
              `/sites/${siteId}/lists/${spList.id}/columns/${field.internalName}`,
              { displayName: field.label, description: field.description || "" }
            );
          } catch (e) {
            if (CONFIG.DEBUG_LOGGING) console.warn(`[SafeEdit] Column patch skipped for "${field.internalName}":`, e.message);
          }
        }
      }
    } catch (e) {
      if (CONFIG.DEBUG_LOGGING) console.warn("[SafeEdit] SP column sync failed (non-fatal):", e.message);
    }

    hideProgress();
    closeModal();
    showToast("success", "Form details updated — submissions are unaffected");
    renderAdminLive(document.getElementById("main-content"));

  } catch (e) {
    hideProgress();
    showToast("error", "Save failed: " + e.message);
  }
}