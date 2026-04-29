// Form Studio — Admin: Review & Approval

// =============================================================
// ADMIN REVIEW
// All form lifecycle stages live in CONFIG.FORMS_LIST.
// Status flow: Draft → Submitted → Approved for Preview → Approved/Rejected
// Approved items get Status set to "Preview" or "Live" after SP list creation.
// =============================================================
async function renderAdminReview(container) {
  if (!AppState.isAdmin) {
    container.innerHTML = `<div class="empty-state"><h3>Access Denied</h3><p>You don't have admin access.</p></div>`;
    return;
  }

  container.innerHTML = `
    <div class="flex items-center justify-between mb-4">
      <div>
        <h1 style="font-size:22px;font-weight:600;letter-spacing:-0.02em;">Review Form Requests</h1>
        <p style="color:var(--text2);font-size:13.5px;margin-top:2px;">Approve, preview, or reject submitted form definitions</p>
      </div>
    </div>
    <div class="card" id="admin-table">
      <div style="padding:40px;text-align:center;"><span class="spinner"></span></div>
    </div>
  `;

  try {
    const items = await getListItems(CONFIG.FORMS_LIST);
    AppState.allRequests = items;
    const card = document.getElementById("admin-table");

    if (!items.length) {
      card.innerHTML = `<div class="empty-state"><h3>No forms yet</h3></div>`;
      return;
    }

    card.innerHTML = `
      <div class="table-wrap">
        <table>
          <thead><tr><th>Form</th><th>Submitter</th><th>Status</th><th>Modified</th><th>Actions</th></tr></thead>
          <tbody>
            ${items.map(item => {
              const f = item.fields || {};
              const status = f[CONFIG.COL_STATUS] || "Draft";
              const author = item.createdBy?.user?.displayName || "—";
              const actions = CONFIG.STATUS_FLOW[status] || [];
              const createdByEmail = (item.createdBy?.user?.email || "").toLowerCase();
              const currentEmail   = (AppState.currentUser?.email || "").toLowerCase();
              const isAuthor       = createdByEmail && createdByEmail === currentEmail;
              const canEdit        = (status === "Approved for Preview" || (status === "Preview" && AppState.isAdmin)) && (AppState.isAdmin || isAuthor);
              return `<tr>
                <td><strong>${escHtml(f.Title||"—")}</strong></td>
                <td style="color:var(--text2);font-size:13px;">${escHtml(author)}</td>
                <td>${statusBadge(status)}</td>
                <td style="color:var(--text2);font-size:12.5px;">${formatDate(f.Modified)}</td>
                <td>
                  <div class="flex gap-2">
                    <button class="btn btn-sm btn-ghost" onclick="adminPreviewForm('${item.id}','${escAttr(f.Title||"")}')">
                      <svg width="12" height="12" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.5"><circle cx="8" cy="8" r="3"/><path d="M1 8s2.5-5 7-5 7 5 7 5-2.5 5-7 5-7-5-7-5z"/></svg>
                      View
                    </button>
                    ${canEdit ? `
                      <button class="btn btn-sm btn-secondary" onclick="editFormRequest('${item.id}')">
                        <svg width="12" height="12" viewBox="0 0 13 13" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M9.5 1.5l2 2-8 8H1.5v-2l8-8z"/></svg>
                        Edit
                      </button>
                    ` : ""}
                    ${actions.map(a => `
                      <button class="btn btn-sm ${a==="Reject"?"btn-danger":"btn-secondary"}"
                        onclick="adminChangeStatus('${item.id}','${a}')">
                        ${a}
                      </button>
                    `).join("")}
                    ${!actions.length && !canEdit ? `<span style="font-size:12px;color:var(--text3);">Locked</span>` : ""}
                  </div>
                </td>
              </tr>`;
            }).join("")}
          </tbody>
        </table>
      </div>
    `;
  } catch (e) {
    document.getElementById("admin-table").innerHTML = `<div class="empty-state"><p style="color:var(--red)">Error: ${escHtml(e.message)}</p></div>`;
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

    const allFields = sections.flatMap(s => s.fields || []).filter(f => f.type !== "InfoText");
    const summaryHtml = `
      <div style="background:var(--bg3);border:1px solid var(--border);border-radius:var(--radius-sm);padding:14px 16px;margin-bottom:20px;display:flex;flex-wrap:wrap;gap:16px 32px;">
        <div><span style="font-size:11px;text-transform:uppercase;letter-spacing:0.06em;color:var(--text3);">Sections</span><br><strong>${sections.length}</strong></div>
        <div><span style="font-size:11px;text-transform:uppercase;letter-spacing:0.06em;color:var(--text3);">Fields</span><br><strong>${allFields.length}</strong></div>
        <div><span style="font-size:11px;text-transform:uppercase;letter-spacing:0.06em;color:var(--text3);">Layout</span><br><strong>${layout === "multistep" ? "Multi-step" : "Single page"}</strong></div>
        <div><span style="font-size:11px;text-transform:uppercase;letter-spacing:0.06em;color:var(--text3);">Submission</span><br><strong>${def.submissionType === "SubmitEdit" ? "Submit & Edit" : "Submit Only"}</strong></div>
        <div><span style="font-size:11px;text-transform:uppercase;letter-spacing:0.06em;color:var(--text3);">Access</span><br><strong>${CONFIG.ACCESS_OPTIONS.find(a=>a.value===def.access)?.label || def.access || "—"}</strong></div>
      </div>
    `;

    body.innerHTML = summaryHtml + previewHtml;

    if (footer && actions.length) {
      footer.innerHTML = `
        <button class="btn btn-ghost" onclick="closeModal()">Close</button>
        ${actions.map(a => `
          <button class="btn btn-sm ${a==="Reject"?"btn-danger":"btn-primary"}"
            onclick="closeModal();adminChangeStatus('${itemId}','${a}')">
            ${a}
          </button>
        `).join("")}
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
async function editFormRequest(itemId) {
  openModal(`
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
        <li>Remove this form entry — you will re-submit for review</li>
        <li>Require re-approval before the form goes live again</li>
      </ul>
      <div style="margin-top:16px;padding:12px 14px;background:rgba(220,38,38,0.06);border:1px solid rgba(220,38,38,0.2);border-radius:var(--radius-sm);">
        <strong style="font-size:13px;color:var(--red);">This cannot be undone.</strong>
        <span style="font-size:13px;color:var(--text2);"> Export any submitted data before continuing.</span>
      </div>
    </div>
    <div class="modal-footer">
      <button class="btn btn-ghost" onclick="closeModal()">Cancel</button>
      <button class="btn btn-danger" onclick="closeModal();doEditFormRequest('${itemId}')">I understand — Edit anyway</button>
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

    // Delete the Forms list item itself
    updateProgress("Removing form entry…");
    try {
      const siteId   = await getSiteId();
      const formsListId = await getListId(CONFIG.FORMS_LIST);
      await graphDelete(`/sites/${siteId}/lists/${formsListId}/items/${itemId}`);
    } catch (e) {
      console.warn("Could not delete Forms item:", e.message);
    }

    hideProgress();

    // Load definition into a fresh builder — new item created on re-submit
    resetBuilderForm();
    AppState.builderMode                    = "create";
    AppState.builderItemId                  = null;
    AppState.builderForm.title              = def.title              || "";
    AppState.builderForm.listName           = def.listName           || generateListName(def.title);
    AppState.builderForm.access             = def.access             || "StaffStudents";
    AppState.builderForm.submissionType     = def.submissionType     || "Submit";
    AppState.builderForm.layout             = def.layout             || "single";
    AppState.builderForm.sections           = def.sections           || [];
    AppState.builderForm.conditions         = def.conditions         || [];
    AppState.builderForm.dependentDropdowns = def.dependentDropdowns || [];
    AppState.builderForm.specificPeople     = def.specificPeople     || [];

    renderBuilder(main);
    showToast("info", "Old entry removed — edit and re-submit for review when ready");

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
  const needsProvisioning = newStatus === "Approved for Preview" || newStatus === "Approved";

  if (needsProvisioning) showProgress("Processing", "Updating status…");
  try {
    await updateListItem(CONFIG.FORMS_LIST, itemId, { [CONFIG.COL_STATUS]: newStatus });

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

  const siteId = await getSiteId();

  // Delete existing data list if present (always recreate on approval)
  try {
    const lists = await graphGet(`/sites/${siteId}/lists`);
    const existing = (lists.value || []).find(l => l.displayName === listName);
    if (existing) {
      updateProgress(`Deleting existing list "${listName}"…`);
      await graphDelete(`/sites/${siteId}/lists/${existing.id}`);
    }
  } catch (_) {}

  // Create the data list and write resolved internalNames back into def
  await doCreateSharePointList(listName, def, def.access || "StaffStudents");

  // Persist the updated definition (with resolved internalNames) back to the Forms item
  updateProgress("Saving resolved column names…");
  await uploadJsonAttachment(CONFIG.FORMS_LIST, itemId, "form-definition.json", def);

  // Set live status on the Forms item ("Preview" or "Live")
  const liveStatus = newStatus === "Approved" ? "Live" : "Preview";
  await updateListItem(CONFIG.FORMS_LIST, itemId, { [CONFIG.COL_STATUS]: liveStatus });
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
      return s === "Preview" || s === "Live";
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

    card.innerHTML = `
      <div class="table-wrap">
        <table>
          <thead><tr><th>Form</th><th>List Name</th><th>Status</th><th>Access</th><th>Actions</th></tr></thead>
          <tbody>
            ${deployed.map((item, idx) => {
              const f = item.fields || {};
              const def = defs[idx];
              const accessLabel = CONFIG.ACCESS_OPTIONS.find(a => a.value === def?.access)?.label || def?.access || "—";
              return `<tr>
                <td><strong>${escHtml(f.Title||"—")}</strong></td>
                <td><span style="font-family:var(--mono);font-size:12px;color:var(--text2)">${escHtml(f[CONFIG.COL_LISTNAME]||"—")}</span></td>
                <td>${statusBadge(f[CONFIG.COL_STATUS]||"—")}</td>
                <td style="color:var(--text2);font-size:13px;">${escHtml(accessLabel)}</td>
                <td>
                  ${f[CONFIG.COL_STATUS] === "Preview" ? `
                    <button class="btn btn-sm btn-primary" onclick="promoteToLive('${item.id}')">Promote to Live</button>
                  ` : `<span style="font-size:12px;color:var(--text3);">Live</span>`}
                </td>
              </tr>`;
            }).join("")}
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

async function forceRecreateList(existingListId, defJson, access) {
  try {
    const siteId = await getSiteId();
    await graphDelete(`/sites/${siteId}/lists/${existingListId}`);
    const def = JSON.parse(defJson);
    await doCreateSharePointList(def.listName, def, def.access || access);
    showToast("success", "List recreated successfully");
  } catch (e) {
    showToast("error", "Recreation failed: " + e.message);
  }
}

// =============================================================
// SP DATA LIST CREATION
// =============================================================
async function doCreateSharePointList(listName, def, access) {
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
    const baseName = (field.label || "Field")
      .replace(/[^a-zA-Z0-9_]/g, "_").replace(/_+/g, "_").replace(/^_|_$/g, "") || "Field";
    let uniqueName = baseName;
    let counter = 2;
    while (usedNames.has(uniqueName.toLowerCase())) {
      uniqueName = `${baseName}_${counter++}`;
    }
    if (uniqueName !== baseName) {
      nameClashes.push(`"${field.label}" → ${uniqueName}`);
    }
    usedNames.add(uniqueName.toLowerCase());
    field.internalName = uniqueName;
  }

  // Second pass: create columns
  for (const field of allFields) {
    const colDef = buildColumnDefinition(field);
    if (!colDef) continue;
    try {
      await graphPost(`/sites/${siteId}/lists/${newListId}/columns`, colDef);
    } catch (e) {
      console.warn(`Column "${field.label}" failed:`, e.message);
    }
  }

  if (nameClashes.length) {
    console.info(`[Columns] ${nameClashes.length} field name(s) auto-renamed: ${nameClashes.join(", ")}`);
  }

  try {
    updateProgress("Setting permissions…");
    await setListPermissions(siteId, newListId, access, def.specificPeople || []);
  } catch (e) {
    console.warn("Permission assignment failed:", e.message);
    showToast("info", `List created but permissions could not be set: ${e.message}`);
  }
}

function buildColumnDefinition(field) {
  const base = {
    name: field.internalName || field.label.replace(/[^a-zA-Z0-9_]/g, "_").replace(/_+/g, "_").replace(/^_|_$/g, "") || "Field",
    displayName: field.label,
    required: !!field.required,
    description: field.description || "",
  };
  switch (field.type) {
    case "InfoText":  return null;
    case "Text":      return { ...base, text: {} };
    case "Note":
    case "RichText":  return { ...base, text: { allowMultipleLines: true, linesForEditing: 6 } };
    case "Number":    return { ...base, number: {} };
    case "Currency":  return { ...base, currency: {} };
    case "DateTime":  return { ...base, dateTime: { displayAs: "default", format: "dateOnly" } };
    case "Boolean":   return { ...base, boolean: {} };
    case "URL":       return { ...base, hyperlinkOrPicture: { isPicture: false } };
    case "User":      return { ...base, personOrGroup: { allowMultipleSelection: false } };
    case "Choice":
    case "MultiChoice": return {
      ...base,
      choice: {
        choices: field.choices || [],
        allowTextEntry: false,
        displayAs: field.type === "MultiChoice" ? "checkBoxes" : "dropDownMenu",
      }
    };
    default: return { ...base, text: {} };
  }
}

// =============================================================
// LIST PERMISSIONS
// =============================================================
async function setListPermissions(siteId, graphListId, access, specificPeople = []) {
  updateProgress("Configuring list permissions…");
  const listBase = `/_api/web/lists(guid'${graphListId}')`;
  await spPost(`${listBase}/breakroleinheritance(copyRoleAssignments=false,clearSubscopes=true)`);

  const rolesData = await spGet(`/_api/web/roledefinitions`);
  const roles = rolesData?.d?.results || [];
  const contributeRole = roles.find(r => r.Name === "Contribute") ||
                         roles.find(r => r.Name === "Edit") ||
                         roles.find(r => r.BasePermissions?.Low === "1011028719");
  if (!contributeRole) throw new Error("Could not find Contribute role definition on site");
  const roleDefId = contributeRole.Id;

  if (access === "StaffStudents" || access === "StaffOnly") {
    const groupNames = access === "StaffStudents"
      ? [CONFIG.STAFF_GROUP, CONFIG.STUDENT_GROUP]
      : [CONFIG.STAFF_GROUP];
    for (const groupName of groupNames) {
      updateProgress(`Granting access to ${groupName}…`);
      try {
        // Use ensureuser with the c:0t.c|tenant| claim + Entra Object ID to resolve
        // M365/Entra groups — no directory read permissions required, avoids sitegroups
        // API which only works for classic SharePoint groups.
        const userData = await spPost(`/_api/web/ensureuser`, { logonName: `c:0t.c|tenant|${groupName}` });
        const principalId = userData?.d?.Id;
        if (!principalId) throw new Error(`ensureuser returned no ID`);
        await spPost(`${listBase}/roleassignments/addroleassignment(principalid=${principalId},roledefid=${roleDefId})`);
      } catch (e) {
        console.warn(`Permission assignment failed for "${groupName}": ${e.message}`);
        showToast("info", `Could not assign permissions to "${groupName}" — add manually in SharePoint`);
      }
    }
  } else if (access === "Specific") {
    if (!specificPeople.length) return;
    for (const person of specificPeople) {
      updateProgress(`Granting access to ${person.displayName}…`);
      try {
        const userEmail = person.email || person.mail || "";
        if (!userEmail) continue;
        const userData = await spPost(`/_api/web/ensureuser`, { logonName: `i:0#.f|membership|${userEmail}` });
        const spUserId = userData?.d?.Id;
        if (!spUserId) throw new Error(`ensureUser returned no ID for ${userEmail}`);
        await spPost(`${listBase}/roleassignments/addroleassignment(principalid=${spUserId},roledefid=${roleDefId})`);
      } catch (e) {
        console.warn(`Could not grant access to ${person.displayName}: ${e.message}`);
        showToast("info", `Could not grant access to ${person.displayName} — add manually`);
      }
    }
  }
  updateProgress("Permissions set successfully");
}
