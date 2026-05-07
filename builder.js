// Form Studio — Form Builder Wizard

// =============================================================
// MY REQUESTS VIEW (kept for internal use)
// =============================================================
// =============================================================
// FORM BUILDER — WIZARD
// =============================================================

const WIZARD_STEPS = [
  { key: "identity",      label: "Identity" },
  { key: "governance",    label: "Governance" },
  { key: "sections",      label: "Sections & Fields" },
  { key: "conditions",    label: "Conditions" },
  { key: "dependents",    label: "Linked Dropdowns" },
  { key: "layout",        label: "Layout" },
  { key: "access",        label: "Access" },
  { key: "onsubmit",      label: "On Submit" },
  { key: "review",        label: "Review" },
];

function startNewForm(container) {
  resetBuilderForm();
  AppState.currentView = "new-form";
  renderBuilder(container);
}

function renderBuilder(container) {
  const step = AppState.builderStep;
  container.innerHTML = html`
    <div class="flex items-center justify-between mb-4">
      <div>
        <h1 style="font-size:22px;font-weight:600;letter-spacing:-0.02em;">${AppState.builderMode === "edit" ? "Edit Form Request" : "New Form Request"}</h1>
        <p style="color:var(--text2);font-size:13.5px;margin-top:2px;">${WIZARD_STEPS[step].label}</p>
      </div>
      <button class="btn btn-ghost" onclick="navigateTo('admin-review')">
        <svg width="14" height="14" viewBox="0 0 14 14" fill="currentColor"><path d="M1 7h12M1 7l5-5M1 7l5 5"/><path d="M1 7h12M1 7l5-5M1 7l5 5" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" fill="none"/></svg>
        Back
      </button>
    </div>

    <!-- Wizard Steps -->
    <div class="wizard-steps" role="list" aria-label="Form builder steps">
      ${safeHtml(WIZARD_STEPS.map((s, i) => html`
        <div class="wizard-step ${i < step ? "done" : i === step ? "active" : ""}"
          role="listitem"
          aria-current="${i === step ? "step" : "false"}"
          aria-label="Step ${i + 1} of ${WIZARD_STEPS.length}: ${s.label}${i < step ? " — completed" : i === step ? " — current" : ""}">
          <div class="wizard-step-num" aria-hidden="true">
            ${safeHtml(i < step
              ? `<svg width="12" height="12" viewBox="0 0 12 12" fill="currentColor"><path d="M10 3L4.5 8.5 2 6"/><path d="M10 3L4.5 8.5 2 6" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" fill="none"/></svg>`
              : String(i + 1)
            )}
          </div>
          <span class="wizard-step-label">${s.label}</span>
        </div>
        ${safeHtml(i < WIZARD_STEPS.length - 1 ? '<div class="wizard-connector" aria-hidden="true"></div>' : "")}
      `).join(""))}
    </div>

    <!-- Step content -->
    <div id="wizard-step-content"></div>

    <!-- Navigation -->
    <div class="flex items-center justify-between mt-4" style="margin-top:24px;">
      <div class="flex gap-2">
        ${safeHtml(step > 0 ? `<button class="btn btn-secondary" onclick="wizardBack()">← Previous</button>` : "")}
      </div>
      <div class="flex gap-2">
        <span id="autosave-indicator" style="font-size:12px;color:var(--text3);align-self:center;"></span>
        ${safeHtml(step < WIZARD_STEPS.length - 1
          ? `<button class="btn btn-primary" onclick="wizardNext()">Next →</button>`
          : `<button class="btn btn-primary" onclick="submitBuilderForReview()">Save Form</button>`
        )}
      </div>
    </div>
  `;

  renderWizardStep();
}

function renderWizardStep() {
  const container = document.getElementById("wizard-step-content");
  if (!container) return;
  const step = WIZARD_STEPS[AppState.builderStep].key;
  switch (step) {
    case "identity":    renderStepIdentity(container); break;
    case "governance":  renderStepGovernance(container); break;
    case "sections":    renderStepSections(container); break;
    case "conditions":  renderStepConditions(container); break;
    case "dependents":  renderStepDependents(container); break;
    case "layout":      renderStepLayout(container); break;
    case "access":      renderStepAccess(container); break;
    case "onsubmit":    renderStepOnSubmit(container); break;
    case "review":      renderStepReview(container); break;
  }
}

async function wizardNext() {
  if (!await validateCurrentStep()) return;
  AppState.builderStep = Math.min(AppState.builderStep + 1, WIZARD_STEPS.length - 1);
  renderBuilder(document.getElementById("main-content"));
  autoSaveBuilder();
}

function wizardBack() {
  AppState.builderStep = Math.max(AppState.builderStep - 1, 0);
  renderBuilder(document.getElementById("main-content"));
  autoSaveBuilder();
}

// Silent auto-save — no progress modal, just a small indicator
async function autoSaveBuilder() {
  const indicator = document.getElementById("autosave-indicator");
  if (indicator) indicator.textContent = "Saving…";
  try {
    await saveBuilderDraft();
    if (indicator) {
      indicator.textContent = "✓ Saved";
      setTimeout(() => { if (indicator) indicator.textContent = ""; }, 2000);
    }
  } catch (_) {
    if (indicator) indicator.textContent = "Save failed";
  }
}

async function validateCurrentStep() {
  const step = WIZARD_STEPS[AppState.builderStep].key;
  if (step === "identity") {
    const title = document.getElementById("form-title")?.value?.trim();
    if (!title) { showToast("error", "Form Name is required"); return false; }
    AppState.builderForm.title    = title;
    AppState.builderForm.listName = generateListName(title);

    // Check for duplicate form title — skip when editing an existing form
    if (AppState.builderMode !== "edit") {
      try {
        const existing = await getListItems(CONFIG.FORMS_LIST);
        const duplicate = existing.some(i =>
          i.id !== AppState.builderItemId &&
          (i.fields?.Title || "").trim().toLowerCase() === title.toLowerCase()
        );
        if (duplicate) {
          showToast("error", `A form called "${title}" already exists. Please choose a different name.`);
          return false;
        }
      } catch (_) {
        // Non-fatal — if the check fails don't block the author
      }
    }
  }
  if (step === "governance") {
    // Capture all governance fields from the DOM into state
    const g = AppState.builderForm.governance;
    g.existingProcess       = document.getElementById("gov-existing-process")?.value || "";
    g.existingProcessDetail = document.getElementById("gov-existing-process-detail")?.value?.trim() || "";
    g.retention             = document.getElementById("gov-retention")?.value || "";
    g.sensitiveData         = document.getElementById("gov-sensitive-data")?.value || "";
    g.privacyAssessment     = document.getElementById("gov-privacy-assessment")?.value || "";
    g.externalAccess        = document.getElementById("gov-external-access")?.value || "";
    g.continuityPlan        = document.getElementById("gov-continuity-plan")?.value?.trim() || "";
    g.expectedVolume        = document.getElementById("gov-expected-volume")?.value || "";
    // g.dataOwner is set directly into state by the people-picker — no DOM read needed

    // These fields are required — without them the admin cannot make an informed decision
    if (!g.retention)          { showToast("error", "Please select a data retention period"); return false; }
    if (!g.sensitiveData)      { showToast("error", "Please select a data sensitivity level"); return false; }
    if (!g.privacyAssessment)  { showToast("error", "Please select whether a Privacy Impact Assessment has been completed"); return false; }
    if (!g.externalAccess)     { showToast("error", "Please select an external access option"); return false; }
    if (!g.dataOwner)          { showToast("error", "Please select a data owner — e.g. your Dept Head"); return false; }
  }
  if (step === "sections") {
    // Block field labels that would clash with system-managed columns
    // (AssignedTo, IsDeleted, Title, ID, Status). Sanitised, case-insensitive —
    // "Assigned To", "assigned-to", "AssignedTo!" all collapse to "assignedto"
    // because that's what would land in SharePoint as the internal name.
    const reserved = (CONFIG.RESERVED_FIELD_NAMES || []).map(n => n.toLowerCase());
    for (const section of (AppState.builderForm.sections || [])) {
      for (const field of (section.fields || [])) {
        if (field.type === "InfoText") continue; // display-only, no SP column created
        const sanitised = sanitiseColumnName(field.internalName || field.label).toLowerCase();
        if (reserved.includes(sanitised)) {
          showToast("error", `"${field.label}" is a reserved field name and cannot be used. Please rename it.`);
          return false;
        }
      }
    }
  }
  if (step === "onsubmit") {
    const emails = document.getElementById("onsubmit-notify-emails")?.value?.trim() || "";
    // Validate each entry is a plausible email address
    if (emails) {
      const invalid = emails.split(",").map(e => e.trim()).filter(e => e && !e.includes("@"));
      if (invalid.length) {
        showToast("error", `Invalid email address: "${invalid[0]}"`);
        return false;
      }
    }
    AppState.builderForm.submitNotifyEmails = emails;
    AppState.builderForm.notifySubmitter    = document.getElementById("onsubmit-notify-submitter")?.checked ?? true;
  }
  return true;
}
// ---- Step 1: Identity ----
function renderStepIdentity(container) {
  container.innerHTML = html`
    <div class="card">
      <div class="card-header">
        <div>
          <div class="card-title">Form Identity</div>
          <div class="card-subtitle">Give your form a name</div>
        </div>
      </div>
      <div class="card-body" style="display:flex;flex-direction:column;gap:20px;">
        <div class="form-group">
          <label for="form-title">Form Name *</label>
          <input id="form-title" class="input" placeholder="e.g. Room Booking Request"
            value="${AppState.builderForm.title}"
            oninput="AppState.builderForm.title=this.value;AppState.builderForm.listName=generateListName(this.value);">
          <span class="input-hint">Shown to end users in the form directory</span>
        </div>
      </div>
    </div>
  `;
}

function generateListName(title) {
  // PascalCase, max 60 chars, SharePoint-safe
  return (title || "")
    .replace(/[^a-zA-Z0-9 ]/g, " ")   // strip special chars
    .trim()
    .split(/\s+/)
    .filter(Boolean)
    .map(w => w.charAt(0).toUpperCase() + w.slice(1).toLowerCase())
    .join("")
    .slice(0, 60) || "FormList";
}
// ---- Step 2: Governance ----
function renderStepGovernance(container) {
  const g = AppState.builderForm.governance;
  const showDetail = g.existingProcess === "yes";

  container.innerHTML = html`
    <div style="display:flex;flex-direction:column;gap:16px;">

      <div class="card">
        <div class="card-header">
          <div>
            <div class="card-title">Governance & Suitability</div>
            <div class="card-subtitle">These answers help us determine whether SharePoint Lists is the right storage for this form's data</div>
          </div>
        </div>
        <div class="card-body" style="display:flex;flex-direction:column;gap:20px;">

          <!-- Existing process -->
          <div class="form-group">
            <label for="gov-existing-process">Is this form based on an existing process?</label>
            <select id="gov-existing-process" class="select"
              onchange="
                AppState.builderForm.governance.existingProcess=this.value;
                const d=document.getElementById('gov-existing-process-detail-group');
                if(d) d.style.display=this.value==='yes'?'':'none';
              ">
              <option value="" ${!g.existingProcess ? "selected" : ""}>— Select —</option>
              <option value="yes" ${g.existingProcess === "yes" ? "selected" : ""}>Yes — this replaces or digitises an existing process</option>
              <option value="no"  ${g.existingProcess === "no"  ? "selected" : ""}>No — this is a brand new process</option>
            </select>
          </div>

          <div class="form-group" id="gov-existing-process-detail-group" style="${showDetail ? "" : "display:none"}">
            <label for="gov-existing-process-detail">Briefly describe the existing process</label>
            <textarea id="gov-existing-process-detail" class="textarea" rows="2"
              placeholder="e.g. Currently handled via a paper form submitted to the departmental office">${g.existingProcessDetail}</textarea>
          </div>

          <!-- Retention -->
          <div class="form-group">
            <label for="gov-retention">Data retention — how long do you need to store submissions? *</label>
            <select id="gov-retention" class="select">
              <option value=""          ${!g.retention            ? "selected" : ""}>— Select —</option>
              <option value="under1"    ${g.retention==="under1"  ? "selected" : ""}>Less than 1 year</option>
              <option value="1to3"      ${g.retention==="1to3"    ? "selected" : ""}>1 – 3 years</option>
              <option value="3to7"      ${g.retention==="3to7"    ? "selected" : ""}>3 – 7 years</option>
              <option value="indefinite"${g.retention==="indefinite"?"selected":""}>Indefinitely / unknown</option>
            </select>
            <span class="input-hint">SharePoint has no automated purge — longer retention periods require a manual process or Purview retention policy</span>
          </div>

          <!-- Sensitive data -->
          <div class="form-group">
            <label for="gov-sensitive-data">Data sensitivity — what type of sensitive information will this form collect? *</label>
            <select id="gov-sensitive-data" class="select">
              <option value=""         ${!g.sensitiveData             ? "selected" : ""}>— Select —</option>
              <option value="none"     ${g.sensitiveData==="none"     ? "selected" : ""}>No sensitive data</option>
              <option value="personal" ${g.sensitiveData==="personal" ? "selected" : ""}>Special category personal data (e.g. bank details, health information)</option>
              <option value="commercial"${g.sensitiveData==="commercial"?"selected":""}>Commercially sensitive information</option>
              <option value="both"     ${g.sensitiveData==="both"     ? "selected" : ""}>Both personal and commercially sensitive</option>
            </select>
            <span class="input-hint">Bank details and health data should not be stored in a standard SharePoint list — select this to flag for admin review</span>
          </div>

          <!-- Privacy Impact Assessment -->
          <div class="form-group">
            <label for="gov-privacy-assessment">Has a Privacy Impact Assessment (PIA / DPIA) been completed for this process? *</label>
            <select id="gov-privacy-assessment" class="select">
              <option value=""    ${!g.privacyAssessment         ? "selected" : ""}>— Select —</option>
              <option value="yes" ${g.privacyAssessment==="yes"  ? "selected" : ""}>Yes — a PIA / DPIA has been completed</option>
              <option value="no"  ${g.privacyAssessment==="no"   ? "selected" : ""}>No — one has not been done</option>
              <option value="na"  ${g.privacyAssessment==="na"   ? "selected" : ""}>Not applicable — no personal data is collected</option>
            </select>
            <span class="input-hint">Required under UK GDPR for high-risk processing. A "No" answer will be flagged for the Data Protection team</span>
          </div>

          <!-- External access -->
          <div class="form-group">
            <label for="gov-external-access">External access — will people outside the University be involved with this data? *</label>
            <select id="gov-external-access" class="select">
              <option value=""           ${!g.externalAccess               ? "selected" : ""}>— Select —</option>
              <option value="none"       ${g.externalAccess==="none"       ? "selected" : ""}>No — internal use only</option>
              <option value="recipients" ${g.externalAccess==="recipients" ? "selected" : ""}>Yes — submission data will be shared with external parties</option>
              <option value="submitters" ${g.externalAccess==="submitters" ? "selected" : ""}>Yes — external people will submit this form</option>
            </select>
            <span class="input-hint">External sharing requires specific SharePoint tenant settings and may have data transfer implications</span>
          </div>

          <!-- Continuity plan -->
          <div class="form-group">
            <label for="gov-continuity-plan">Continuity plan — if this form stops working, what is the workaround? *</label>
            <textarea id="gov-continuity-plan" class="textarea" rows="2"
              placeholder="e.g. Staff would revert to emailing requests directly to the team inbox">${g.continuityPlan}</textarea>
            <span class="input-hint">Helps us understand how business-critical this form is and plan maintenance windows accordingly</span>
          </div>

        </div>
      </div>

      <div class="card">
        <div class="card-header">
          <div>
            <div class="card-title">Additional Information</div>
            <div class="card-subtitle">Optional but recommended</div>
          </div>
        </div>
        <div class="card-body" style="display:flex;flex-direction:column;gap:20px;">

          <!-- Expected volume -->
          <div class="form-group">
            <label for="gov-expected-volume">Expected submission volume</label>
            <select id="gov-expected-volume" class="select">
              <option value=""      ${!g.expectedVolume          ? "selected" : ""}>— Select —</option>
              <option value="low"   ${g.expectedVolume==="low"   ? "selected" : ""}>Low — fewer than 100 submissions per month</option>
              <option value="medium"${g.expectedVolume==="medium"? "selected" : ""}>Medium — 100 – 1,000 per month</option>
              <option value="high"  ${g.expectedVolume==="high"  ? "selected" : ""}>High — more than 1,000 per month</option>
            </select>
            <span class="input-hint">SharePoint list performance can degrade at high volumes — large-scale forms may need a different storage solution</span>
          </div>

          <!-- Data owner -->
          <div class="form-group">
            <label for="gov-data-owner-search">Data owner *</label>
            ${safeHtml(g.dataOwner ? html`
              <div style="display:flex;align-items:center;gap:8px;margin-bottom:8px;">
                <span class="person-chip">
                  <div class="avatar" style="width:20px;height:20px;font-size:9px;">${g.dataOwner.displayName.split(" ").map(n=>n[0]).join("").slice(0,2)}</div>
                  ${g.dataOwner.displayName}
                  <button onclick="removeDataOwner()" aria-label="Remove data owner">
                    <svg width="10" height="10" viewBox="0 0 10 10" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M1 1l8 8M9 1L1 9"/></svg>
                  </button>
                </span>
              </div>
            ` : html`
              <div class="flex gap-2">
                <input id="gov-data-owner-search" class="input" placeholder="Search by name or email…"
                  oninput="debouncedDataOwnerSearch(this.value)">
                <button class="btn btn-secondary" onclick="searchDataOwnerNow()">Search</button>
              </div>
              <div id="gov-data-owner-results" style="margin-top:8px;"></div>
            `)}
            <span class="input-hint">The person accountable for this data if a Subject Access Request or data breach occurs — e.g. your Dept Head</span>
          </div>

        </div>
      </div>

    </div>
  `;
}

// ---- Data Owner people-picker (governance step) ----
// Exactly one person, not a list — selecting a person replaces any previous selection.

// Data owner search — single-select, replaces existing selection
const _dataOwnerSearch = createPeopleSearch({
  inputId:   "gov-data-owner-search",
  resultsId: "gov-data-owner-results",
  onClickFn: "setDataOwnerFromEl",
});
function debouncedDataOwnerSearch(val) { _dataOwnerSearch.debounced(val); }
async function searchDataOwnerNow(q)   { await _dataOwnerSearch.search(q); }

// Snapshot all governance DOM values into state before any re-render.
// validateCurrentStep does this on Next — but mid-step re-renders (picker select/remove)
// happen before Next is clicked, so we must capture first or the selects reset.
function snapshotGovernanceDom() {
  const g = AppState.builderForm.governance;
  const r = (id) => document.getElementById(id)?.value;
  const t = (id) => document.getElementById(id)?.value?.trim() || "";
  if (r("gov-existing-process")        !== undefined) g.existingProcess       = r("gov-existing-process");
  if (r("gov-existing-process-detail") !== undefined) g.existingProcessDetail = t("gov-existing-process-detail");
  if (r("gov-retention")               !== undefined) g.retention             = r("gov-retention");
  if (r("gov-sensitive-data")          !== undefined) g.sensitiveData         = r("gov-sensitive-data");
  if (r("gov-privacy-assessment")      !== undefined) g.privacyAssessment     = r("gov-privacy-assessment");
  if (r("gov-external-access")         !== undefined) g.externalAccess        = r("gov-external-access");
  if (r("gov-continuity-plan")         !== undefined) g.continuityPlan        = t("gov-continuity-plan");
  if (r("gov-expected-volume")         !== undefined) g.expectedVolume        = r("gov-expected-volume");
}

function setDataOwnerFromEl(el) {
  snapshotGovernanceDom();
  AppState.builderForm.governance.dataOwner = {
    id:          el.dataset.id,
    displayName: el.dataset.name,
    email:       el.dataset.email,
  };
  renderStepGovernance(document.getElementById("wizard-step-content"));
}

function removeDataOwner() {
  snapshotGovernanceDom();
  AppState.builderForm.governance.dataOwner = null;
  renderStepGovernance(document.getElementById("wizard-step-content"));
}

// ---- Step 3: Sections & Fields ----
function renderStepSections(container) {
  const { sections } = AppState.builderForm;

  container.innerHTML = html`
    <div style="display:flex;flex-direction:column;gap:16px;" id="sections-container">
      ${safeHtml(sections.map((sec, si) => renderSectionBlock(sec, si)).join(""))}
    </div>
    <button class="btn btn-secondary mt-4" onclick="addSection()" style="margin-top:16px;">
      <svg width="14" height="14" viewBox="0 0 14 14" fill="currentColor"><path d="M7 1v12M1 7h12" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" fill="none"/></svg>
      Add Section
    </button>
    ${safeHtml(!sections.length ? `<div class="empty-state" style="padding:32px;"><h3>No sections yet</h3><p>Add a section to start building your form</p></div>` : "")}
  `;
}

function renderSectionBlock(sec, si) {
  const nonSystemFieldCount = sec.fields.filter(f => !f.system).length;

  // Build sub-sections as plain strings to avoid nested template escaping issues

  const notifyToggle = sec.managerOnly ? `
    <label style="display:flex;align-items:center;gap:5px;font-size:12px;color:${sec.notify ? "var(--accent)" : "var(--text3)"};cursor:pointer;white-space:nowrap;border-left:1px solid var(--border);padding-left:10px;"
      title="When enabled, managers can mark this section complete and notify a team by email">
      <input type="checkbox" style="width:13px;height:13px;cursor:pointer;accent-color:var(--accent);" data-si="${si}"
        ${sec.notify ? "checked" : ""}
        onchange="toggleSectionNotify(+this.dataset.si, this.checked)">
      Notify
    </label>
  ` : "";

  const notifyRow = (sec.managerOnly && sec.notify) ? `
    <div class="section-notify-row">
      <svg width="12" height="12" viewBox="0 0 16 16" fill="none" stroke="var(--text3)" stroke-width="1.5"><path d="M2 4l6 5 6-5M2 4h12v9a1 1 0 01-1 1H3a1 1 0 01-1-1V4z"/></svg>
      <label style="font-size:12px;color:var(--text3);white-space:nowrap;">Notification emails:</label>
      <input
        class="input"
        style="font-size:12px;padding:4px 8px;height:28px;flex:1;"
        placeholder="e.g. dept@le.ac.uk, manager@le.ac.uk"
        value="${escAttr(sec.deptEmail || "")}"
        title="Comma-separated email addresses notified when this section is marked complete"
        data-si="${si}"
        oninput="AppState.builderForm.sections[+this.dataset.si].deptEmail=this.value">
    </div>
  ` : "";

  const infoBar = sec.managerOnly ? `
    <div style="padding:6px 14px;background:rgba(79,124,255,0.06);border-bottom:1px solid var(--border);font-size:12px;color:var(--accent);display:flex;align-items:center;gap:6px;">
      <svg width="12" height="12" viewBox="0 0 16 16" fill="currentColor"><path d="M7.122.392a1.75 1.75 0 011.756 0l5.25 3.045c.54.313.872.89.872 1.514V8.64c0 2.048-1.19 3.914-3.05 4.856l-2.5 1.286a1.75 1.75 0 01-1.6 0l-2.5-1.286C3.19 12.554 2 10.688 2 8.64V4.951c0-.624.332-1.2.872-1.514L7.122.392z"/></svg>
      This section is only visible to Form Managers and Admins
      ${sec.notify && sec.deptEmail
        ? `&nbsp;·&nbsp;<svg width="11" height="11" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M2 4l6 5 6-5M2 4h12v9a1 1 0 01-1 1H3a1 1 0 01-1-1V4z"/></svg> Notifies: ${escHtml(sec.deptEmail)}`
        : sec.notify
          ? `&nbsp;·&nbsp;<span style="color:var(--red,#ef4444);">⚠ No notification emails set</span>`
          : ""}
    </div>
  ` : "";

  return html`
    <div class="section-block${sec.managerOnly ? " section-manager-only" : ""}" id="section-${sec.id}">

      <div class="section-header">
        <svg width="14" height="14" viewBox="0 0 14 14" fill="none" stroke="var(--text3)" stroke-width="1.5"><path d="M1 4h12M1 7h12M1 10h8"/></svg>
        <input class="section-title-input" placeholder="Section title (optional)"
          value="${sec.title}"
          oninput="AppState.builderForm.sections[${si}].title=this.value">
        <span style="color:var(--text3);font-size:12px;">${nonSystemFieldCount} field${nonSystemFieldCount !== 1 ? "s" : ""}</span>

        <label style="display:flex;align-items:center;gap:5px;font-size:12px;color:${sec.managerOnly ? "var(--accent)" : "var(--text3)"};cursor:pointer;white-space:nowrap;"
          title="When enabled, submitters cannot see this section — only form managers and admins can view and fill it">
          <input type="checkbox" style="width:13px;height:13px;cursor:pointer;accent-color:var(--accent);" data-si="${si}"
            ${sec.managerOnly ? "checked" : ""}
            onchange="toggleManagerOnly(+this.dataset.si, this.checked)">
          Managers only
        </label>

        ${safeHtml(notifyToggle)}

        <button class="btn btn-ghost btn-sm btn-icon" data-si="${si}" onclick="removeSection(+this.dataset.si)" aria-label="Remove section">
          <svg width="14" height="14" viewBox="0 0 14 14" fill="currentColor"><path d="M1 1l12 12M13 1L1 13" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" fill="none"/></svg>
        </button>
      </div>

      ${safeHtml(notifyRow)}
      ${safeHtml(infoBar)}

      <div class="section-body">
        <div class="field-list" id="fields-${sec.id}">
          ${safeHtml(sec.fields.map((field, fi) => renderFieldItem(field, si, fi)).join(""))}
        </div>
        <button class="btn btn-ghost btn-sm" style="margin-top:10px;" data-si="${si}" onclick="openAddFieldModal(+this.dataset.si)">
          <svg width="12" height="12" viewBox="0 0 12 12" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M6 1v10M1 6h10"/></svg>
          Add Field
        </button>
      </div>
    </div>
  `;
}


function renderFieldItem(field, si, fi) {
  const typeLabel = CONFIG.FIELD_TYPES.find(t => t.value === field.type)?.label || field.type;
  const displayLabel = field.type === "InfoText"
    ? `ℹ ${(field.infoContent || "").replace(/<[^>]*>/g, "").slice(0, 50) || "Info block"}`
    : (field.label || "Untitled Field");

  // System fields (injected by toggleManagerOnly) are protected —
  // they cannot be edited or deleted by the form author.
  if (field.system) {
    return html`
      <div class="field-item" style="opacity:0.7;background:var(--bg3);">
        <span class="field-drag-handle" style="cursor:default;opacity:0.3;">
          <svg width="14" height="14" viewBox="0 0 14 14" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M3 4h8M3 7h8M3 10h8"/></svg>
        </span>
        <span class="field-type-badge" style="background:var(--bg2);color:var(--text3);">${typeLabel}</span>
        <div class="field-info">
          <div class="field-label" style="color:var(--text2);">${displayLabel}</div>
          <div class="field-meta"><span style="color:var(--text3);font-size:11px;">System field — auto-managed</span></div>
        </div>
        <div class="field-actions">
          <span title="System field — cannot be edited or deleted" style="color:var(--text3);padding:4px;">
            <svg width="13" height="13" viewBox="0 0 16 16" fill="currentColor"><path d="M11 7V5a3 3 0 10-6 0v2H3v8h10V7h-2zm-4-2a1 1 0 112 0v2H7V5zm1 6.5a1 1 0 110-2 1 1 0 010 2z"/></svg>
          </span>
        </div>
      </div>
    `;
  }

  return html`
    <div class="field-item">
      <span class="field-drag-handle">
        <svg width="14" height="14" viewBox="0 0 14 14" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M3 4h8M3 7h8M3 10h8"/></svg>
      </span>
      <span class="field-type-badge">${typeLabel}</span>
      <div class="field-info">
        <div class="field-label">${displayLabel}</div>
        <div class="field-meta">
          ${safeHtml(field.required ? `<span style="color:var(--red);font-size:11px;">Required</span>` : "")}
          ${safeHtml(field.description ? html`<span style="color:var(--text3);font-size:11px;">${field.description.slice(0,60)}…</span>` : "")}
        </div>
      </div>
      <div class="field-actions">
        <button class="btn btn-ghost btn-sm btn-icon" data-si="${si}" data-fi="${fi}" onclick="openEditFieldModal(+this.dataset.si,+this.dataset.fi)" aria-label="Edit field">
          <svg width="13" height="13" viewBox="0 0 13 13" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M9.5 1.5l2 2-8 8H1.5v-2l8-8z"/></svg>
        </button>
        <button class="btn btn-ghost btn-sm btn-icon" data-si="${si}" data-fi="${fi}" onclick="removeField(+this.dataset.si,+this.dataset.fi)" aria-label="Remove field">
          <svg width="13" height="13" viewBox="0 0 13 13" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M1 1l11 11M12 1L1 12"/></svg>
        </button>
      </div>
    </div>
  `;
}

function addSection() {
  const si = AppState.builderForm.sections.length;
  AppState.builderForm.sections.push({
    id:          "s" + Date.now(),
    title:       "",
    stepIndex:   si,
    managerOnly: false,
    notify:      false,
    deptEmail:   "",
    fields:      []
  });
  renderStepSections(document.getElementById("wizard-step-content"));
}

// =============================================================
// MANAGER SECTION SYSTEM FIELDS
// When a section is toggled to managerOnly, four protected system
// fields are automatically appended to it. They are marked
// system:true so renderFieldItem shows them as locked.
//
// Column names are derived via sectionKey() so they are unique
// per section and SharePoint-safe:
//   {key}_DeptEmail      — Text    — comma-separated notification emails
//   {key}_Completed      — Boolean — has this section been completed?
//   {key}_CompletedDate  — DateTime — when it was completed
//   {key}_CompletedBy    — Text    — display name of the completing user
//
// These are NOT in CONFIG.FIELD_TYPES (not user-selectable).
// They are provisioned separately in admin.js doCreateSharePointList.
// =============================================================
function buildSystemFields(section) {
  const key = sectionKey(section);
  return [
    {
      id:           `${section.id}_deptEmail`,
      label:        "Dept Notification Emails",
      type:         "Text",
      description:  "Comma-separated email addresses notified when this section is completed",
      required:     false,
      choices:      [],
      system:       true,
      systemRole:   "DeptEmail",
      internalName: `${key}_DeptEmail`,
    },
    {
      id:           `${section.id}_completed`,
      label:        "Completed",
      type:         "Boolean",
      description:  "Set automatically when the Complete button is clicked",
      required:     false,
      choices:      [],
      system:       true,
      systemRole:   "Completed",
      internalName: `${key}_Completed`,
    },
    {
      id:           `${section.id}_completedDate`,
      label:        "Completed Date",
      type:         "DateTime",
      description:  "Set automatically when the Complete button is clicked",
      required:     false,
      choices:      [],
      system:       true,
      systemRole:   "CompletedDate",
      internalName: `${key}_CompletedDate`,
    },
    {
      id:           `${section.id}_completedBy`,
      label:        "Completed By",
      type:         "Text",
      description:  "Display name of the user who clicked Complete",
      required:     false,
      choices:      [],
      system:       true,
      systemRole:   "CompletedBy",
      internalName: `${key}_CompletedBy`,
    },
  ];
}

function toggleManagerOnly(si, checked) {
  const section = AppState.builderForm.sections[si];
  section.managerOnly = checked;

  if (checked) {
    // Ensure notify and deptEmail exist on older section objects
    if (section.notify === undefined) section.notify    = false;
    if (section.deptEmail === undefined) section.deptEmail = "";
    // Remove stale system fields — they're only injected when notify is also on
    section.fields = section.fields.filter(f => !f.system);
    if (section.notify) {
      section.fields.push(...buildSystemFields(section));
    }
  } else {
    // Turning off managerOnly also resets notify — workflow makes no
    // sense on a public section
    section.notify    = false;
    section.deptEmail = "";
    section.fields    = section.fields.filter(f => !f.system);
  }

  renderStepSections(document.getElementById("wizard-step-content"));
}

// Toggling Notify on a managerOnly section injects or removes the
// system fields that support the completion workflow.
function toggleSectionNotify(si, checked) {
  const section = AppState.builderForm.sections[si];
  section.notify = checked;

  // Remove stale system fields first regardless of direction
  section.fields = section.fields.filter(f => !f.system);

  if (checked) {
    section.fields.push(...buildSystemFields(section));
  } else {
    // Turning notify off clears the email list — no orphaned config
    section.deptEmail = "";
  }

  renderStepSections(document.getElementById("wizard-step-content"));
}

function removeSection(si) {
  AppState.builderForm.sections.splice(si, 1);
  renderStepSections(document.getElementById("wizard-step-content"));
}

function removeField(si, fi) {
  const field = AppState.builderForm.sections[si]?.fields[fi];
  if (field?.system) return; // system fields cannot be removed manually
  AppState.builderForm.sections[si].fields.splice(fi, 1);
  renderStepSections(document.getElementById("wizard-step-content"));
}

function openAddFieldModal(sectionIndex) {
  openFieldModal(null, sectionIndex, null);
}

function openEditFieldModal(si, fi) {
  const field = AppState.builderForm.sections[si].fields[fi];
  openFieldModal(field, si, fi);
}

function openFieldModal(field, si, fi) {
  const isEdit = field !== null;
  const f = field || { label: "", type: "Text", description: "", required: false, choices: [] };

  // Check if a FileUpload field already exists in any section (only one allowed per form)
  const allFields = AppState.builderForm.sections.flatMap(s => s.fields || []);
  const existingFileUpload = allFields.find(fl => fl.type === "FileUpload" && fl.id !== f.id);

  openModal(html`
    <div class="modal-header">
      <span class="modal-title">${isEdit ? "Edit Field" : "Add Field"}</span>
      <button class="btn btn-ghost btn-icon btn-sm" onclick="closeModal()">
        <svg width="14" height="14" viewBox="0 0 14 14" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M1 1l12 12M13 1L1 13"/></svg>
      </button>
    </div>
    <div class="modal-body">
      <div class="form-group" id="field-label-group" style="${f.type === "InfoText" ? "display:none" : ""}">
        <label>Field Label *</label>
        <input id="field-label" class="input" placeholder="e.g. Full Name" value="${f.label}">
      </div>
      <div class="form-group">
        <label>Field Type *</label>
        <select id="field-type" class="select" onchange="toggleFieldOptions(this.value)">
          ${safeHtml(CONFIG.FIELD_TYPES.filter(t => t.value !== "RichText").map(t =>
            html`<option value="${t.value}" ${f.type === t.value ? "selected" : ""} ${t.value === "FileUpload" && existingFileUpload ? "disabled" : ""}>${t.label}${t.value === "FileUpload" && existingFileUpload ? " (already added)" : ""}</option>`
          ).join(""))}
        </select>
      </div>
      <div id="choice-options-block" style="${f.type === "Choice" || f.type === "MultiChoice" ? "" : "display:none"}">
        <div class="form-group">
          <label>Choices (one per line)</label>
          <textarea id="field-choices" class="textarea" placeholder="Option 1&#10;Option 2&#10;Option 3">${(f.choices||[]).join("\n")}</textarea>
        </div>
      </div>
      <div id="fileupload-options-block" style="${f.type === "FileUpload" ? "" : "display:none"}">
        <div class="form-group">
          <label>Accepted File Types</label>
          <input id="field-accept" class="input" placeholder="e.g. .pdf,.docx,.png (leave blank for any)"
            value="${f.accept || ""}">
          <span class="input-hint">Comma-separated extensions. Leave blank to allow any file type.</span>
        </div>
      </div>
      <div id="infotext-content-block" style="${f.type === "InfoText" ? "" : "display:none"}">
        <div class="form-group">
          <label>Notice Content (HTML supported)</label>
          <textarea id="field-infotext-content" class="textarea" style="min-height:120px;font-family:var(--mono);font-size:12.5px;" placeholder="e.g. &lt;p&gt;This form must be completed when...&lt;/p&gt;&lt;ul&gt;&lt;li&gt;Condition A&lt;/li&gt;&lt;/ul&gt;">${f.infoContent || ""}</textarea>
          <span class="input-hint">Rendered as read-only HTML in the form. Supports &lt;p&gt;, &lt;ul&gt;, &lt;li&gt;, &lt;strong&gt;, &lt;a&gt;, etc.</span>
        </div>
        <div class="form-group">
          <label>Style</label>
          <select id="field-infotext-style" class="select">
            <option value="info"    ${(f.infoStyle||"info")==="info"    ? "selected" : ""}>ℹ Info (blue)</option>
            <option value="warning" ${f.infoStyle==="warning" ? "selected" : ""}>⚠ Warning (amber)</option>
            <option value="success" ${f.infoStyle==="success" ? "selected" : ""}>✓ Success (green)</option>
            <option value="neutral" ${f.infoStyle==="neutral" ? "selected" : ""}>Neutral (grey)</option>
          </select>
        </div>
      </div>
      <div class="form-group" id="field-description-group" style="${f.type === "InfoText" ? "display:none" : ""}">
        <label>Description / Help Text</label>
        <input id="field-description" class="input" placeholder="Optional helper text shown below the field"
          value="${f.description || ""}">
      </div>
      <div class="flex items-center gap-2" id="field-required-group" style="${f.type === "InfoText" ? "display:none" : ""}">
        <input type="checkbox" id="field-required" class="toggle" ${f.required ? "checked" : ""}>
        <label for="field-required" style="font-size:13.5px;color:var(--text);cursor:pointer;">Required field</label>
      </div>
    </div>
    <div class="modal-footer">
      <button class="btn btn-ghost" onclick="closeModal()">Cancel</button>
      <button class="btn btn-primary" data-si="${si}" data-fi="${fi !== null ? fi : ""}" data-edit="${isEdit}" onclick="saveField(+this.dataset.si, this.dataset.fi===''?null:+this.dataset.fi, this.dataset.edit==='true')">
        ${isEdit ? "Update Field" : "Add Field"}
      </button>
    </div>
  `, true);
}


function toggleFieldOptions(type) {
  const isInfoText   = type === "InfoText";
  const isChoice     = type === "Choice" || type === "MultiChoice";
  const isFileUpload = type === "FileUpload";

  const show = (id, visible) => { const el = document.getElementById(id); if (el) el.style.display = visible ? "" : "none"; };
  show("choice-options-block",    isChoice);
  show("fileupload-options-block", isFileUpload);
  show("infotext-content-block",  isInfoText);
  show("field-label-group",       !isInfoText);
  show("field-description-group", !isInfoText);
  show("field-required-group",    !isInfoText);
}

// Keep legacy name as alias in case anything still calls it

function saveField(si, fi, isEdit) {
  const type = document.getElementById("field-type")?.value;
  if (!type) { showToast("error", "Field type is required"); return; }

  const isInfoText   = type === "InfoText";
  const isFileUpload = type === "FileUpload";
  const label        = isInfoText ? "" : (document.getElementById("field-label")?.value?.trim() || "");
  const description  = isInfoText ? "" : (document.getElementById("field-description")?.value?.trim() || "");
  const required     = isInfoText ? false : (document.getElementById("field-required")?.checked || false);
  const choicesRaw   = document.getElementById("field-choices")?.value || "";
  const choices      = choicesRaw.split("\n").map(c => c.trim()).filter(Boolean);
  const accept       = isFileUpload ? (document.getElementById("field-accept")?.value?.trim() || "") : "";
  const infoContent  = isInfoText
    ? (typeof DOMPurify !== "undefined"
        ? DOMPurify.sanitize(document.getElementById("field-infotext-content")?.value || "")
        : (document.getElementById("field-infotext-content")?.value || ""))
    : "";
  const infoStyle    = isInfoText ? (document.getElementById("field-infotext-style")?.value || "info") : "";

  if (!isInfoText && !label) { showToast("error", "Field label is required"); return; }
  if (isInfoText && !infoContent.trim()) { showToast("error", "Info content cannot be empty"); return; }

  const fieldObj = {
    id: (isEdit && AppState.builderForm.sections[si].fields[fi]?.id) || "f" + Date.now(),
    label, type, description, required, choices, accept, infoContent, infoStyle,
  };

  if (isEdit && fi !== null) {
    AppState.builderForm.sections[si].fields[fi] = fieldObj;
  } else {
    AppState.builderForm.sections[si].fields.push(fieldObj);
  }

  closeModal();
  renderStepSections(document.getElementById("wizard-step-content"));
  showToast("success", isEdit ? "Field updated" : "Field added");
}
// ---- Step 3: Conditions ----
function renderStepConditions(container) {
  const allFields = getAllFields();
  const { conditions } = AppState.builderForm;

  if (allFields.length < 2) {
    container.innerHTML = `<div class="card"><div class="card-body"><p style="color:var(--text2)">You need at least 2 fields to set up conditional visibility. Add more fields in the Sections & Fields step.</p></div></div>`;
    return;
  }

  container.innerHTML = html`
    <div class="card">
      <div class="card-header">
        <div>
          <div class="card-title">Conditional Visibility</div>
          <div class="card-subtitle">Show or hide fields/sections based on other field values</div>
        </div>
        <button class="btn btn-secondary btn-sm" onclick="addCondition()" style="margin-left:auto">Add Rule</button>
      </div>
      <div class="card-body" style="display:flex;flex-direction:column;gap:8px;" id="conditions-list">
        ${safeHtml(conditions.map((c, ci) => renderConditionRule(c, ci, allFields)).join(""))}
        ${safeHtml(!conditions.length ? `<p style="color:var(--text2);font-size:13.5px;">No conditions yet. Click "Add Rule" to create one.</p>` : "")}
      </div>
    </div>
  `;
}

function renderConditionRule(c, ci, allFields) {
  const fieldOpts = allFields.map(f =>
    html`<option value="${f.id}" ${c.showFieldId === f.id ? "selected" : ""} ${f.id === c.whenFieldId ? "disabled" : ""}>${f.label}</option>`
  ).join("");
  const triggerOpts = allFields.filter(f => f.type !== "InfoText" && f.type !== "FileUpload").map(f =>
    html`<option value="${f.id}" ${c.whenFieldId === f.id ? "selected" : ""} ${f.id === c.showFieldId ? "disabled" : ""}>${f.label}</option>`
  ).join("");

  // Determine the type of the trigger field so we can show the right operator set and value input
  const triggerField = allFields.find(f => f.id === c.whenFieldId);
  const triggerType = triggerField?.type || "Text";

  // Operators available per field type
  const numericTypes = ["Number", "Currency"];
  const dateTypes    = ["DateTime"];
  const boolTypes    = ["Boolean"];
  const choiceTypes  = ["Choice", "MultiChoice"];
  const textTypes    = ["Text", "Note", "URL"];

  let operators = [];
  if (numericTypes.includes(triggerType)) {
    operators = [
      { value: "eq",  label: "equals" },
      { value: "neq", label: "not equals" },
      { value: "gt",  label: "greater than" },
      { value: "gte", label: "greater than or equal to" },
      { value: "lt",  label: "less than" },
      { value: "lte", label: "less than or equal to" },
    ];
  } else if (dateTypes.includes(triggerType)) {
    operators = [
      { value: "eq",  label: "equals" },
      { value: "gt",  label: "after" },
      { value: "lt",  label: "before" },
    ];
  } else if (boolTypes.includes(triggerType) || choiceTypes.includes(triggerType)) {
    operators = [
      { value: "eq",  label: "equals" },
      { value: "neq", label: "not equals" },
    ];
  } else {
    operators = [
      { value: "eq",       label: "equals" },
      { value: "neq",      label: "not equals" },
      { value: "contains", label: "contains" },
    ];
  }


  const op = c.operator || "eq";

  // FIX: build operatorSelect and valueInput as plain string concatenation, not with the
  // html`` tag. The html`` tag auto-escapes all interpolated values, so using it here
  // caused the inner .map().join("") result (a string of <option> tags) to be HTML-escaped
  // a second time — turning <option> into &lt;option&gt; and leaving the select blank.
  // Plain string concatenation with explicit escAttr/escHtml calls is correct here;
  // safeHtml() in the outer html`` template (lines below) then passes them through unmodified.
  const operatorSelect =
    `<select class="select" style="min-width:110px;" data-ci="${ci}" onchange="AppState.builderForm.conditions[+this.dataset.ci].operator=this.value;renderStepConditions(document.getElementById('wizard-step-content'))">` +
    operators.map(o => `<option value="${escAttr(o.value)}"${op === o.value ? " selected" : ""}>${escHtml(o.label)}</option>`).join("") +
    `</select>`;

  // Value input — smart based on trigger field type
  let valueInput = "";
  if (boolTypes.includes(triggerType)) {
    const bVal = c.equalsValue === "true" || c.equalsValue === true;
    valueInput =
      `<select class="select" data-ci="${ci}" onchange="AppState.builderForm.conditions[+this.dataset.ci].equalsValue=this.value">` +
      `<option value="true"${bVal ? " selected" : ""}>Yes</option>` +
      `<option value="false"${!bVal ? " selected" : ""}>No</option>` +
      `</select>`;
  } else if (choiceTypes.includes(triggerType) && triggerField?.choices?.length) {
    valueInput =
      `<select class="select" data-ci="${ci}" onchange="AppState.builderForm.conditions[+this.dataset.ci].equalsValue=this.value">` +
      `<option value="">— choose —</option>` +
      triggerField.choices.map(ch => `<option value="${escAttr(ch)}"${c.equalsValue === ch ? " selected" : ""}>${escHtml(ch)}</option>`).join("") +
      `</select>`;
  } else if (numericTypes.includes(triggerType)) {
    valueInput = `<input class="input" type="number" step="any" placeholder="0" style="max-width:120px;" value="${escAttr(c.equalsValue||"")}" data-ci="${ci}" oninput="AppState.builderForm.conditions[+this.dataset.ci].equalsValue=this.value">`;
  } else if (dateTypes.includes(triggerType)) {
    valueInput = `<input class="input" type="date" style="max-width:160px;" value="${escAttr(c.equalsValue||"")}" data-ci="${ci}" oninput="AppState.builderForm.conditions[+this.dataset.ci].equalsValue=this.value">`;
  } else {
    valueInput = `<input class="input" placeholder="value" value="${escAttr(c.equalsValue||"")}" data-ci="${ci}" oninput="AppState.builderForm.conditions[+this.dataset.ci].equalsValue=this.value">`;
  }


  return html`
    <div class="condition-rule" style="flex-wrap:wrap;gap:8px;">
      <span style="font-size:12px;color:var(--text3);white-space:nowrap;align-self:center;">Show</span>
      <select class="select" data-ci="${ci}" onchange="setConditionShow(+this.dataset.ci,this.value)"><option value="">— field —</option>${safeHtml(fieldOpts)}</select>
      <span style="font-size:12px;color:var(--text3);white-space:nowrap;align-self:center;">when</span>
      <select class="select" data-ci="${ci}" onchange="setConditionWhen(+this.dataset.ci,this.value)"><option value="">— trigger field —</option>${safeHtml(triggerOpts)}</select>
      ${safeHtml(operatorSelect)}
      ${safeHtml(valueInput)}
      <button class="btn btn-danger btn-sm btn-icon" data-ci="${ci}" onclick="removeCondition(+this.dataset.ci)">
        <svg width="13" height="13" viewBox="0 0 13 13" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M1 1l11 11M12 1L1 12"/></svg>
      </button>
    </div>
  `;
}

function setConditionShow(ci, value) {
  // Prevent same field for show and when
  if (value && value === AppState.builderForm.conditions[ci].whenFieldId) {
    showToast("error", "Show and When fields cannot be the same");
    renderStepConditions(document.getElementById("wizard-step-content"));
    return;
  }
  AppState.builderForm.conditions[ci].showFieldId = value;
  renderStepConditions(document.getElementById("wizard-step-content"));
}

function setConditionWhen(ci, value) {
  // Prevent same field for show and when
  if (value && value === AppState.builderForm.conditions[ci].showFieldId) {
    showToast("error", "Show and When fields cannot be the same");
    renderStepConditions(document.getElementById("wizard-step-content"));
    return;
  }
  AppState.builderForm.conditions[ci].whenFieldId = value;
  renderStepConditions(document.getElementById("wizard-step-content"));
}

function addCondition() {
  AppState.builderForm.conditions.push({ showFieldId: "", whenFieldId: "", operator: "eq", equalsValue: "" });
  renderStepConditions(document.getElementById("wizard-step-content"));
}

function removeCondition(ci) {
  AppState.builderForm.conditions.splice(ci, 1);
  renderStepConditions(document.getElementById("wizard-step-content"));
}
// ---- Step 4: Dependent Dropdowns ----
function renderStepDependents(container) {
  const choiceFields = getAllFields().filter(f => f.type === "Choice");
  const { dependentDropdowns } = AppState.builderForm;

  if (choiceFields.length < 2) {
    container.innerHTML = `<div class="card"><div class="card-body"><p style="color:var(--text2)">You need at least 2 Choice fields to set up dependent dropdowns.</p></div></div>`;
    return;
  }

  container.innerHTML = html`
    <div class="card">
      <div class="card-header">
        <div>
          <div class="card-title">Dependent Dropdowns</div>
          <div class="card-subtitle">Set child dropdown options based on the selected parent value</div>
        </div>
        <button class="btn btn-secondary btn-sm" onclick="addDependentDropdown()" style="margin-left:auto">Add Mapping</button>
      </div>
      <div class="card-body" id="dependents-list">
        ${safeHtml(dependentDropdowns.map((dd, di) => renderDependentBlock(dd, di, choiceFields)).join(""))}
        ${safeHtml(!dependentDropdowns.length ? `<p style="color:var(--text2);font-size:13.5px;">No dependent dropdowns configured.</p>` : "")}
      </div>
    </div>
  `;
}

function renderDependentBlock(dd, di, choiceFields) {
  const parentField = choiceFields.find(f => f.id === dd.parentFieldId);
  const parentChoices = parentField ? parentField.choices || [] : [];
  const opts = choiceFields.map(f => html`<option value="${f.id}" ${dd.parentFieldId === f.id ? "selected" : ""}>${f.label}</option>`).join("");
  const childOpts = choiceFields.map(f => html`<option value="${f.id}" ${dd.childFieldId === f.id ? "selected" : ""}>${f.label}</option>`).join("");

  return html`
    <div style="background:var(--bg3);border:1px solid var(--border);border-radius:var(--radius-sm);padding:16px;margin-bottom:12px;">
      <div class="flex items-center gap-3 mb-4">
        <div class="form-group" style="flex:1">
          <label>Parent Field</label>
          <select class="select" data-di="${di}" onchange="setDependentParent(+this.dataset.di,this.value)">
            <option value="">Select parent field</option>${safeHtml(opts)}
          </select>
        </div>
        <div style="padding-top:20px;color:var(--text3)">→</div>
        <div class="form-group" style="flex:1">
          <label>Child Field</label>
          <select class="select" data-di="${di}" onchange="AppState.builderForm.dependentDropdowns[+this.dataset.di].childFieldId=this.value">
            <option value="">Select child field</option>${safeHtml(childOpts)}
          </select>
        </div>
        <button class="btn btn-danger btn-sm btn-icon" style="margin-top:16px;" data-di="${di}" onclick="removeDependentDropdown(+this.dataset.di)">
          <svg width="13" height="13" viewBox="0 0 13 13" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M1 1l11 11M12 1L1 12"/></svg>
        </button>
      </div>
      ${safeHtml(parentChoices.length ? html`
        <div>
          <label style="font-size:11px;text-transform:uppercase;letter-spacing:0.06em;color:var(--text3);">Value Mappings</label>
          <div style="margin-top:8px;display:flex;flex-direction:column;gap:8px;">
            ${safeHtml(parentChoices.map(choice => html`
              <div class="flex items-center gap-2">
                <span style="font-size:12.5px;min-width:120px;color:var(--text2);">${choice}</span>
                <span style="color:var(--text3)">→</span>
                <input class="input" placeholder="Child options (comma-separated)"
                  value="${(dd.mapping && dd.mapping[choice] ? dd.mapping[choice] : []).join(", ")}"
                  data-di="${di}" data-choice="${choice}"
                  oninput="setDependentMapping(+this.dataset.di, this.dataset.choice, this.value)">
              </div>
            `).join(""))}
          </div>
        </div>
      ` : `<p style="font-size:12.5px;color:var(--text3);">Select a parent Choice field with options to configure mappings.</p>`)}
    </div>
  `;
}

function addDependentDropdown() {
  AppState.builderForm.dependentDropdowns.push({ parentFieldId: "", childFieldId: "", mapping: {} });
  renderStepDependents(document.getElementById("wizard-step-content"));
}

function removeDependentDropdown(di) {
  AppState.builderForm.dependentDropdowns.splice(di, 1);
  renderStepDependents(document.getElementById("wizard-step-content"));
}

function setDependentParent(di, parentId) {
  AppState.builderForm.dependentDropdowns[di].parentFieldId = parentId;
  AppState.builderForm.dependentDropdowns[di].mapping = {};
  const choiceFields = getAllFields().filter(f => f.type === "Choice");
  renderStepDependents(document.getElementById("wizard-step-content"));
}

function setDependentMapping(di, parentVal, rawVal) {
  if (!AppState.builderForm.dependentDropdowns[di].mapping) AppState.builderForm.dependentDropdowns[di].mapping = {};
  AppState.builderForm.dependentDropdowns[di].mapping[parentVal] = rawVal.split(",").map(v => v.trim()).filter(Boolean);
}
// Auto-assign incremental stepIndexes if all sections are on the same step
// (happens when switching to multistep for the first time, or when sections
// were created before the per-section stepIndex defaulting was fixed)
function autoAssignSteps() {
  const sections = AppState.builderForm.sections;
  if (!sections.length) return;
  const allSame = sections.every(s => (s.stepIndex || 0) === (sections[0].stepIndex || 0));
  if (allSame) {
    sections.forEach((s, i) => { s.stepIndex = i; });
  }
}

// ---- Step 5: Layout ----
function renderStepLayout(container) {
  const { layout, sections } = AppState.builderForm;

  // Auto-assign incremental stepIndexes if all sections share the same value —
  // this covers both first-time multistep selection AND editing existing forms
  // where sections were saved with stepIndex:0 before the fix was in place.
  if (layout === "multistep") autoAssignSteps();

  container.innerHTML = html`
    <div class="card">
      <div class="card-header">
        <div>
          <div class="card-title">Layout Preference</div>
          <div class="card-subtitle">Choose how the form is presented to users</div>
        </div>
      </div>
      <div class="card-body" style="display:flex;flex-direction:column;gap:20px;">
        <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px;">
          <div onclick="AppState.builderForm.layout='single';renderStepLayout(document.getElementById('wizard-step-content'))"
            style="border:2px solid ${layout==='single'?'var(--accent)':'var(--border)'};border-radius:var(--radius);padding:20px;cursor:pointer;background:${layout==='single'?'rgba(79,124,255,0.06)':'var(--bg3)'}">
            <div style="font-weight:600;margin-bottom:4px;">Single Page</div>
            <div style="font-size:12.5px;color:var(--text2);">All sections shown at once. Good for short forms.</div>
          </div>
          <div onclick="AppState.builderForm.layout='multistep';autoAssignSteps();renderStepLayout(document.getElementById('wizard-step-content'))"
            style="border:2px solid ${layout==='multistep'?'var(--accent)':'var(--border)'};border-radius:var(--radius);padding:20px;cursor:pointer;background:${layout==='multistep'?'rgba(79,124,255,0.06)':'var(--bg3)'}">
            <div style="font-weight:600;margin-bottom:4px;">Multi-Step Wizard</div>
            <div style="font-size:12.5px;color:var(--text2);">Sections split across steps with Previous / Next navigation.</div>
          </div>
        </div>

        ${safeHtml(layout === "multistep" && sections.length ? html`
          <div>
            <label style="display:block;margin-bottom:12px;">Assign sections to steps:</label>
            <div style="display:flex;flex-direction:column;gap:8px;">
              ${safeHtml(sections.map((sec, si) => html`
                <div class="flex items-center gap-3" style="background:var(--bg3);border:1px solid var(--border);border-radius:var(--radius-sm);padding:12px 14px;">
                  <span style="flex:1;font-size:13.5px;">${sec.title || "Section " + (si+1)}</span>
                  <label style="font-size:12px;">Step</label>
                  <input type="number" class="input" min="1" max="10" style="width:64px;"
                    value="${(sec.stepIndex||0)+1}"
                    data-si="${si}"
                    onchange="AppState.builderForm.sections[+this.dataset.si].stepIndex=Math.max(0,parseInt(this.value)-1)||0">
                </div>
              `).join(""))}
            </div>
          </div>
        ` : "")}
      </div>
    </div>
  `;
}
// ---- Step 6: Access ----
async function renderStepAccess(container) {
  const { access, specificPeople, submissionType, formManagers } = AppState.builderForm;

  container.innerHTML = html`
    <div style="display:flex;flex-direction:column;gap:16px;">
      <div class="card">
        <div class="card-header">
          <div>
            <div class="card-title">Access Permissions</div>
            <div class="card-subtitle">Who can access and submit this form?</div>
          </div>
        </div>
        <div class="card-body" style="display:flex;flex-direction:column;gap:12px;">
          ${safeHtml(CONFIG.ACCESS_OPTIONS.map(opt => html`
            <div data-val="${opt.value}" onclick="setAccess(this.dataset.val)"
              style="border:2px solid ${access===opt.value?'var(--accent)':'var(--border)'};border-radius:var(--radius-sm);padding:14px 16px;cursor:pointer;background:${access===opt.value?'rgba(79,124,255,0.06)':'var(--bg3)'}">
              <div style="font-weight:500;">${opt.label}</div>
            </div>
          `).join(""))}

          ${safeHtml(access === "Specific" ? html`
            <div id="people-picker-section">
              <label style="display:block;margin-bottom:8px;">Search and add people:</label>
              <div class="flex gap-2">
                <input id="people-search" class="input" placeholder="Search by name or email…" oninput="debouncedPeopleSearch(this.value)">
                <button class="btn btn-secondary" onclick="searchPeopleNow()">Search</button>
              </div>
              <div id="people-results" style="margin-top:8px;"></div>
              <div id="selected-people" style="margin-top:12px;display:flex;flex-wrap:wrap;gap:6px;">
                ${safeHtml(specificPeople.map(p => html`
                  <span class="person-chip">
                    <div class="avatar" style="width:20px;height:20px;font-size:9px;">${p.displayName.split(" ").map(n=>n[0]).join("").slice(0,2)}</div>
                    ${p.displayName}
                    <button data-id="${p.id}" onclick="removePerson(this.dataset.id)">
                      <svg width="10" height="10" viewBox="0 0 10 10" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M1 1l8 8M9 1L1 9"/></svg>
                    </button>
                  </span>
                `).join(""))}
              </div>
            </div>
          ` : "")}
        </div>
      </div>

      <!-- Form Managers -->
      <div class="card">
        <div class="card-header">
          <div>
            <div class="card-title">Form Managers</div>
            <div class="card-subtitle">Colleagues who can view and manage all submissions for this form</div>
          </div>
        </div>
        <div class="card-body">
          <div class="flex gap-2" style="margin-bottom:8px;">
            <input id="managers-search" class="input" placeholder="Search by name or email…" oninput="debouncedManagerSearch(this.value)">
            <button class="btn btn-secondary" onclick="searchManagersNow()">Search</button>
          </div>
          <div id="managers-results" style="margin-top:8px;"></div>
          <div id="selected-managers" style="margin-top:12px;display:flex;flex-wrap:wrap;gap:6px;">
            ${safeHtml(formManagers.map(p => html`
              <span class="person-chip">
                <div class="avatar" style="width:20px;height:20px;font-size:9px;">${p.displayName.split(" ").map(n=>n[0]).join("").slice(0,2)}</div>
                ${p.displayName}
                <button data-id="${p.id}" onclick="removeManager(this.dataset.id)">
                  <svg width="10" height="10" viewBox="0 0 10 10" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M1 1l8 8M9 1L1 9"/></svg>
                </button>
              </span>
            `).join(""))}
          </div>
          <p style="font-size:12px;color:var(--text3);margin-top:8px;">Form Managers can view, edit and soft-delete all submissions. Standard submitters can only see their own.</p>
        </div>
      </div>

      <div class="card">
        <div class="card-header">
          <div>
            <div class="card-title">Submission Type</div>
            <div class="card-subtitle">Can users edit their submission after submitting?</div>
          </div>
        </div>
        <div class="card-body" style="display:flex;flex-direction:column;gap:12px;">
          ${safeHtml(CONFIG.SUBMISSION_TYPES.map(opt => html`
            <div data-val="${opt.value}" onclick="AppState.builderForm.submissionType=this.dataset.val;renderStepAccess(document.getElementById('wizard-step-content'))"
              style="border:2px solid ${submissionType===opt.value?'var(--accent)':'var(--border)'};border-radius:var(--radius-sm);padding:14px 16px;cursor:pointer;background:${submissionType===opt.value?'rgba(79,124,255,0.06)':'var(--bg3)'}">
              <div style="font-weight:500;">${opt.label}</div>
              <div style="font-size:12.5px;color:var(--text2);margin-top:2px;">${opt.desc}</div>
            </div>
          `).join(""))}
        </div>
      </div>
    </div>
  `;
}

function setAccess(value) {
  AppState.builderForm.access = value;
  renderStepAccess(document.getElementById("wizard-step-content"));
}

function removePerson(id) {
  AppState.builderForm.specificPeople = AppState.builderForm.specificPeople.filter(p => p.id !== id);
  renderStepAccess(document.getElementById("wizard-step-content"));
}

// Specific people search (Access step — who can submit this form)
const _peopleSearch = createPeopleSearch({
  inputId:   "people-search",
  resultsId: "people-results",
  onClickFn: "addPersonFromEl",
});
function debouncedPeopleSearch(val) { _peopleSearch.debounced(val); }
async function searchPeopleNow(q)   { await _peopleSearch.search(q); }

function removeManager(id) {
  AppState.builderForm.formManagers = AppState.builderForm.formManagers.filter(p => p.id !== id);
  renderStepAccess(document.getElementById("wizard-step-content"));
}

// Form managers search (Access step — who can manage submissions)
const _managersSearch = createPeopleSearch({
  inputId:   "managers-search",
  resultsId: "managers-results",
  onClickFn: "addManagerFromEl",
});
function debouncedManagerSearch(val) { _managersSearch.debounced(val); }
async function searchManagersNow(q)  { await _managersSearch.search(q); }

function addManagerFromEl(el) {
  addManager(el.dataset.id, el.dataset.name, el.dataset.email);
}

function addManager(id, displayName, email) {
  if (!AppState.builderForm.formManagers.find(p => p.id === id)) {
    AppState.builderForm.formManagers.push({ id, displayName, email });
  }
  renderStepAccess(document.getElementById("wizard-step-content"));
}

function addPersonFromEl(el) {
  addPerson(el.dataset.id, el.dataset.name, el.dataset.email);
}

function addPerson(id, displayName, email) {
  if (!AppState.builderForm.specificPeople.find(p => p.id === id)) {
    AppState.builderForm.specificPeople.push({ id, displayName, email });
  }
  renderStepAccess(document.getElementById("wizard-step-content"));
}
function renderStepOnSubmit(container) {
  const { submitNotifyEmails, notifySubmitter } = AppState.builderForm;
  container.innerHTML = html`
    <div class="card">
      <div class="card-header">
        <div>
          <div class="card-title">On Submit</div>
          <div class="card-subtitle">Configure what happens when a user submits this form</div>
        </div>
      </div>
      <div class="card-body" style="display:flex;flex-direction:column;gap:24px;">

        <!-- Submitter confirmation -->
        <div class="form-group">
          <label style="display:flex;align-items:center;gap:10px;cursor:pointer;font-weight:500;">
            <input type="checkbox" id="onsubmit-notify-submitter"
              style="width:15px;height:15px;accent-color:var(--accent);cursor:pointer;"
              ${notifySubmitter ? "checked" : ""}
              onchange="AppState.builderForm.notifySubmitter = this.checked">
            Send submitter a confirmation email
          </label>
          <div style="font-size:12.5px;color:var(--text2);margin-top:4px;padding-left:25px;">
            The submitter receives a read-only HTML summary of their submission. System fields are excluded.
          </div>
        </div>

        <!-- Notify addresses -->
        <div class="form-group">
          <label for="onsubmit-notify-emails" style="color:var(--text1);font-weight:500;">Notification emails</label>
          <input id="onsubmit-notify-emails" class="input" type="text"
            placeholder="e.g. admin@example.com, team@example.com"
            value="${escAttr(submitNotifyEmails)}"
            oninput="AppState.builderForm.submitNotifyEmails = this.value.trim()">
          <div style="font-size:12px;color:var(--text3);margin-top:4px;">
            Comma-separated. Leave blank to skip. These addresses are notified every time the form is submitted.
          </div>
        </div>

      </div>
    </div>
  `;
}

function renderStepReview(container) {
  const { title, listName, sections, layout, access, specificPeople, formManagers, submissionType, conditions, dependentDropdowns, governance: g, submitNotifyEmails, notifySubmitter } = AppState.builderForm;
  const allFields = getAllFields();

  // Human-readable labels for governance select values
  const retentionLabels   = { under1: "< 1 year", "1to3": "1–3 years", "3to7": "3–7 years", indefinite: "Indefinite" };
  const sensitiveLabels   = { none: "No sensitive data", personal: "Special category personal data", commercial: "Commercially sensitive", both: "Personal & commercial" };
  const privacyLabels     = { yes: "Yes — completed", no: "No — not done", na: "Not applicable" };
  const externalLabels    = { none: "Internal only", recipients: "Data shared externally", submitters: "External submitters" };
  const volumeLabels      = { low: "Low (< 100/month)", medium: "Medium (100–1,000/month)", high: "High (> 1,000/month)" };

  container.innerHTML = html`
    <div style="display:flex;flex-direction:column;gap:16px;">
      <div class="card">
        <div class="card-header"><div class="card-title">Review & Confirm</div></div>
        <div class="card-body">
          <div class="grid-2" style="gap:24px;">
            <div>
              <div style="font-size:11px;text-transform:uppercase;letter-spacing:0.06em;color:var(--text3);margin-bottom:4px;">Form Name</div>
              <div style="font-weight:500;">${title}</div>
            </div>
            <div>
              <div style="font-size:11px;text-transform:uppercase;letter-spacing:0.06em;color:var(--text3);margin-bottom:4px;">Layout</div>
              <div>${layout === "multistep" ? "Multi-Step Wizard" : "Single Page"}</div>
            </div>
            <div>
              <div style="font-size:11px;text-transform:uppercase;letter-spacing:0.06em;color:var(--text3);margin-bottom:4px;">Submission Type</div>
              <div>${submissionType === "SubmitEdit" ? "Submit & Edit" : "Submit Only"}</div>
            </div>
            <div>
              <div style="font-size:11px;text-transform:uppercase;letter-spacing:0.06em;color:var(--text3);margin-bottom:4px;">Access</div>
              <div>${CONFIG.ACCESS_OPTIONS.find(a=>a.value===access)?.label || access}
                ${safeHtml(specificPeople.length ? `(${specificPeople.length} people)` : "")}
              </div>
            </div>
            <div>
              <div style="font-size:11px;text-transform:uppercase;letter-spacing:0.06em;color:var(--text3);margin-bottom:4px;">Form Managers</div>
              <div>${formManagers.length ? `${formManagers.length} manager${formManagers.length !== 1 ? "s" : ""}` : "None (author only)"}</div>
            </div>
            <div>
              <div style="font-size:11px;text-transform:uppercase;letter-spacing:0.06em;color:var(--text3);margin-bottom:4px;">Total Fields</div>
              <div>${allFields.length} across ${sections.length} section${sections.length !== 1 ? "s" : ""}</div>
            </div>
          </div>
        </div>
      </div>

      <!-- Governance summary -->
      <div class="card">
        <div class="card-header"><div class="card-title">Governance</div></div>
        <div class="card-body">
          <div class="grid-2" style="gap:24px;">
            <div>
              <div style="font-size:11px;text-transform:uppercase;letter-spacing:0.06em;color:var(--text3);margin-bottom:4px;">Existing Process</div>
              <div>${g.existingProcess === "yes" ? "Yes" : g.existingProcess === "no" ? "No" : "—"}${safeHtml(g.existingProcess === "yes" && g.existingProcessDetail ? html`<div style="font-size:12px;color:var(--text2);margin-top:2px;">${g.existingProcessDetail}</div>` : "")}</div>
            </div>
            <div>
              <div style="font-size:11px;text-transform:uppercase;letter-spacing:0.06em;color:var(--text3);margin-bottom:4px;">Data Retention</div>
              <div>${retentionLabels[g.retention] || "—"}</div>
            </div>
            <div>
              <div style="font-size:11px;text-transform:uppercase;letter-spacing:0.06em;color:var(--text3);margin-bottom:4px;">Data Sensitivity</div>
              <div style="${(g.sensitiveData === "personal" || g.sensitiveData === "both") ? "color:var(--amber,#d97706);font-weight:500;" : ""}">${sensitiveLabels[g.sensitiveData] || "—"}</div>
            </div>
            <div>
              <div style="font-size:11px;text-transform:uppercase;letter-spacing:0.06em;color:var(--text3);margin-bottom:4px;">Privacy Assessment</div>
              <div style="${g.privacyAssessment === "no" ? "color:var(--amber,#d97706);font-weight:500;" : ""}">${privacyLabels[g.privacyAssessment] || "—"}</div>
            </div>
            <div>
              <div style="font-size:11px;text-transform:uppercase;letter-spacing:0.06em;color:var(--text3);margin-bottom:4px;">External Access</div>
              <div>${externalLabels[g.externalAccess] || "—"}</div>
            </div>
            <div>
              <div style="font-size:11px;text-transform:uppercase;letter-spacing:0.06em;color:var(--text3);margin-bottom:4px;">Expected Volume</div>
              <div>${volumeLabels[g.expectedVolume] || "—"}</div>
            </div>
            ${safeHtml(g.dataOwner ? html`
            <div>
              <div style="font-size:11px;text-transform:uppercase;letter-spacing:0.06em;color:var(--text3);margin-bottom:4px;">Data Owner</div>
              <div>${g.dataOwner.displayName}</div>
            </div>` : "")}
            ${safeHtml(g.continuityPlan ? html`
            <div style="grid-column:1/-1;">
              <div style="font-size:11px;text-transform:uppercase;letter-spacing:0.06em;color:var(--text3);margin-bottom:4px;">Continuity Plan</div>
              <div style="font-size:13.5px;color:var(--text2);">${g.continuityPlan}</div>
            </div>` : "")}
          </div>
        </div>
      </div>

      <!-- On Submit summary -->
      <div class="card">
        <div class="card-header"><div class="card-title">On Submit</div></div>
        <div class="card-body">
          <div class="grid-2" style="gap:24px;">
            <div>
              <div style="font-size:11px;text-transform:uppercase;letter-spacing:0.06em;color:var(--text3);margin-bottom:4px;">Submitter Confirmation</div>
              <div>${notifySubmitter ? "✓ Confirmation email enabled" : "Disabled"}</div>
            </div>
            <div>
              <div style="font-size:11px;text-transform:uppercase;letter-spacing:0.06em;color:var(--text3);margin-bottom:4px;">Notify on Submit</div>
              <div>${submitNotifyEmails || "—"}</div>
            </div>
          </div>
        </div>
      </div>

      <!-- Preview -->
      <div class="card">
        <div class="card-header">
          <div class="card-title">Form Preview</div>
        </div>
        <div class="card-body">
          ${safeHtml(renderFormPreviewHtml())}
        </div>
      </div>
    </div>
  `;
}
// =============================================================
// FORM PREVIEW
// =============================================================
function renderFormPreviewHtml() {
  const { title, sections, layout } = AppState.builderForm;
  if (!sections.length) return `<p style="color:var(--text3);">No sections defined.</p>`;

  return html`
    <div class="form-preview">
      <h2>${title}</h2>
      ${safeHtml(layout === "multistep" ? html`<p style="color:var(--text2);font-size:13px;margin-top:4px;">Multi-step form — ${sections.length} step${sections.length!==1?"s":""}</p>` : "")}
      <div style="margin-top:24px;">
        ${safeHtml(sections.map(sec => html`
          <div class="preview-section" style="${sec.managerOnly ? "border:1px dashed var(--accent);border-radius:var(--radius-sm);padding:12px;" : ""}">
            ${safeHtml(sec.managerOnly ? `
              <div style="font-size:11px;color:var(--accent);font-weight:600;text-transform:uppercase;letter-spacing:0.06em;margin-bottom:8px;display:flex;align-items:center;gap:4px;">
                <svg width="11" height="11" viewBox="0 0 16 16" fill="currentColor"><path d="M7.122.392a1.75 1.75 0 011.756 0l5.25 3.045c.54.313.872.89.872 1.514V8.64c0 2.048-1.19 3.914-3.05 4.856l-2.5 1.286a1.75 1.75 0 01-1.6 0l-2.5-1.286C3.19 12.554 2 10.688 2 8.64V4.951c0-.624.332-1.2.872-1.514L7.122.392z"/></svg>
                Managers only
              </div>
            ` : "")}
            ${safeHtml(sec.title ? html`<div class="preview-section-title">${sec.title}</div>` : "")}
            ${safeHtml(sec.fields.map(field => renderPreviewField(field)).join(""))}
          </div>
        `).join(""))}
      </div>
      ${safeHtml(layout === "multistep" ? `
        <div class="flex gap-2" style="margin-top:16px;padding-top:16px;border-top:1px solid var(--border);">
          <button class="btn btn-secondary" disabled>← Previous</button>
          <button class="btn btn-primary">Next →</button>
        </div>
      ` : `<button class="btn btn-primary" style="margin-top:16px;" disabled>Submit</button>`)}
    </div>
  `;
}

function renderPreviewField(field) {
  if (field.type === "InfoText") {
    const safeStyle = ["info","warning","success","neutral"].includes(field.infoStyle) ? field.infoStyle : "info";
    // InfoText deliberately renders user-authored HTML — sanitise it
    const cleanContent = typeof DOMPurify !== "undefined"
      ? DOMPurify.sanitize(field.infoContent || "")
      : escHtml(field.infoContent || "");
    return html`<div class="infotext-block infotext-${safeStyle}">${safeHtml(cleanContent)}</div>`;
  }
  const typeLabel = CONFIG.FIELD_TYPES.find(t => t.value === field.type)?.label || field.type;
  let inputHtml = "";
  switch (field.type) {
    case "Note":
    case "RichText":
      inputHtml = `<div class="preview-input-mock" style="min-height:72px;">Enter text…</div>`; break;
    case "Boolean":
      inputHtml = `<div class="flex items-center gap-2"><input type="checkbox" class="toggle" disabled><span style="font-size:13px;color:var(--text3);">Yes / No</span></div>`; break;
    case "Choice":
    case "MultiChoice":
      inputHtml = html`<div class="preview-input-mock">${field.choices?.length ? field.choices[0] + " ▾" : "Select an option ▾"}</div>`; break;
    case "Text":
      inputHtml = `<div class="preview-input-mock">${field.description || "Enter " + (field.label || "text").toLowerCase() + "…"}</div>`; break;
    case "Number":
      inputHtml = `<div class="preview-input-mock" style="max-width:200px;font-variant-numeric:tabular-nums;">0</div>`; break;
    case "DateTime":
      inputHtml = `<div class="preview-input-mock">DD/MM/YYYY</div>`; break;
    case "User":
      inputHtml = `<div class="preview-input-mock">Search people…</div>`; break;
    case "Currency":
      inputHtml = `<div class="preview-input-mock" style="display:inline-flex;align-items:center;gap:4px;max-width:200px;font-variant-numeric:tabular-nums;"><span style="color:var(--text3);">£</span><span>0.00</span></div>`; break;
    case "URL":
      inputHtml = `<div class="preview-input-mock">https://…</div>`; break;
    case "FileUpload":
      inputHtml = html`<div class="preview-input-mock" style="display:flex;align-items:center;gap:8px;">
        <svg width="14" height="14" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M8 10V3M5 6l3-3 3 3M3 13h10"/></svg>
        Choose files${field.accept ? ` (${field.accept})` : ""}
      </div>`; break;
    default:
      inputHtml = html`<div class="preview-input-mock">${field.description || "Enter " + field.label.toLowerCase() + "…"}</div>`;
  }

  return html`
    <div class="preview-field">
      <label class="preview-label">${field.label}${safeHtml(field.required ? '<span class="req">*</span>' : "")}</label>
      ${safeHtml(field.description ? html`<span style="font-size:11.5px;color:var(--text3);display:block;margin-bottom:5px;">${field.description}</span>` : "")}
      ${safeHtml(inputHtml)}
    </div>
  `;
}

// =============================================================
// LOAD FORM INTO BUILDER
// Single shared function that loads a form definition from
// SharePoint into AppState.builderForm. Used by:
//   - doEditFormRequest (admin.js) — after its cleanup steps
//   - previewRequest (builder.js) — jumps to the review step
//
// Replaces the duplicated editRequest (builder.js) and the
// inline loading block in doEditFormRequest (admin.js) that
// was missing governance, causing it to be lost on edit.
// =============================================================
async function loadFormIntoBuilder(itemId) {
  resetBuilderForm();
  AppState.builderMode   = "edit";
  AppState.builderItemId = itemId;

  const def = await getFormDefinition(CONFIG.FORMS_LIST, itemId);
  if (def) {
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
    if (def.governance) {
      Object.assign(AppState.builderForm.governance, def.governance);
    }
  }
  return def;
}

function previewRequest(itemId) {
  // Load the form into the builder then jump straight to the Review step
  loadFormIntoBuilder(itemId).then(() => {
    AppState.builderStep = WIZARD_STEPS.findIndex(s => s.key === "review");
    renderBuilder(document.getElementById("main-content"));
  });
}
// =============================================================
// SAVE / SUBMIT
// =============================================================
async function saveBuilderDraft() {
  // Silent auto-save — no progress modal (doSubmitForReview adds its own)
  try {
    const form = AppState.builderForm;
    const jsonDef = {
      title: form.title,
      listName: form.listName,
      sections: form.sections,
      layout: form.layout,
      access: form.access,
      specificPeople: form.specificPeople,
      formManagers: form.formManagers,
      submissionType: form.submissionType,
      conditions: form.conditions,
      dependentDropdowns: form.dependentDropdowns,
      governance: form.governance,
    };

    // SP columns written: Title, Status (new only), ListName, FormDefinition (via uploadJsonAttachment)
    // Governance columns promoted to SP for filtering/reporting by admins
    const gov = form.governance;
    const govFields = {
      [CONFIG.COL_GOV_RETENTION]:    gov.retention       || "",
      [CONFIG.COL_GOV_SENSITIVE]:    gov.sensitiveData   || "",
      [CONFIG.COL_GOV_PRIVACY]:      gov.privacyAssessment || "",
      [CONFIG.COL_GOV_EXTERNAL]:     gov.externalAccess  || "",
      [CONFIG.COL_GOV_VOLUME]:       gov.expectedVolume  || "",
    };

    // Data owner is a Person column — resolve their SP integer user ID first
    if (gov.dataOwner?.email) {
      try {
        const spUserId = await resolveSpUserId(gov.dataOwner.email);
        if (spUserId) {
          govFields[CONFIG.COL_GOV_DATA_OWNER + "LookupId"] = spUserId;
        }
      } catch (e) {
        console.warn("[saveBuilderDraft] Could not resolve SP user ID for data owner:", e.message);
      }
    }

    const fields = {
      Title:               form.title || "Untitled",
      [CONFIG.COL_LISTNAME]: form.listName || generateListName(form.title),
      ...govFields,
    };

    let itemId = AppState.builderItemId;
    if (itemId) {
      // Editing existing — do NOT touch Status
      await updateListItem(CONFIG.FORMS_LIST, itemId, fields);
    } else {
      // New form — set initial status to Created
      const result = await createListItem(CONFIG.FORMS_LIST, {
        ...fields,
        [CONFIG.COL_STATUS]: "Created",
      });
      itemId = result.id;
      AppState.builderItemId = itemId;
    }

    await uploadJsonAttachment(CONFIG.FORMS_LIST, itemId, "form-definition.json", jsonDef);

  } catch (e) {
    throw e; // Let caller handle errors
  }
}

async function submitRequest(itemId) {
  openModal(html`
    <div class="modal-header">
      <span class="modal-title">Submit for Review</span>
      <button class="btn btn-ghost btn-icon btn-sm" onclick="closeModal()">
        <svg width="14" height="14" viewBox="0 0 14 14" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M1 1l12 12M13 1L1 13"/></svg>
      </button>
    </div>
    <div class="modal-body">
      <p style="color:var(--text2);">This will submit your form request to the admin team for review. You won't be able to edit it until they respond.</p>
    </div>
    <div class="modal-footer">
      <button class="btn btn-ghost" onclick="closeModal()">Cancel</button>
      <button class="btn btn-primary" data-id="${itemId}" onclick="closeModal();doSubmitRequest(this.dataset.id)">Submit for Review</button>
    </div>
  `);
}

async function doSubmitRequest(itemId) {
  try {
    await updateListItem(CONFIG.FORMS_LIST, itemId, { [CONFIG.COL_STATUS]: "Submitted" });
    showToast("success", "Form submitted for review!");
    renderAdminReview(document.getElementById("main-content"));
  } catch (e) {
    showToast("error", "Submit failed: " + e.message);
  }
}

async function recallRequest(itemId) {
  openModal(html`
    <div class="modal-header">
      <span class="modal-title">Recall Submission</span>
      <button class="btn btn-ghost btn-icon btn-sm" onclick="closeModal()">
        <svg width="14" height="14" viewBox="0 0 14 14" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M1 1l12 12M13 1L1 13"/></svg>
      </button>
    </div>
    <div class="modal-body">
      <p style="color:var(--text2);">This will recall your form from the admin review queue and return it to Created. You can then edit and re-submit when ready.</p>
    </div>
    <div class="modal-footer">
      <button class="btn btn-ghost" onclick="closeModal()">Cancel</button>
      <button class="btn btn-secondary" data-id="${itemId}" onclick="closeModal();doRecallRequest(this.dataset.id)">Recall</button>
    </div>
  `);
}

async function doRecallRequest(itemId) {
  try {
    await updateListItem(CONFIG.FORMS_LIST, itemId, { [CONFIG.COL_STATUS]: "Created" });
    showToast("success", "Form recalled — you can now edit and re-submit.");
    renderAdminReview(document.getElementById("main-content"));
  } catch (e) {
    showToast("error", "Recall failed: " + e.message);
  }
}

async function submitBuilderForReview() {
  if (!validateCurrentStep()) return;
  const form = AppState.builderForm;
  if (!form.title) { showToast("error", "Please complete the form identity step"); return; }

  openModal(`
    <div class="modal-header"><span class="modal-title">Save Form</span></div>
    <div class="modal-body">
      <p style="color:var(--text2);">This will save your form. You can then submit it for admin review from Form Requests.</p>
      <p style="margin-top:8px;font-size:13px;color:var(--text3);">Form: <strong style="color:var(--text)">${escHtml(form.title)}</strong></p>
    </div>
    <div class="modal-footer">
      <button class="btn btn-ghost" onclick="closeModal()">Cancel</button>
      <button class="btn btn-primary" onclick="doSubmitForReview()">Save</button>
    </div>
  `);
}

async function doSubmitForReview() {
  closeModal();
  showProgress("Saving Form", "Saving form definition…");
  try {
    await saveBuilderDraft();
    updateProgress("Saving…");
    await updateListItem(CONFIG.FORMS_LIST, AppState.builderItemId, { [CONFIG.COL_STATUS]: "Created" });
    hideProgress();
    showToast("success", "Form saved — submit it for review from Form Requests when ready.");
    navigateTo("admin-review");
  } catch (e) {
    hideProgress();
    showToast("error", "Submit failed: " + e.message);
  }
}