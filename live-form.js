// Form Studio — Live Form Renderer

// =============================================================
// LIVE FORMS BROWSER
// =============================================================
async function renderLiveForms(container) {
  container.innerHTML = `
    <div class="flex items-center justify-between mb-4">
      <div>
        <h1 style="font-size:22px;font-weight:600;letter-spacing:-0.02em;">Available Forms</h1>
        <p style="color:var(--text2);font-size:13.5px;margin-top:2px;">Browse and submit live forms</p>
      </div>
    </div>
    <div id="live-forms-grid" style="display:grid;grid-template-columns:repeat(auto-fill,minmax(280px,1fr));gap:16px;">
      <div style="padding:40px;text-align:center;"><span class="spinner"></span></div>
    </div>
  `;

  try {
    const items = await getListItems(CONFIG.FORMS_LIST);
    const grid = document.getElementById("live-forms-grid");
    AppState.liveForms = items;

    // Determine whether the current user is a student based on email domain.
    // Students: @student.le.ac.uk or @student.leicester.ac.uk
    // Staff:    @leicester.ac.uk without the "student" subdomain
    const currentEmail = (AppState.currentUser?.email || "").toLowerCase();
    const isStudent = currentEmail.includes("@student.le.ac.uk") ||
                      currentEmail.includes("@student.leicester.ac.uk");

    // Load definitions for all Live non-retro items in parallel so we can check the access field.
    // Retro forms have no JSON definition and are only shown on the home page, not here.
    // Preview items don't need this — they are already restricted to admin/author only.
    const liveItems = items.filter(i => i.fields?.Status === "Live" && !i.fields?.[CONFIG.COL_RETRO]);
    const defResults = await Promise.allSettled(
      liveItems.map(i => getFormDefinition(CONFIG.FORMS_LIST, i.id))
    );
    // Build a Map of itemId → definition for O(1) lookup in the filter below.
    const defMap = new Map(
      liveItems.map((i, idx) => [
        i.id,
        defResults[idx].status === "fulfilled" ? defResults[idx].value : null,
      ])
    );

    // Returns true if the current user is allowed to see this form card.
    function canSeeForm(item) {
      const s = item.fields?.Status;
      const isOwn = (item.createdBy?.user?.email || "").toLowerCase() === currentEmail;

      // Retro forms are only shown on the home page — exclude them from this grid.
      if (item.fields?.[CONFIG.COL_RETRO]) return false;

      // Preview: only admins and the form's own author see it.
      if (s === "Preview") return AppState.isAdmin || isOwn;
      if (s !== "Live") return false;

      // Admins and the form's own author always see Live forms.
      if (AppState.isAdmin || isOwn) return true;

      const def = defMap.get(item.id);
      const access = def?.access || "StaffStudents"; // default to broadest if definition missing

      if (access === "StaffStudents") return true;        // everyone can see it

      if (access === "StaffOnly") return !isStudent;      // hidden from students

      if (access === "Specific") {
        // Show only if the current user is explicitly listed in specificPeople.
        const people = def?.specificPeople || [];
        return people.some(p => (p.email || "").toLowerCase() === currentEmail);
      }

      return true; // unknown access value — fail open
    }

    const visible = items.filter(canSeeForm);


    if (!visible.length) {
      grid.innerHTML = `<div class="empty-state" style="grid-column:1/-1;"><h3>No forms available</h3><p>Check back later.</p></div>`;
      return;
    }

    grid.innerHTML = visible.map(item => {
      const f = item.fields || {};
      const isPreview = f.Status === "Preview";
      return html`
        <div class="card" style="cursor:pointer;transition:var(--transition);" data-id="${item.id}" onclick="openLiveForm(this.dataset.id)"
          onmouseover="this.style.borderColor='var(--border2)'" onmouseout="this.style.borderColor='var(--border)'">
          <div class="card-body" style="padding:20px;">
            <div class="flex items-center gap-2 mb-3">
              ${safeHtml(isPreview ? `<span class="badge badge-amber">Preview</span>` : `<span class="badge badge-green">Live</span>`)}
            </div>
            <div style="font-size:16px;font-weight:600;margin-bottom:4px;">${f.Title||"Untitled"}</div>
            <div style="font-size:12px;color:var(--text3);font-family:var(--mono);">${f.ListName||""}</div>
          </div>
        </div>
      `;
    }).join("");
  } catch (e) {
    document.getElementById("live-forms-grid").innerHTML = `<div class="empty-state"><p style="color:var(--red)">Error: ${escHtml(e.message)}</p></div>`;
  }
}
// =============================================================
// LIVE FORM RENDERER
// =============================================================
async function openLiveForm(itemId, editItemId) {
  const main = document.getElementById("main-content");
  main.innerHTML = `<div style="padding:60px;text-align:center;"><span class="spinner" style="width:32px;height:32px;border-width:3px;"></span><p style="margin-top:16px;color:var(--text2)">Loading form…</p></div>`;

  try {
    const siteId = await getSiteId();
    const listId = await getListId(CONFIG.FORMS_LIST);
    const item = await graphGet(`/sites/${siteId}/lists/${listId}/items/${itemId}?expand=fields,createdBy`);
    if (!item || !item.id) throw new Error("Form item not found");
    const f = { ...item.fields || {}, createdBy: item.createdBy };

    const def = await getFormDefinition(CONFIG.FORMS_LIST, itemId);
    if (!def) throw new Error("Form definition not found. Ensure the FormDefinition column exists on the Forms list.");

    // If editing an existing submission, load its values for pre-population
    let prefillValues = undefined;
    if (editItemId) {
      const listName = f[CONFIG.COL_LISTNAME] || def.listName;
      try {
        const siteId2 = await getSiteId();
        const dataListId = await getListId(listName);
        const subItem = await graphGet(`/sites/${siteId2}/lists/${dataListId}/items/${editItemId}?expand=fields`);
        prefillValues = subItem.fields || {};
        if (CONFIG.DEBUG_LOGGING) console.log("[Edit] prefillValues from SP:", JSON.stringify(prefillValues));
      } catch (e) {
        console.warn("Could not load submission for editing:", e.message);
      }
    }

    renderLiveFormUI(main, f, def, itemId, prefillValues, editItemId || null);
  } catch (e) {
    main.innerHTML = `<div class="empty-state"><p style="color:var(--red)">Failed to load form: ${escHtml(e.message)}</p><button class="btn btn-secondary mt-4" onclick="navigateTo('live-forms')">← Back</button></div>`;
  }
}

// =============================================================
// CONDITION EVALUATOR
// Supports operators: eq, neq, gt, gte, lt, lte, contains
// Works for Text, Choice, Boolean, Number, Currency, DateTime
// =============================================================
function evaluateCondition(cond, formValues) {
  const actual   = formValues[cond.whenFieldId];
  const expected = cond.equalsValue;
  const op       = cond.operator || "eq"; // default to "eq" for legacy conditions

  // Boolean: compare loosely against "true"/"false" strings
  if (typeof actual === "boolean") {
    const expBool = expected === "true" || expected === true;
    if (op === "eq")  return actual === expBool;
    if (op === "neq") return actual !== expBool;
    return actual === expBool;
  }

  // Numeric comparison (Number, Currency)
  const actualNum   = parseFloat(actual);
  const expectedNum = parseFloat(expected);
  if (!isNaN(actualNum) && !isNaN(expectedNum) && (op === "gt" || op === "gte" || op === "lt" || op === "lte")) {
    if (op === "gt")  return actualNum >  expectedNum;
    if (op === "gte") return actualNum >= expectedNum;
    if (op === "lt")  return actualNum <  expectedNum;
    if (op === "lte") return actualNum <= expectedNum;
  }

  // Date comparison — compare ISO strings lexicographically (works for YYYY-MM-DD)
  if (op === "gt" || op === "lt") {
    const aStr = String(actual  ?? "").slice(0, 10);
    const eStr = String(expected ?? "").slice(0, 10);
    if (aStr && eStr) {
      if (op === "gt") return aStr > eStr;
      if (op === "lt") return aStr < eStr;
    }
  }

  // String comparison
  const actualStr   = String(actual   ?? "").toLowerCase();
  const expectedStr = String(expected ?? "").toLowerCase();
  if (op === "eq")       return actualStr === expectedStr;
  if (op === "neq")      return actualStr !== expectedStr;
  if (op === "contains") return actualStr.includes(expectedStr);

  return actualStr === expectedStr;
}

function renderLiveFormUI(container, formMeta, def, liveFormItemId, prefillValues, editItemId, existingFormValues) {
  const isMultiStep = (def.layout || "single") === "multistep";
  const allSections = def.sections || [];

  // Determine if current user is a manager or admin — they can see manager-only sections
  const currentEmail = (AppState.currentUser?.email || "").toLowerCase();
  const isAuthor = (formMeta.createdBy?.user?.email || "").toLowerCase() === currentEmail;
  const isManager = AppState.isAdmin || isAuthor ||
    (def.formManagers || []).some(m => (m.email || "").toLowerCase() === currentEmail);

  // Filter out manager-only sections for non-managers
  const sections = isManager
    ? allSections
    : allSections.filter(s => !s.managerOnly);

  // If existingFormValues passed (step navigation), use them directly.
  // Otherwise build from prefillValues (fresh SP load).
  let formValues;
  if (existingFormValues) {
    formValues = existingFormValues;
  } else {
    formValues = {};
    if (prefillValues) {
      const allFieldsList = allSections.flatMap(s => s.fields || []);
      for (const field of allFieldsList) {
        if (field.type === "InfoText" || field.type === "FileUpload") continue;
        const raw = prefillValues[field.internalName || field.label];
        if (raw === undefined || raw === null) continue;
        if (field.type === "User") {
          const normalise = r => ({
            id: String(r.LookupId),
            displayName: r.LookupValue || r.Email || "",
            email: r.Email || "",
          });
          if (Array.isArray(raw)) {
            formValues[field.id] = raw.filter(r => r.LookupId).map(normalise);
          } else if (raw && typeof raw === "object" && raw.LookupId !== undefined) {
            formValues[field.id] = [normalise(raw)];
          }
        } else {
          formValues[field.id] = raw;
        }
      }
    }
  }

  const isEditMode = !!editItemId;

  // currentStep: use existing state's step when navigating, otherwise start at 0
  const currentStep = (existingFormValues && window._liveFormState)
    ? window._liveFormState.currentStep
    : 0;

  // Group sections by step for multi-step
  const steps = [];
  if (isMultiStep) {
    // If all sections share the same stepIndex (e.g. all 0 from older forms),
    // auto-assign each section its own incremental step so multi-step actually works
    const allSame = sections.length > 0 &&
      sections.every(s => (s.stepIndex || 0) === (sections[0].stepIndex || 0));
    if (allSame && sections.length > 1) {
      sections.forEach((s, i) => { s.stepIndex = i; });
    }
    const maxStep = sections.length > 0
      ? Math.max(...sections.map(s => s.stepIndex || 0)) + 1
      : 1;
    for (let i = 0; i < maxStep; i++) {
      steps.push(sections.filter(s => (s.stepIndex || 0) === i));
    }
  } else {
    steps.push(sections);
  }

  function renderCurrentStep() {
    const stepSections = steps[currentStep] || [];

    return stepSections.map(sec => {
      const visibleFields = sec.fields.filter(field => {
        // System fields that are purely read-only (Completed, CompletedDate, CompletedBy)
        // are never rendered as form inputs — they are displayed via the completion panel.
        // DeptEmail (systemRole "DeptEmail") IS rendered as a normal editable text input.
        if (field.system && field.systemRole !== "DeptEmail") return false;

        // Check conditions — find ALL conditions that target this field (show when ALL pass)
        const conds = def.conditions?.filter(c => c.showFieldId === field.id) || [];
        if (!conds.length) return true;
        return conds.every(cond => evaluateCondition(cond, formValues));
      });

      // For managerOnly sections, check completion state from prefillValues (SP item fields).
      // prefillValues uses SP internal column names as keys.
      let completionPanel = "";
      if (sec.managerOnly && isManager) {
        const key = sectionKey(sec);
        const completedColName  = `${key}_Completed`;
        const dateColName       = `${key}_CompletedDate`;
        const byColName         = `${key}_CompletedBy`;

        // Read completion state from the raw SP prefill values — not formValues —
        // because these are system-managed fields never entered via the form UI.
        const isCompleted  = prefillValues?.[completedColName] === true;
        const completedBy  = prefillValues?.[byColName] || "";
        const completedRaw = prefillValues?.[dateColName] || "";
        const completedDate = completedRaw
          ? new Date(completedRaw).toLocaleString("en-GB", {
              day: "2-digit", month: "short", year: "numeric",
              hour: "2-digit", minute: "2-digit"
            })
          : "";

        // Read the stored comment from formValues — it's stored in a Comment field
        // we write alongside the completion fields.
        const commentColName = `${key}_CompletedComment`;
        const completedComment = prefillValues?.[commentColName] || "";

        if (isCompleted) {
          // Show the completion info panel — no Complete button
          completionPanel = `
            <div style="margin:16px 0 8px;padding:14px 16px;background:rgba(34,197,94,0.08);border:1px solid rgba(34,197,94,0.25);border-radius:8px;display:flex;gap:12px;align-items:flex-start;">
              <svg width="18" height="18" viewBox="0 0 16 16" fill="none" stroke="rgba(34,197,94,0.9)" stroke-width="1.5" style="flex-shrink:0;margin-top:1px;"><path d="M13 4L6 11 3 8" stroke-linecap="round" stroke-linejoin="round"/></svg>
              <div style="font-size:13px;color:var(--text);">
                <div style="font-weight:600;margin-bottom:3px;">Completed by ${escHtml(completedBy)}${completedDate ? ` on ${escHtml(completedDate)}` : ""}</div>
                ${completedComment ? `<div style="color:var(--text2);font-style:italic;">"${escHtml(completedComment)}"</div>` : ""}
              </div>
            </div>
          `;
        } else if (editItemId) {
          // Only show the Complete button when viewing an existing submission (editItemId exists).
          // On a fresh unsaved submission there is no SP item ID to update yet.
          const deptEmailField = sec.fields.find(f => f.system && f.systemRole === "DeptEmail");
          const deptEmailColName = deptEmailField?.internalName || `${key}_DeptEmail`;
          completionPanel = `
            <div style="margin:16px 0 8px;padding:12px 16px;background:var(--bg3);border:1px solid var(--border);border-radius:8px;display:flex;justify-content:flex-end;">
              <button class="btn btn-primary btn-sm"
                data-secid="${escAttr(sec.id)}"
                data-key="${escAttr(key)}"
                data-listname="${escAttr(def.listName || formMeta[CONFIG.COL_LISTNAME] || "")}"
                data-itemid="${escAttr(editItemId)}"
                data-deptemailcol="${escAttr(deptEmailColName)}"
                onclick="openSectionCompleteModal(this.dataset.secid, this.dataset.key, this.dataset.listname, this.dataset.itemid, this.dataset.deptemailcol)">
                ✓ Mark Section Complete
              </button>
            </div>
          `;
        }
      }

      return `
        <div class="preview-section">
          ${sec.title ? `<div class="preview-section-title">${escHtml(sec.title)}</div>` : ""}
          ${visibleFields.map(field => renderLiveField(field, formValues, def)).join("")}
          ${completionPanel}
        </div>
      `;
    }).join("");
  }

  function update() {
    document.getElementById("live-form-body").innerHTML = renderCurrentStep();
    attachLiveFieldListeners(def, formValues, update);
    updateStepIndicator();
  }

  container.innerHTML = `
    <div class="flex items-center gap-3 mb-4">
      <button class="btn btn-ghost btn-sm" onclick="navigateTo('${isEditMode ? "my-forms" : "home"}')">
        <svg width="14" height="14" viewBox="0 0 14 14" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M9 2L2 7l7 5"/></svg>
        ${isEditMode ? "My Forms" : "Back"}
      </button>
      <h1 style="font-size:22px;font-weight:600;letter-spacing:-0.02em;">${escHtml(formMeta.Title||def.title||"Form")}</h1>
      ${(formMeta[CONFIG.COL_STATUS]||formMeta.Status) === "Preview" ? `<span class="badge badge-amber">Preview</span>` : ""}

    </div>
    ${isMultiStep ? `
      <div class="wizard-steps mb-4" id="live-step-indicator">
        ${steps.map((_, i) => `
          <div class="wizard-step ${i < currentStep ? "done" : i === currentStep ? "active" : ""}">
            <div class="wizard-step-num">${i+1}</div>
          </div>
          ${i < steps.length-1 ? '<div class="wizard-connector"></div>' : ""}
        `).join("")}
      </div>
    ` : ""}
    <div class="card">
      <div id="live-form-body" class="card-body" style="display:flex;flex-direction:column;gap:0;"></div>
      <div style="padding:16px 24px;border-top:1px solid var(--border);display:flex;justify-content:space-between;align-items:center;">
        <div>
          ${isMultiStep ? `<button id="live-prev-btn" class="btn btn-secondary" onclick="liveFormPrev()" ${currentStep===0?"disabled":""}>← Previous</button>` : ""}
        </div>
        <div class="flex gap-2">
          ${isMultiStep && currentStep < steps.length-1
            ? `<button class="btn btn-primary" onclick="liveFormNext(${steps.length})">Next →</button>`
            : `<button class="btn btn-primary" id="live-submit-btn"
                data-formid="${liveFormItemId}"
                data-listname="${def.listName||formMeta.ListName||""}"
                data-editid="${editItemId||""}"
                onclick="submitLiveFormFromEl(this)">${isEditMode ? "Save Changes" : "Submit"}</button>`
          }
        </div>
      </div>
    </div>
  `;

  // Store state on window for callbacks — preserve any files already selected across step re-renders
  const existingFiles = window._liveFormState?.files || {};
  window._liveFormState = { def, formMeta, formValues, currentStep, steps, liveFormItemId, editItemId: editItemId||null, update, files: existingFiles };

  function updateStepIndicator() {
    const ind = document.getElementById("live-step-indicator");
    if (!ind) return;
    const cs = window._liveFormState.currentStep;
    ind.innerHTML = steps.map((_, i) => `
      <div class="wizard-step ${i < cs ? "done" : i === cs ? "active" : ""}">
        <div class="wizard-step-num">${i+1}</div>
      </div>
      ${i < steps.length-1 ? '<div class="wizard-connector"></div>' : ""}
    `).join("");
  }

  update();
}

function renderLiveField(field, formValues, def) {
  const val = formValues[field.id] || "";

  switch (field.type) {
    case "InfoText": {
      const safeStyle = ["info","warning","success","neutral"].includes(field.infoStyle) ? field.infoStyle : "info";
      const cleanContent = typeof DOMPurify !== "undefined"
        ? DOMPurify.sanitize(field.infoContent || "")
        : escHtml(field.infoContent || "");
      return html`<div class="infotext-block infotext-${safeStyle}">${safeHtml(cleanContent)}</div>`;
    }
    case "Note":
    case "RichText":
      return `<div class="form-group preview-field">
        <label class="preview-label">${escHtml(field.label)}${field.required?'<span class="req">*</span>':""}</label>
        ${field.description ? `<span class="input-hint">${escHtml(field.description)}</span>` : ""}
        <textarea class="textarea" data-field-id="${field.id}" placeholder="${escAttr(field.description||"")}">${escHtml(val)}</textarea>
      </div>`;
    case "Boolean": {
      // Initialise to false so required validation passes (false is a valid answer)
      if (formValues[field.id] === undefined || formValues[field.id] === "") {
        formValues[field.id] = false;
      }
      const boolVal = formValues[field.id] === true || formValues[field.id] === "true";
      return `<div class="form-group preview-field">
        <div class="flex items-center gap-2">
          <input type="checkbox" class="toggle" data-field-id="${field.id}" ${boolVal?"checked":""}>
          <label style="font-size:13.5px;color:var(--label-color);cursor:pointer;font-weight:600;">
            ${escHtml(field.label)}${field.required?'<span class="req" style="color:var(--red);margin-left:3px;">*</span>':""}
          </label>
        </div>
        ${field.description ? `<span class="input-hint" style="margin-left:44px;">${escHtml(field.description)}</span>` : ""}
      </div>`;
    }
    case "Choice":
      const childMapping = def.dependentDropdowns?.find(dd => dd.childFieldId === field.id);
      let choices = field.choices || [];
      if (childMapping) {
        const parentVal = formValues[childMapping.parentFieldId];
        if (parentVal && childMapping.mapping?.[parentVal]) choices = childMapping.mapping[parentVal];
      }
      return `<div class="form-group preview-field">
        <label class="preview-label">${escHtml(field.label)}${field.required?'<span class="req">*</span>':""}</label>
        <select class="select" data-field-id="${field.id}">
          <option value="">Select an option</option>
          ${choices.map(c => `<option value="${escAttr(c)}" ${val===c?"selected":""}>${escHtml(c)}</option>`).join("")}
        </select>
      </div>`;
    case "MultiChoice":
      return `<div class="form-group preview-field">
        <label class="preview-label">${escHtml(field.label)}${field.required?'<span class="req">*</span>':""}</label>
        <div style="display:flex;flex-direction:column;gap:8px;">
          ${(field.choices||[]).map(c => `
            <label style="display:flex;align-items:center;gap:8px;cursor:pointer;font-size:13.5px;color:var(--text);">
              <input type="checkbox" data-field-id="${field.id}" value="${escAttr(c)}" ${(val||[]).includes(c)?"checked":""}>
              ${escHtml(c)}
            </label>
          `).join("")}
        </div>
      </div>`;
    case "DateTime":
      return `<div class="form-group preview-field">
        <label class="preview-label">${escHtml(field.label)}${field.required?'<span class="req">*</span>':""}</label>
        <input type="date" class="input" data-field-id="${field.id}" value="${escAttr(val ? val.split('T')[0] : '')}"
          style="cursor:pointer;" onclick="this.showPicker&&this.showPicker()">
      </div>`;
    case "User": {
      // formValues[field.id] is an array of {id, displayName, email}
      const people = Array.isArray(val) ? val : (val && typeof val === "object" && val.id ? [val] : []);
      return `<div class="form-group preview-field" id="person-field-group-${field.id}">
        <label class="preview-label">${escHtml(field.label)}${field.required?'<span class="req">*</span>':""}</label>
        ${field.description ? `<span class="input-hint">${escHtml(field.description)}</span>` : ""}
        <div style="position:relative;">
          <input class="input" id="person-search-${field.id}"
            placeholder="Search by name or email…"
            autocomplete="off"
            oninput="debouncedPersonFieldSearch('${field.id}', this.value)">
          <div id="person-results-${field.id}" style="position:absolute;top:100%;left:0;right:0;z-index:200;margin-top:2px;"></div>
        </div>
        <div id="person-selected-${field.id}" style="margin-top:6px;display:flex;flex-wrap:wrap;gap:6px;">
          ${people.map(p => {
            const initials = (p.displayName || "").split(" ").map(n => n[0] || "").join("").slice(0, 2).toUpperCase();
            return `<span class="person-chip">
              <div class="avatar" style="width:20px;height:20px;font-size:9px;">${escHtml(initials)}</div>
              ${escHtml(p.displayName || "")}
              <button type="button" data-fieldid="${escAttr(field.id)}" data-personid="${escAttr(p.id)}" onclick="removePersonFromField(this.dataset.fieldid, this.dataset.personid)" style="background:none;border:none;cursor:pointer;padding:0;line-height:1;">
                <svg width="10" height="10" viewBox="0 0 10 10" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M1 1l8 8M9 1L1 9"/></svg>
              </button>
            </span>`;
          }).join("")}
        </div>
      </div>`;
    }
    case "Currency": {
      // Store raw numeric string in formValues; display with symbol and thousand-separators
      const numVal = val !== "" && val !== undefined && val !== null ? val : "";
      const rawVal = numVal !== "" ? String(parseFloat(numVal)) : "";
      const displayVal = numVal !== "" ? parseFloat(numVal).toLocaleString("en-GB", { minimumFractionDigits: 2, maximumFractionDigits: 2 }) : "";
      return `<div class="form-group preview-field">
        <label class="preview-label">${escHtml(field.label)}${field.required ? '<span class="req">*</span>' : ""}</label>
        ${field.description ? `<span class="input-hint">${escHtml(field.description)}</span>` : ""}
        <div style="position:relative;display:flex;align-items:center;max-width:200px;">
          <span style="position:absolute;left:12px;font-size:14px;color:var(--text2);pointer-events:none;user-select:none;z-index:1;">£</span>
          <input
            type="text"
            inputmode="decimal"
            class="input"
            data-field-id="${field.id}"
            data-field-type="Currency"
            data-raw="${escAttr(rawVal)}"
            placeholder="0.00"
            value="${escAttr(displayVal)}"
            style="padding-left:28px;font-variant-numeric:tabular-nums;"
            onfocus="this.value=this.dataset.raw||'';"
            onblur="updateCurrencyField(this,'${field.id}')"
            oninput="updateCurrencyField(this,'${field.id}',true)"
          >
        </div>
      </div>`;
    }
    case "URL":
      return `<div class="form-group preview-field">
        <label class="preview-label">${escHtml(field.label)}${field.required?'<span class="req">*</span>':""}</label>
        ${field.description ? `<span class="input-hint">${escHtml(field.description)}</span>` : ""}
        <input type="url" class="input" data-field-id="${field.id}" placeholder="https://…" value="${escAttr(val)}">
      </div>`;
    case "Number":
      return `<div class="form-group preview-field">
        <label class="preview-label">${escHtml(field.label)}${field.required?'<span class="req">*</span>':""}</label>
        ${field.description ? `<span class="input-hint">${escHtml(field.description)}</span>` : ""}
        <input type="number" step="any" class="input" data-field-id="${field.id}" placeholder="0" value="${escAttr(val)}" style="max-width:200px;">
      </div>`;
    case "FileUpload": {
      const acceptAttr = field.accept ? escAttr(field.accept) : "";

      // Detect if all accepted types are images — if so, enable camera capture on mobile
      const isImageOnly = (() => {
        if (!field.accept) return false;
        const types = field.accept.split(",").map(t => t.trim().toLowerCase());
        return types.length > 0 && types.every(t =>
          t === "image/*" ||
          t.startsWith("image/") ||
          [".jpg",".jpeg",".png",".gif",".webp",".heic",".heif",".bmp",".svg"].includes(t)
        );
      })();

      const captureAttr = isImageOnly ? 'capture="environment"' : "";
      const fileName = val || "";
      const icon = isImageOnly
        ? `<svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="var(--text3)" stroke-width="1.5" style="margin-bottom:8px;"><path d="M23 19a2 2 0 01-2 2H3a2 2 0 01-2-2V8a2 2 0 012-2h4l2-3h6l2 3h4a2 2 0 012 2z"/><circle cx="12" cy="13" r="4"/></svg>`
        : `<svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="var(--text3)" stroke-width="1.5" style="margin-bottom:8px;"><path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4M17 8l-5-5-5 5M12 3v12"/></svg>`;
      const label = isImageOnly ? "Take photo or choose image" : "Click to select a file";

      return html`
        <div class="form-group preview-field" id="field-group-${field.id}">
          <label class="preview-label">${field.label}${field.required ? '<span class="req">*</span>' : ""}</label>
          ${safeHtml(field.description ? html`<span class="input-hint">${field.description}</span>` : "")}
          <div style="border:2px dashed var(--border);border-radius:var(--radius-sm);padding:20px;text-align:center;background:var(--bg3);">
            ${safeHtml(icon)}
            <div style="font-size:13px;color:var(--text2);margin-bottom:8px;">${label}</div>
            ${safeHtml(acceptAttr ? html`<div style="font-size:11.5px;color:var(--text3);margin-bottom:8px;">Accepted: ${field.accept}</div>` : "")}
            <input type="file" id="file-input-${field.id}" data-field-id="${field.id}" style="display:none;"
              ${acceptAttr ? `accept="${acceptAttr}"` : ""}
              ${captureAttr}
              multiple
              onchange="handleFileSelect(this)">
            <button class="btn btn-secondary btn-sm" data-fieldid="${field.id}" onclick="triggerFileInput(this.dataset.fieldid)">
              ${isImageOnly ? "Take Photo / Choose Images" : "Choose Files"}
            </button>
            <div id="file-selected-${field.id}" style="margin-top:8px;font-size:12.5px;color:var(--accent);">${fileName ? `Selected: ${escHtml(fileName)}` : ""}</div>
          </div>
        </div>
      `;
    }
    default:
      return `<div class="form-group preview-field">
        <label class="preview-label">${escHtml(field.label)}${field.required?'<span class="req">*</span>':""}</label>
        ${field.description ? `<span class="input-hint">${escHtml(field.description)}</span>` : ""}
        <input class="input" data-field-id="${field.id}" placeholder="${escAttr(field.description||"")}" value="${escAttr(val)}">
      </div>`;
  }
}

// =============================================================
// PERSON FIELD — live search
// =============================================================
const _personFieldSearchTimers = {};

function debouncedPersonFieldSearch(fieldId, query) {
  clearTimeout(_personFieldSearchTimers[fieldId]);
  _personFieldSearchTimers[fieldId] = setTimeout(() => personFieldSearch(fieldId, query), 350);
}

async function personFieldSearch(fieldId, query) {
  const resultsEl = document.getElementById(`person-results-${fieldId}`);
  if (!resultsEl) return;
  if (!query || query.length < 2) { resultsEl.innerHTML = ""; return; }

  resultsEl.innerHTML = `<div style="background:var(--bg3);border:1px solid var(--border);border-radius:var(--radius-sm);padding:10px 12px;font-size:12.5px;color:var(--text3);">Searching…</div>`;
  try {
    const people = await searchPeople(query);
    if (!people.length) {
      resultsEl.innerHTML = `<div style="background:var(--bg3);border:1px solid var(--border);border-radius:var(--radius-sm);padding:10px 12px;font-size:12.5px;color:var(--text3);">No results found.</div>`;
      return;
    }
    resultsEl.innerHTML = html`
      <div style="background:var(--bg);border:1px solid var(--border);border-radius:var(--radius-sm);box-shadow:var(--shadow-md);overflow:hidden;">
        ${safeHtml(people.slice(0, 6).map(p => {
          const email = p.scoredEmailAddresses?.[0]?.address || "";
          const initials = p.displayName.split(" ").map(n => n[0] || "").join("").slice(0, 2).toUpperCase();
          return html`
            <div class="flex items-center gap-2" style="padding:8px 12px;cursor:pointer;border-bottom:1px solid var(--border);"
              data-fieldid="${fieldId}"
              data-id="${p.id}"
              data-name="${p.displayName}"
              data-email="${email}"
              onclick="selectPersonField(this)"
              onmouseover="this.style.background='var(--bg3)'" onmouseout="this.style.background=''">
              <div class="avatar" style="width:24px;height:24px;font-size:10px;">${initials}</div>
              <div style="flex:1;min-width:0;">
                <div style="font-size:13px;">${p.displayName}</div>
                <div style="font-size:11.5px;color:var(--text3);">${email}</div>
              </div>
            </div>
          `;
        }).join(""))}
      </div>
    `;
  } catch (e) {
    resultsEl.innerHTML = html`<div style="background:var(--bg3);border:1px solid var(--border);border-radius:var(--radius-sm);padding:10px 12px;font-size:12.5px;color:var(--red);">Search failed.</div>`;
  }
}

function selectPersonField(el) {
  const fieldId = el.dataset.fieldid;
  const person = { id: el.dataset.id, displayName: el.dataset.name, email: el.dataset.email };

  // Ensure formValues[fieldId] is an array and add person if not already present
  if (window._liveFormState) {
    const current = window._liveFormState.formValues[fieldId];
    const arr = Array.isArray(current) ? current : (current && current.id ? [current] : []);
    if (!arr.find(p => p.id === person.id)) {
      arr.push(person);
    }
    window._liveFormState.formValues[fieldId] = arr;
  }

  // Clear search box
  const searchEl = document.getElementById(`person-search-${fieldId}`);
  if (searchEl) searchEl.value = "";

  // Clear results
  const resultsEl = document.getElementById(`person-results-${fieldId}`);
  if (resultsEl) resultsEl.innerHTML = "";

  // Re-render chips
  refreshPersonChips(fieldId);
}

function removePersonFromField(fieldId, personId) {
  if (window._liveFormState) {
    const current = window._liveFormState.formValues[fieldId];
    const arr = Array.isArray(current) ? current : [];
    window._liveFormState.formValues[fieldId] = arr.filter(p => p.id !== personId);
  }
  refreshPersonChips(fieldId);
}

function refreshPersonChips(fieldId) {
  const selectedEl = document.getElementById(`person-selected-${fieldId}`);
  if (!selectedEl) return;
  const arr = Array.isArray(window._liveFormState?.formValues[fieldId])
    ? window._liveFormState.formValues[fieldId] : [];
  selectedEl.innerHTML = arr.map(p => {
    const initials = (p.displayName || "").split(" ").map(n => n[0] || "").join("").slice(0, 2).toUpperCase();
    return `<span class="person-chip">
      <div class="avatar" style="width:20px;height:20px;font-size:9px;">${escHtml(initials)}</div>
      ${escHtml(p.displayName || "")}
      <button type="button"
        data-fieldid="${escAttr(fieldId)}"
        data-personid="${escAttr(p.id)}"
        onclick="removePersonFromField(this.dataset.fieldid, this.dataset.personid)"
        style="background:none;border:none;cursor:pointer;padding:0;line-height:1;">
        <svg width="10" height="10" viewBox="0 0 10 10" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M1 1l8 8M9 1L1 9"/></svg>
      </button>
    </span>`;
  }).join("");
}


// =============================================================
// CURRENCY FIELD HANDLER
// Named function replaces inline onblur/oninput to make
// debugging easier and keep the template string readable.
// isInput=true  → called from oninput (typing, no reformat)
// isInput=false → called from onblur  (finished, reformat)
// =============================================================
function updateCurrencyField(el, fieldId, isInput) {
  const raw = el.value.replace(/,/g, "").replace(/[^0-9.\-]/g, "");
  const n   = parseFloat(raw);
  if (CONFIG.DEBUG_LOGGING) console.log(`[Currency:${fieldId}] input="${el.value}" raw="${raw}" parsed=${n} isInput=${!!isInput}`);
  if (!isInput) {
    // blur — reformat display and update dataset.raw
    el.dataset.raw = isNaN(n) ? "" : String(n);
    el.value = isNaN(n) ? "" : n.toLocaleString("en-GB", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  }
  if (window._liveFormState) {
    window._liveFormState.formValues[fieldId] = isNaN(n) ? "" : n;
  }
}

function attachLiveFieldListeners(def, formValues, update) {
  // Fields that trigger conditional visibility — only these need to re-render on change
  const conditionTriggers = new Set((def.conditions || []).map(c => c.whenFieldId));
  const ddTriggers = new Set((def.dependentDropdowns || []).map(dd => dd.parentFieldId));
  const needsRerender = id => conditionTriggers.has(id) || ddTriggers.has(id);

  document.querySelectorAll("[data-field-id]").forEach(el => {
    const id = el.dataset.fieldId;

    if (el.type === "file") {
      // FileUpload — value stored via handleFileSelect; only re-render if it's a condition trigger
      el.addEventListener("change", () => {
        if (needsRerender(id)) update();
      });
    } else if (el.type === "checkbox" && el.classList.contains("toggle")) {
      // Boolean toggle — store as boolean, compare loosely in conditions
      el.addEventListener("change", () => {
        formValues[id] = el.checked;
        if (needsRerender(id)) update();
      });
    } else if (el.type === "checkbox" && el.value) {
      // Multi-choice checkbox — never a condition trigger
      el.addEventListener("change", () => {
        if (!Array.isArray(formValues[id])) formValues[id] = [];
        if (el.checked) { if (!formValues[id].includes(el.value)) formValues[id].push(el.value); }
        else formValues[id] = formValues[id].filter(v => v !== el.value);
      });
    } else if (el.tagName === "SELECT") {
      // Dropdowns — re-render only if condition/dependent trigger
      el.addEventListener("change", () => {
        formValues[id] = el.value;
        if (needsRerender(id)) update();
      });
    } else if (el.tagName === "INPUT" || el.tagName === "TEXTAREA") {
      // Skip person-field search inputs (they're just UI, value stored via selectPersonField)
      if (el.id && el.id.startsWith("person-search-")) return;
      // Hidden person-field inputs — value already kept in sync by selectPersonField
      if (el.type === "hidden") return;
      // Currency inputs manage formValues via their own inline oninput/onblur handlers
      // (which store the raw numeric value). The generic listener would overwrite with
      // the formatted display string (e.g. "1,234.56"), so skip it here.
      if (el.dataset.fieldType === "Currency") return;
      // Text / number / url / date / textarea
      // Update formValues silently on every keystroke
      el.addEventListener("input", () => { formValues[id] = el.value; });
      // Re-render on blur/change only if this field is a condition trigger
      el.addEventListener("change", () => {
        formValues[id] = el.value;
        if (needsRerender(id)) update();
      });
    }
  });
}

function liveFormNext(totalSteps) {
  if (!window._liveFormState) return;
  const state = window._liveFormState;
  // Validate required fields on current step before advancing
  const allFields = state.steps[state.currentStep].flatMap(s => s.fields || []);
  for (const field of allFields) {
    if (field.required && field.type !== "Boolean") {
      const v = state.formValues[field.id];
      const isEmpty = v === undefined || v === null || v === "" ||
        (Array.isArray(v) && v.length === 0);
      if (isEmpty) {
        showToast("error", `"${field.label}" is required`);
        return;
      }
    }
  }
  // Capture any text input values that may not have fired change events
  document.querySelectorAll("[data-field-id]").forEach(el => {
    const id = el.dataset.fieldId;
    if (el.type === "hidden") return; // person hidden inputs already in formValues
    if (el.id && el.id.startsWith("person-search-")) return; // skip person search UI
    if (el.dataset.fieldType === "Currency") return; // currency inputs manage formValues via inline handlers
    if (el.type !== "checkbox" && el.tagName !== "SELECT") {
      state.formValues[id] = el.value;
    }
  });
  state.currentStep = Math.min(state.currentStep + 1, state.steps.length - 1);
  renderLiveFormUI(
    document.getElementById("main-content"),
    state.formMeta, state.def, state.liveFormItemId,
    null, state.editItemId, state.formValues
  );
}

function liveFormPrev() {
  if (!window._liveFormState) return;
  const state = window._liveFormState;
  // Capture current text values before navigating back
  document.querySelectorAll("[data-field-id]").forEach(el => {
    const id = el.dataset.fieldId;
    if (el.type === "hidden") return;
    if (el.id && el.id.startsWith("person-search-")) return;
    if (el.dataset.fieldType === "Currency") return; // currency inputs manage formValues via inline handlers
    if (el.type !== "checkbox" && el.tagName !== "SELECT") {
      state.formValues[id] = el.value;
    }
  });
  state.currentStep = Math.max(state.currentStep - 1, 0);
  renderLiveFormUI(
    document.getElementById("main-content"),
    state.formMeta, state.def, state.liveFormItemId,
    null, state.editItemId, state.formValues
  );
}

function triggerFileInput(fieldId) {
  document.getElementById(`file-input-${fieldId}`)?.click();
}

function handleFileSelect(input) {
  const fieldId = input.dataset.fieldId;
  const files = Array.from(input.files || []);
  const display = document.getElementById(`file-selected-${fieldId}`);
  if (display) {
    if (!files.length) {
      display.textContent = "";
    } else if (files.length === 1) {
      display.textContent = `Selected: ${files[0].name}`;
    } else {
      display.textContent = `Selected: ${files.length} files (${files.map(f => f.name).join(", ")})`;
    }
  }
  // Store file array in liveFormState
  if (window._liveFormState) {
    if (!window._liveFormState.files) window._liveFormState.files = {};
    window._liveFormState.files[fieldId] = files.length ? files : null;
  }
}

async function uploadFormAttachment(listName, spItemId, file, isEdit) {
  const token = await getSpToken();
  const siteUrl = CONFIG.SITE_URL.replace(/\/$/, "");
  const encodedName = encodeURIComponent(file.name);
  const baseUrl = `${siteUrl}/_api/web/lists/GetByTitle('${encodeURIComponent(listName)}')/items(${spItemId})/AttachmentFiles`;

  // In edit mode, delete any existing attachment with the same filename first
  // (SP REST returns 409 Conflict if you try to add a duplicate filename)
  if (isEdit) {
    try {
      const delRes = await fetch(`${baseUrl}/getByFileName('${encodedName}')`, {
        method: "DELETE",
        headers: { Authorization: "Bearer " + token, Accept: "application/json;odata=verbose" }
      });
      // 404 just means it didn't exist — that's fine
    } catch (_) {}
  }

  const res = await fetch(`${baseUrl}/add(FileName='${encodedName}')`, {
    method: "POST",
    headers: {
      Authorization: "Bearer " + token,
      Accept: "application/json;odata=verbose",
      "Content-Type": "application/octet-stream",
    },
    body: file,
  });
  if (!res.ok) {
    const txt = await res.text().catch(() => "");
    throw new Error(`File upload failed: ${res.status} — ${txt.slice(0, 200)}`);
  }
  return file.name;
}

function submitLiveFormFromEl(el) {
  submitLiveForm(el.dataset.formid, el.dataset.listname, el.dataset.editid || null);
}

async function submitLiveForm(liveFormItemId, listName, editItemId) {
  if (!listName) { showToast("error", "List name not configured"); return; }
  const state = window._liveFormState;
  if (!state) return;

  const isEdit = !!(editItemId || state.editItemId);
  const itemId = editItemId || state.editItemId;

  const btn = document.getElementById("live-submit-btn");
  if (btn) { btn.disabled = true; btn.innerHTML = `<span class="spinner"></span> ${isEdit ? "Saving…" : "Submitting…"}`; }
  showProgress(isEdit ? "Saving Changes" : "Submitting Form", isEdit ? "Saving your changes…" : "Submitting your response…");

  try {
    // Validate required fields — including FileUpload
    const allFields = getAllFieldsFromDef(state.def);
    const currentStepFields = state.steps[state.currentStep]?.flatMap(s => s.fields || []) || allFields;
    for (const field of currentStepFields) {
      if (field.required && field.type !== "Boolean") {
        if (field.type === "FileUpload") {
          const files = state.files?.[field.id];
          const hasFiles = files && (Array.isArray(files) ? files.length > 0 : true);
          if (!hasFiles && !isEdit) {
            showToast("error", `"${field.label}" is required`);
            if (btn) { btn.disabled = false; btn.innerHTML = isEdit ? "Save Changes" : "Submit"; }
            hideProgress();
            return;
          }
        } else {
          const v = state.formValues[field.id];
          const isEmpty = v === undefined || v === null || v === "" ||
            (Array.isArray(v) && v.length === 0) ||
            (typeof v === "object" && !Array.isArray(v) && !v.displayName && !v.email) ||
            (field.type === "RichText" && typeof v === "string" && v.replace(/<[^>]*>/g, "").trim() === "");
          if (isEmpty) {
            showToast("error", `"${field.label}" is required`);
            if (btn) { btn.disabled = false; btn.innerHTML = isEdit ? "Save Changes" : "Submit"; }
            hideProgress();
            return;
          }
        }
      }
    }

    if (CONFIG.DEBUG_LOGGING) console.log("[Submit] raw formValues:", JSON.stringify(state.formValues));

    // Build fields from formValues — sanitise types for SharePoint
    // FileUpload fields are stored as attachments — skip them here
    const submitFields = {};
    const fileUploadField = allFields.find(f => f.type === "FileUpload");
    const filesToUpload = fileUploadField ? (state.files?.[fileUploadField.id] || null) : null;
    // Normalise to array
    const fileArray = filesToUpload
      ? (Array.isArray(filesToUpload) ? filesToUpload : [filesToUpload])
      : [];

    // Pre-compute column names using the EXACT same deduplication logic as provisioning (admin.js).
    // Key rule: InfoText is skipped but FileUpload IS included in the name-reservation pass
    // (it just produces null from buildColumnDefinition, but its name still occupies the namespace).
    // internalName is authoritative when present — the fallback is only for forms provisioned
    // before internalName was written back to JSON.
    const usedColNames = new Set();
    const fieldColNames = new Map();
    for (const field of allFields) {
      if (field.type === "InfoText") continue; // InfoText has no column — matches admin.js
      if (field.internalName) {
        fieldColNames.set(field.id, field.internalName);
        usedColNames.add(field.internalName.toLowerCase());
        continue;
      }
      // Replicate sanitiseColumnName from admin.js
      const baseName = ((field.label || "Field")
        .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
        .replace(/[^a-zA-Z0-9]/g, "")
        .replace(/^[0-9]+/, "") || "Field").slice(0, 32);
      let uniqueName = baseName;
      let counter = 2;
      while (usedColNames.has(uniqueName.toLowerCase())) {
        const suffix = `_${counter++}`;
        uniqueName = baseName.slice(0, 32 - suffix.length) + suffix;
      }
      usedColNames.add(uniqueName.toLowerCase());
      if (field.type !== "FileUpload") {
        // FileUpload columns aren't submitted as fields — reserve the name but don't map it
        fieldColNames.set(field.id, uniqueName);
      }
    }

    for (const field of allFields) {
      if (field.type === "InfoText") continue;
      if (field.type === "FileUpload") continue;
      const val = state.formValues[field.id];
      if (val === undefined || val === null || val === "") continue;

      const colName = fieldColNames.get(field.id);
      if (!colName) continue;

      // Person fields are collected separately — Graph API cannot write multi-value
      // Person columns; they must be patched via SP REST after the item exists.
      if (field.type === "User") continue;

      let cleanVal = val;
      if (field.type === "Boolean") {
        cleanVal = val === true || val === "true" || val === 1;
      } else if (field.type === "Number" || field.type === "Currency") {
        // Strip any thousand-separator commas that may have slipped through before parsing
        const stripped = typeof val === "string" ? val.replace(/,/g, "") : val;
        cleanVal = parseFloat(stripped);
        if (isNaN(cleanVal)) continue;
      } else if (field.type === "DateTime") {
        try { cleanVal = new Date(val).toISOString(); } catch (_) { continue; }
      } else if (field.type === "MultiChoice") {
        const arr = Array.isArray(val) ? val : [val];
        cleanVal = { results: arr };
      } else if (typeof val === "string") {
        cleanVal = val.trim();
        if (!cleanVal) continue;
      }

      submitFields[colName] = cleanVal;
    }

    // Resolve Person field values via SP ensureUser — collect as { colName: [spId, ...] }
    const personFields = {};
    for (const field of allFields) {
      if (field.type !== "User") continue;
      const val = state.formValues[field.id];
      const people = Array.isArray(val) ? val : (val && val.email ? [val] : []);
      if (!people.length) continue;
      const colName = fieldColNames.get(field.id);
      if (!colName) continue;
      const resolvedIds = [];
      for (const person of people) {
        if (!person.email) continue;
        try {
          updateProgress(`Resolving ${person.displayName || person.email}…`);
          const userData = await spPost(`/_api/web/ensureuser`,
            { logonName: `i:0#.f|membership|${person.email}` });
          const spUserId = userData?.d?.Id;
          if (spUserId) resolvedIds.push(parseInt(spUserId, 10));
        } catch (e) {
          console.warn(`[Submit] Could not resolve user "${person.email}":`, e.message);
          showToast("error", `Could not resolve "${person.displayName || person.email}" — check they exist in SharePoint.`);
        }
      }
      if (resolvedIds.length) personFields[colName] = resolvedIds;
    }

    if (CONFIG.DEBUG_LOGGING) console.log("[Submit] listName:", listName, "fields:", JSON.stringify(submitFields), "personFields:", JSON.stringify(personFields));

    // Helper: patch person fields on an SP item via REST MERGE
    async function patchPersonFields(spItemId) {
      if (!Object.keys(personFields).length) return;
      updateProgress("Saving people…");
      const siteUrl = CONFIG.SITE_URL.replace(/\/$/, "");
      // Get the list's entity type for SP REST metadata
      const listMeta = await spGet(`/_api/web/lists/GetByTitle('${encodeURIComponent(listName)}')?$select=ListItemEntityTypeFullName`);
      const entityType = listMeta?.d?.ListItemEntityTypeFullName || "SP.Data.ListItem";
      // Build the REST body — multi-value Person: { results: [id1, id2] }, single: { results: [id] }
      const body = { "__metadata": { "type": entityType } };
      for (const [colName, ids] of Object.entries(personFields)) {
        body[`${colName}Id`] = { results: ids };
      }
      await spMerge(
        `/_api/web/lists/GetByTitle('${encodeURIComponent(listName)}')/items(${spItemId})`,
        body
      );
    }

    if (isEdit) {
      await updateListItem(listName, itemId, submitFields);
      const spItemId = await getSpItemId(listName, itemId);
      await patchPersonFields(spItemId);
      if (fileArray.length) {
        updateProgress("Uploading files…");
        for (const file of fileArray) {
          await uploadFormAttachment(listName, spItemId, file, true);
        }
      }
      hideProgress();
      showToast("success", "Changes saved successfully!");
    } else {
      const newItem = await createListItem(listName, submitFields);
      const spItemId = await getSpItemId(listName, newItem.id);
      await patchPersonFields(spItemId);
      if (fileArray.length && newItem?.id) {
        updateProgress("Uploading files…");
        for (const file of fileArray) {
          await uploadFormAttachment(listName, spItemId, file, false);
        }
      }
      hideProgress();
      showToast("success", "Form submitted successfully!");
    }
    navigateTo(isEdit ? "my-forms" : "home");
  } catch (e) {
    hideProgress();
    showToast("error", `${isEdit ? "Save" : "Submit"} failed: ` + e.message);
    if (btn) { btn.disabled = false; btn.innerHTML = isEdit ? "Save Changes" : "Submit"; }
  }
}

// =============================================================
// SECTION COMPLETE BUTTON
// Shown on managerOnly sections when viewing an existing submission.
// Clicking it opens a modal for the Form Manager to enter a comment,
// then writes the completion fields to SharePoint and fires sendMail
// to the comma-separated addresses in the section's DeptEmail column.
// =============================================================

function openSectionCompleteModal(secId, key, listName, itemId, deptEmailColName) {
  openModal(html`
    <div class="modal-header">
      <span class="modal-title">Mark Section Complete</span>
      <button class="btn btn-ghost btn-icon btn-sm" onclick="closeModal()">
        <svg width="14" height="14" viewBox="0 0 14 14" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M1 1l12 12M13 1L1 13"/></svg>
      </button>
    </div>
    <div class="modal-body">
      <p style="color:var(--text2);font-size:13.5px;margin-bottom:16px;">
        This will mark this section as complete and notify the relevant team.
        This action cannot be undone.
      </p>
      <div class="form-group">
        <label for="section-complete-comment" style="font-weight:600;">Comment</label>
        <textarea id="section-complete-comment" class="textarea" style="min-height:90px;"
          placeholder="e.g. Work completed — ready for Dept B review"></textarea>
        <span class="input-hint">This comment will be included in the notification email.</span>
      </div>
    </div>
    <div class="modal-footer">
      <button class="btn btn-ghost" onclick="closeModal()">Cancel</button>
      <button class="btn btn-primary"
        data-secid="${secId}"
        data-key="${key}"
        data-listname="${listName}"
        data-itemid="${itemId}"
        data-deptemailcol="${deptEmailColName}"
        onclick="doSectionComplete(this)">
        Confirm Complete
      </button>
    </div>
  `);
}

async function doSectionComplete(btn) {
  const secId         = btn.dataset.secid;
  const key           = btn.dataset.key;
  const listName      = btn.dataset.listname;
  const itemId        = btn.dataset.itemid;
  const deptEmailCol  = btn.dataset.deptemailcol;
  const comment       = document.getElementById("section-complete-comment")?.value?.trim() || "";

  closeModal();

  // Resolve the section from the current form definition
  const state = window._liveFormState;
  if (!state) { showToast("error", "Form state lost — please reload"); return; }

  const def      = state.def;
  const formMeta = state.formMeta;
  const section  = (def.sections || []).find(s => s.id === secId);
  if (!section) { showToast("error", "Section not found"); return; }

  showProgress("Completing section", "Saving completion status…");

  try {
    const now         = new Date().toISOString();
    const displayName = AppState.currentUser?.displayName || AppState.currentUser?.email || "Unknown";

    // Build the fields to write — 4 system columns + the comment column
    const completionFields = {
      [`${key}_Completed`]:        true,
      [`${key}_CompletedDate`]:    now,
      [`${key}_CompletedBy`]:      displayName,
      [`${key}_CompletedComment`]: comment,
    };

    await updateListItem(listName, itemId, completionFields);

    // ── Fire notification email ──────────────────────────────────
    // Read the DeptEmail value from the current SP item.
    // We use the prefillValues already loaded — or re-fetch if needed.
    updateProgress("Sending notification…");

    // Get the DeptEmail value from the saved item
    const siteId    = await getSiteId();
    const dataListId = await getListId(listName);
    const spItem    = await graphGet(
      `/sites/${siteId}/lists/${dataListId}/items/${itemId}?expand=fields`
    );
    const deptEmailRaw = spItem?.fields?.[deptEmailCol] || "";

    // Parse comma-separated emails — trim each, filter empty strings
    const recipients = deptEmailRaw
      .split(",")
      .map(e => e.trim())
      .filter(e => e.includes("@"));

    if (recipients.length) {
      // Build the deep links
      const appUrl      = (CONFIG.APP_URL || "").replace(/\/$/, "");
      const formId      = state.liveFormItemId;
      const openThisUrl = `${appUrl}?view=my-forms&formId=${encodeURIComponent(formId)}&itemId=${encodeURIComponent(itemId)}`;
      const openAllUrl  = `${appUrl}?view=my-forms&formId=${encodeURIComponent(formId)}`;

      const formTitle    = escHtml(formMeta.Title || def.title || "Form");
      const sectionTitle = escHtml(section.title || key);

      // HTML email body
      const emailBody = `
        <div style="font-family:sans-serif;font-size:14px;color:#1a1a1a;max-width:600px;">
          <h2 style="font-size:18px;font-weight:600;margin-bottom:8px;">
            ${formTitle} — ${sectionTitle} Completed
          </h2>
          <p style="color:#555;margin-bottom:16px;">
            <strong>${escHtml(displayName)}</strong> has marked the
            <strong>${sectionTitle}</strong> section as complete.
          </p>
          ${comment ? `
          <div style="background:#f4f4f2;border-left:3px solid #888;padding:12px 16px;border-radius:4px;margin-bottom:20px;">
            <p style="margin:0;color:#333;font-style:italic;">"${escHtml(comment)}"</p>
          </div>` : ""}
          <div style="margin-bottom:8px;">
            <strong>Completed by:</strong> ${escHtml(displayName)}<br>
            <strong>Date:</strong> ${new Date(now).toLocaleString("en-GB", {
              day: "2-digit", month: "short", year: "numeric",
              hour: "2-digit", minute: "2-digit"
            })}
          </div>
          <div style="margin-top:24px;display:flex;gap:12px;">
            <a href="${openThisUrl}"
              style="display:inline-block;padding:10px 20px;background:#002147;color:#fff;text-decoration:none;border-radius:5px;font-weight:600;">
              Open This Form
            </a>
            <a href="${openAllUrl}"
              style="display:inline-block;padding:10px 20px;background:#f4f4f2;color:#002147;text-decoration:none;border-radius:5px;font-weight:600;border:1px solid #ddd;">
              Open All Forms
            </a>
          </div>
          <p style="margin-top:24px;font-size:12px;color:#aaa;">
            Sent by Form Studio on behalf of ${escHtml(displayName)}
          </p>
        </div>
      `;

      // Send via Graph /me/sendMail — fires as the logged-in user
      await graphPost(`/me/sendMail`, {
        message: {
          subject: `[Form Studio] ${formMeta.Title || def.title || "Form"} — ${section.title || key} Completed`,
          body: {
            contentType: "HTML",
            content: emailBody,
          },
          toRecipients: recipients.map(email => ({
            emailAddress: { address: email }
          })),
        },
        saveToSentItems: false,
      });
    }

    hideProgress();
    showToast("success", "Section marked complete" + (recipients.length ? " — notification sent" : ""));

    // Reload the form to show the completion panel instead of the button
    openLiveForm(state.liveFormItemId, itemId);

  } catch (e) {
    hideProgress();
    showToast("error", "Could not complete section: " + e.message);
  }
}