// Form Studio — My Submissions

// =============================================================
// MY SUBMISSIONS
// =============================================================
async function renderMySubmissions(container) {
  container.innerHTML = `
    <div class="flex items-center justify-between mb-4">
      <h1 style="font-size:22px;font-weight:600;letter-spacing:-0.02em;">My Submissions</h1>
    </div>
    <div id="submissions-container">
      <div style="padding:40px;text-align:center;"><span class="spinner"></span></div>
    </div>
  `;

  try {
    const allForms = await getListItems(CONFIG.FORMS_LIST);

    // Only query data lists for forms that are actually deployed (Preview or Live)
    // Draft/Submitted items have no data list yet — querying them would always fail
    const deployedForms = allForms.filter(lf => {
      const s = lf.fields?.[CONFIG.COL_STATUS] || "";
      return s === "Live" || s === "Preview";
    });

    const results = await Promise.all(
      deployedForms.map(async lf => {
        const listName = lf.fields?.[CONFIG.COL_LISTNAME] || lf.fields?.ListName;
        if (!listName) return [];
        try {
          const items = await getListItems(listName);
          return items.map(i => ({ ...i, _formTitle: lf.fields?.Title }));
        } catch (_) { return []; }
      })
    );

    const allSubmissions = results.flat();
    const cont = document.getElementById("submissions-container");

    if (!allSubmissions.length) {
      cont.innerHTML = `<div class="empty-state"><h3>No submissions yet</h3><p>Forms you submit will appear here.</p></div>`;
      return;
    }

    cont.innerHTML = `
      <div class="card">
        <div class="table-wrap">
          <table>
            <thead><tr><th>Form</th><th>Submitted</th></tr></thead>
            <tbody>
              ${allSubmissions.map(i => `<tr>
                <td><strong>${escHtml(i._formTitle || "—")}</strong></td>
                <td style="color:var(--text2);font-size:12.5px;">${formatDate(i.fields?.Created)}</td>
              </tr>`).join("")}
            </tbody>
          </table>
        </div>
      </div>
    `;
  } catch (e) {
    document.getElementById("submissions-container").innerHTML =
      `<div class="empty-state"><p style="color:var(--red)">Error: ${escHtml(e.message)}</p></div>`;
  }
}
