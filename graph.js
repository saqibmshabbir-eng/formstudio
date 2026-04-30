// Form Studio — Graph API Helpers

// =============================================================
// GRAPH API HELPERS
// =============================================================
async function graphGet(path) {
  const token = await getToken();
  const res = await fetch(CONFIG.GRAPH_BASE + path, {
    headers: {
      Authorization: "Bearer " + token,
      "Content-Type": "application/json",
      "Prefer": "HonorNonIndexedQueriesWarningMayFailRandomly",
    }
  });
  if (!res.ok) {
    const err = await res.json().catch(() => ({}));
    throw new Error(err?.error?.message || `Graph error ${res.status} on ${path}`);
  }
  return res.json();
}

async function graphPost(path, body) {
  const token = await getWriteToken();
  const res = await fetch(CONFIG.GRAPH_BASE + path, {
    method: "POST",
    headers: {
      Authorization: "Bearer " + token,
      "Content-Type": "application/json",
      "Prefer": "HonorNonIndexedQueriesWarningMayFailRandomly",
    },
    body: JSON.stringify(body)
  });
  if (!res.ok) {
    const txt = await res.text().catch(() => "");
    console.error("[graphPost error]", res.status, path, txt);
    let err = {};
    try { err = JSON.parse(txt); } catch (_) {}
    const detail = err?.error?.innererror?.message || err?.error?.message || `Graph error ${res.status} on ${path}`;
    throw new Error(detail);
  }
  if (res.status === 204) return {};
  return res.json();
}

async function graphPatch(path, body) {
  const token = await getWriteToken();
  const res = await fetch(CONFIG.GRAPH_BASE + path, {
    method: "PATCH",
    headers: {
      Authorization: "Bearer " + token,
      "Content-Type": "application/json",
      "Prefer": "HonorNonIndexedQueriesWarningMayFailRandomly",
    },
    body: JSON.stringify(body)
  });
  if (!res.ok) {
    const err = await res.json().catch(() => ({}));
    throw new Error(err?.error?.message || `Graph error ${res.status} on ${path}`);
  }
  if (res.status === 204) return {};
  return res.json();
}

async function graphDelete(path) {
  const token = await getWriteToken();
  const res = await fetch(CONFIG.GRAPH_BASE + path, {
    method: "DELETE",
    headers: { Authorization: "Bearer " + token }
  });
  if (!res.ok) {
    const err = await res.json().catch(() => ({}));
    throw new Error(err?.error?.message || `Graph error ${res.status}`);
  }
  return {};
}

// Get SharePoint site ID
let _siteId = null;
async function getSiteId() {
  if (_siteId) return _siteId;
  const hostname = new URL(CONFIG.SITE_URL).hostname;
  const sitePath = new URL(CONFIG.SITE_URL).pathname;
  const data = await graphGet(`/sites/${hostname}:${sitePath}`);
  _siteId = data.id;
  return _siteId;
}

// Get list ID by name — always fetches fresh to avoid stale GUID issues
async function getListId(listName) {
  const siteId = await getSiteId();
  const data = await graphGet(`/sites/${siteId}/lists`);
  const match = (data.value || []).find(l =>
    l.displayName === listName ||
    l.name === listName ||
    l.displayName?.toLowerCase() === listName?.toLowerCase()
  );
  if (!match) throw new Error(`List "${listName}" not found on site. Ensure it has been created via form approval.`);
  return match.id;
}

// Get list items
async function getListItems(listName, filter = "", expand = "") {
  const siteId = await getSiteId();
  const listId = await getListId(listName);
  let url = `/sites/${siteId}/lists/${listId}/items?expand=fields,createdBy`;
  if (filter) url += `&$filter=${encodeURIComponent(filter)}`;
  if (expand) url += `&${expand}`;
  const data = await graphGet(url);
  return data.value || [];
}

// Get SharePoint numeric item ID from a Graph item GUID
// The attachment REST API requires the SP numeric ID, not the Graph GUID
async function getSpItemId(listName, graphItemId) {
  const siteId = await getSiteId();
  const listId = await getListId(listName);
  const item = await graphGet(`/sites/${siteId}/lists/${listId}/items/${graphItemId}?$select=id&expand=fields($select=id)`);
  // SP numeric ID is in fields.id
  const spId = item?.fields?.id || item?.fields?.ID;
  if (!spId) throw new Error(`Could not resolve SharePoint item ID for Graph item ${graphItemId}`);
  return parseInt(spId);
}

// Create list item
async function createListItem(listName, fields) {
  const siteId = await getSiteId();
  const listId = await getListId(listName);
  return graphPost(`/sites/${siteId}/lists/${listId}/items`, { fields });
}

// Update list item
async function updateListItem(listName, itemId, fields) {
  const siteId = await getSiteId();
  const listId = await getListId(listName);
  return graphPatch(`/sites/${siteId}/lists/${listId}/items/${itemId}/fields`, fields);
}

// ── JSON Definition Storage ──────────────────────────────────────────────────
// Form definitions are stored as JSON in the FormDefinition column of the
// Forms list. All form lifecycle stages live in this one list.
// ─────────────────────────────────────────────────────────────────────────────

// Save JSON definition into the Forms list item
async function uploadJsonAttachment(listName, itemId, fileName, jsonData) {
  const siteId = await getSiteId();
  const listId = await getListId(listName);
  const json   = JSON.stringify(jsonData);
  await graphPatch(
    `/sites/${siteId}/lists/${listId}/items/${itemId}/fields`,
    { [CONFIG.COL_FORM_DEF]: json }
  );
}

// Read JSON definition from a Forms list item
async function getFormDefinition(listName, itemId) {
  const siteId = await getSiteId();
  const listId = await getListId(listName);
  const item   = await graphGet(`/sites/${siteId}/lists/${listId}/items/${itemId}?expand=fields`);
  const raw    = item?.fields?.[CONFIG.COL_FORM_DEF];
  if (!raw) return null;
  try { return JSON.parse(raw); } catch (_) { return null; }
}

// ── SharePoint REST API Helpers ───────────────────────────────────────────────
// SharePoint REST permissions APIs are not available via Graph — they require
// a token scoped to the SharePoint tenant (not graph.microsoft.com).

async function getSpToken() {
  const spScope = `${new URL(CONFIG.SITE_URL).origin}/.default`;
  try {
    const result = await msalInstance.acquireTokenSilent({
      scopes: [spScope],
      account: currentAccount,
      forceRefresh: false,
    });
    return result.accessToken;
  } catch (e) {
    const result = await msalInstance.acquireTokenPopup({
      scopes: [spScope],
      account: currentAccount,
    });
    return result.accessToken;
  }
}

// GET to SharePoint REST API
async function spGet(relUrl) {
  const token = await getSpToken();
  const url = CONFIG.SITE_URL.replace(/\/$/, "") + relUrl;
  const res = await fetch(url, {
    headers: {
      Authorization: "Bearer " + token,
      Accept: "application/json;odata=verbose",
    },
  });
  if (!res.ok) {
    const txt = await res.text().catch(() => "");
    throw new Error(`SP REST ${res.status} on ${relUrl}: ${txt.slice(0, 300)}`);
  }
  return res.json();
}

// POST to SharePoint REST API
async function spPost(relUrl, body) {
  const token = await getSpToken();
  const url = CONFIG.SITE_URL.replace(/\/$/, "") + relUrl;
  const res = await fetch(url, {
    method: "POST",
    headers: {
      Authorization: "Bearer " + token,
      Accept: "application/json;odata=verbose",
      "Content-Type": "application/json;odata=verbose",
    },
    body: body !== undefined ? JSON.stringify(body) : undefined,
  });
  if (!res.ok) {
    const txt = await res.text().catch(() => "");
    throw new Error(`SP REST ${res.status} on ${relUrl}: ${txt.slice(0, 300)}`);
  }
  if (res.status === 204) return {};
  return res.json().catch(() => ({}));
}

async function spDelete(relUrl) {
  const token = await getSpToken();
  const url = CONFIG.SITE_URL.replace(/\/$/, "") + relUrl;
  const res = await fetch(url, {
    method: "POST",
    headers: {
      Authorization: "Bearer " + token,
      Accept: "application/json;odata=verbose",
      "X-HTTP-Method": "DELETE",
      "If-Match": "*",
    },
  });
  if (!res.ok) {
    const txt = await res.text().catch(() => "");
    throw new Error(`SP DELETE ${res.status} on ${relUrl}: ${txt.slice(0, 300)}`);
  }
  return {};
}

async function spMerge(relUrl, body) {
  const token = await getSpToken();
  const url = CONFIG.SITE_URL.replace(/\/$/, "") + relUrl;
  const res = await fetch(url, {
    method: "POST",
    headers: {
      Authorization: "Bearer " + token,
      Accept: "application/json;odata=verbose",
      "Content-Type": "application/json;odata=verbose",
      "X-HTTP-Method": "MERGE",
      "If-Match": "*",
    },
    body: JSON.stringify(body),
  });
  if (!res.ok) {
    const txt = await res.text().catch(() => "");
    throw new Error(`SP MERGE ${res.status} on ${relUrl}: ${txt.slice(0, 300)}`);
  }
  return {};
}

// Check if current user has read access to a list
async function checkListReadAccess(listName) {
  try {
    await getListItems(listName);
    return true;
  } catch (e) {
    if (e.message && (e.message.includes("403") || e.message.includes("Unauthorized") || e.message.includes("AccessDenied"))) return false;
    throw e;
  }
}

async function checkListWriteAccess(listName) {
  // Attempt to read the list's effective permissions via REST
  // A 400/403 on a write probe means read-only access
  try {
    const siteId = await getSiteId();
    const listId = await getListId(listName);
    const data = await graphGet(`/sites/${siteId}/lists/${listId}?$select=id,name`);
    // Try fetching effective base permissions via SharePoint REST
    const token = await getSpToken();
    const url = CONFIG.SITE_URL.replace(/\/$/, "") + `/_api/web/lists(guid'${listId}')/EffectiveBasePermissions`;
    const res = await fetch(url, {
      headers: { Authorization: "Bearer " + token, Accept: "application/json;odata=verbose" }
    });
    if (!res.ok) return false;
    const perms = await res.json();
    // Low permission bit 2 = AddListItems
    const low = parseInt(perms?.d?.Low || "0");
    return !!(low & 0x2);
  } catch (_) {
    return false;
  }
}

// People search via Graph
async function searchPeople(query) {
  const data = await graphGet(`/me/people?$search="${query}"&$top=10`);
  return data.value || [];
}

// Check admin status
async function checkIsAdmin(userDisplayName, userEmail) {
  try {
    // Fetch all items — no filter to avoid index issues
    const items = await getListItems(CONFIG.FORM_ADMINS_LIST);
    if (CONFIG.DEBUG_LOGGING) console.log("[AdminCheck] FormAdmins items:", JSON.stringify(items.map(i => i.fields)));
    if (CONFIG.DEBUG_LOGGING) console.log("[AdminCheck] Checking user:", userDisplayName, userEmail);

    return items.some(item => {
      const fields = item.fields || {};
      // Approvers is a Person column — array of {LookupId, LookupValue, Email}
      const approvers = Array.isArray(fields.Approvers) ? fields.Approvers : [];
      const matchEmail = approvers.some(a => a.Email?.toLowerCase() === userEmail?.toLowerCase());
      const matchName  = approvers.some(a => a.LookupValue?.toLowerCase() === userDisplayName?.toLowerCase());
      if (CONFIG.DEBUG_LOGGING) console.log("[AdminCheck] Approvers:", JSON.stringify(approvers), "matchEmail:", matchEmail, "matchName:", matchName);
      return matchEmail || matchName;
    });
  } catch (e) {
    console.error("[AdminCheck] Error:", e.message);
    return false;
  }
}
