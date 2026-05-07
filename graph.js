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
  if (res.status === 401) {
    // Session expired — redirect to login
    await msalInstance.loginRedirect({ scopes: CONFIG.SCOPES });
    return;
  }
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
  if (res.status === 401) {
    await msalInstance.loginRedirect({ scopes: CONFIG.SCOPES });
    return;
  }
  if (!res.ok) {
    const txt = await res.text().catch(() => "");
    console.error("[graphPost error]", res.status, path, txt);
    let err = {};
    try { err = JSON.parse(txt); } catch (_) {}
    const detail = err?.error?.innererror?.message || err?.error?.message || `Graph error ${res.status} on ${path}`;
    throw new Error(detail);
  }
  // 202 Accepted (e.g. /me/sendMail) and 204 No Content both return empty bodies
  if (res.status === 204 || res.status === 202) return {};
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
      "If-Match": "*",
    },
    body: JSON.stringify(body)
  });
  if (res.status === 401) {
    await msalInstance.loginRedirect({ scopes: CONFIG.SCOPES });
    return;
  }
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
  if (res.status === 401) {
    await msalInstance.loginRedirect({ scopes: CONFIG.SCOPES });
    return;
  }
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

// POST a file (binary/text) to SharePoint REST API
async function spPut(relUrl, body) {
  const token = await getSpToken();
  const url = CONFIG.SITE_URL.replace(/\/$/, "") + relUrl;
  const res = await fetch(url, {
    method: "POST",
    headers: {
      Authorization: "Bearer " + token,
      Accept: "application/json;odata=verbose",
      "Content-Type": "application/octet-stream",
    },
    body,
  });
  if (!res.ok) {
    const txt = await res.text().catch(() => "");
    throw new Error(`SP PUT ${res.status} on ${relUrl}: ${txt.slice(0, 300)}`);
  }
  return res.json().catch(() => ({}));
}

// ── JSON Definition Storage ──────────────────────────────────────────────────
// Form definitions are stored as a JSON file attachment on the Forms list item.
// SP list item attachments are used rather than a column value because the JSON
// can grow very large (e.g. dropdowns with 1000+ entries) and would exceed the
// SharePoint column size limit (~64KB).
//
// The attachment filename is always "form-definition.json".
// getFormDefinition falls back to the legacy FormDefinition column for forms
// saved before this migration — so existing forms continue to work.
// ─────────────────────────────────────────────────────────────────────────────

// Save JSON definition as a file attachment on the Forms list item.
// Deletes any existing "form-definition.json" attachment first (SP returns 409
// Conflict if you try to add a duplicate filename).
async function uploadJsonAttachment(listName, graphItemId, fileName, jsonData) {
  const spItemId  = await getSpItemId(listName, graphItemId);
  const siteUrl   = CONFIG.SITE_URL.replace(/\/$/, "");
  const token     = await getSpToken();
  const baseUrl   = `${siteUrl}/_api/web/lists/GetByTitle('${encodeURIComponent(listName)}')/items(${spItemId})/AttachmentFiles`;
  const encoded   = encodeURIComponent(fileName);
  const jsonBytes = new TextEncoder().encode(JSON.stringify(jsonData));

  // Delete existing attachment — SP doesn't support overwrite.
  // SP REST requires X-HTTP-Method: DELETE and If-Match: * headers.
  try {
    await fetch(`${baseUrl}/getByFileName('${encoded}')`, {
      method: "POST",
      headers: {
        Authorization: "Bearer " + token,
        Accept: "application/json;odata=verbose",
        "X-HTTP-Method": "DELETE",
        "If-Match": "*",
      },
    });
  } catch (_) {} // 404 = didn't exist — fine

  // Upload new attachment
  const res = await fetch(`${baseUrl}/add(FileName='${encoded}')`, {
    method: "POST",
    headers: {
      Authorization: "Bearer " + token,
      Accept: "application/json;odata=verbose",
      "Content-Type": "application/octet-stream",
    },
    body: jsonBytes,
  });
  if (!res.ok) {
    const txt = await res.text().catch(() => "");
    throw new Error(`Definition upload failed: ${res.status} — ${txt.slice(0, 200)}`);
  }

  // Clear the legacy column value to free space — not critical if it fails
  try {
    const siteId = await getSiteId();
    const listId = await getListId(listName);
    await graphPatch(`/sites/${siteId}/lists/${listId}/items/${graphItemId}/fields`, {
      [CONFIG.COL_FORM_DEF]: "",
    });
  } catch (_) {}
}

// Read JSON definition — tries the attachment first, falls back to the legacy
// FormDefinition column for forms saved before the attachment migration.
async function getFormDefinition(listName, itemId) {
  // ── Try attachment first ──────────────────────────────────────
  try {
    const spItemId = await getSpItemId(listName, itemId);
    const siteUrl  = CONFIG.SITE_URL.replace(/\/$/, "");
    const token    = await getSpToken();
    const res = await fetch(
      `${siteUrl}/_api/web/lists/GetByTitle('${encodeURIComponent(listName)}')/items(${spItemId})/AttachmentFiles/getByFileName('form-definition.json')/$value`,
      { headers: { Authorization: "Bearer " + token, Accept: "application/json" } }
    );
    if (res.ok) {
      const text = await res.text();
      try { return JSON.parse(text); } catch (_) {}
    }
  } catch (_) {}

  // ── Fall back to legacy column ────────────────────────────────
  try {
    const siteId = await getSiteId();
    const listId = await getListId(listName);
    const item   = await graphGet(`/sites/${siteId}/lists/${listId}/items/${itemId}?expand=fields`);
    const raw    = item?.fields?.[CONFIG.COL_FORM_DEF];
    if (!raw) return null;
    return JSON.parse(raw);
  } catch (_) { return null; }
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

// Resolve a user's SharePoint integer ID from their email.
// SharePoint Person columns written via the Graph API require a SP numeric user ID
// (as ColumnNameLookupId), not an Entra object ID or email string.
// Results are cached in memory for the session to avoid redundant lookups.
const _spUserIdCache = {};
async function resolveSpUserId(email) {
  if (!email) return null;
  if (_spUserIdCache[email] !== undefined) return _spUserIdCache[email];
  try {
    const siteId = await getSiteId();
    const data = await graphGet(
      `/sites/${siteId}/lists('User Information List')/items?$filter=fields/EMail eq '${email}'&$expand=fields&$select=id,fields`
    );
    const item = data.value?.[0];
    // fields.ID is the SharePoint integer user ID — item.id is the Graph GUID (not useful here)
    const rawId = item?.fields?.ID ?? item?.fields?.id;
    const spId = rawId != null ? parseInt(rawId, 10) : null;
    _spUserIdCache[email] = spId;
    return spId;
  } catch (e) {
    console.warn("[resolveSpUserId] Could not resolve SP user ID for", email, e.message);
    _spUserIdCache[email] = null;
    return null;
  }
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