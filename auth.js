// Form Studio — Authentication (MSAL)

// =============================================================
// MSAL SETUP
// =============================================================
let msalInstance = null;
let currentAccount = null;
let accessToken = null;

function initMsal() {
  // Wipe any stale MSAL interaction locks BEFORE initialize() reads localStorage.
  // localStorage survives hard refreshes, so a lock left by a crashed or interrupted
  // interaction (e.g. user closed the popup, or logout fired before MSAL finished)
  // causes "interaction_in_progress" on the next loginPopup call even after a full
  // page reload. MSAL caches this state internally during initialize(), so the cleanup
  // must happen here — before that call — not later in login().
  Object.keys(localStorage)
    .filter(key => key.startsWith("msal.") && key.includes("interaction.status"))
    .forEach(key => localStorage.removeItem(key));

  const msalConfig = {
    auth: {
      clientId: CONFIG.CLIENT_ID,
      authority: `https://login.microsoftonline.com/${CONFIG.TENANT_ID}`,
      // Use a blank redirect URI so the popup closes itself cleanly
      // and the token is never placed in the main window URL.
      // Register "https://login.microsoftonline.com/common/oauth2/nativeclient"
      // as a SPA redirect URI in your app registration.
      // postMessage-based redirect: MSAL opens popup, popup posts token back,
      // main window never navigates. The blank.html just needs to exist at this path.
      redirectUri: window.location.origin + window.location.pathname,
      navigateToLoginRequestUrl: false,
      postLogoutRedirectUri: "about:blank",
    },
    // localStorage persists across tabs and hard refreshes — users stay signed in
    // when opening new tabs and don't need to re-authenticate after a page reload.
    // sessionStorage scopes auth to a single tab and is cleared inconsistently
    // (survives hard refresh but not tab close), which caused stale interaction locks.
    cache: { cacheLocation: "localStorage", storeAuthStateInCookie: false },
    system: {
      allowRedirectInIframe: true,
      windowHashTimeout: 60000,
      iframeHashTimeout: 6000,
      loadFrameTimeout: 0,
    },
  };
  msalInstance = new msal.PublicClientApplication(msalConfig);
  return msalInstance.initialize();
}

async function getToken() {
  const request = { scopes: CONFIG.SCOPES, account: currentAccount };
  try {
    // forceRefresh ensures new scopes (e.g. Sites.Manage.All) are included
    const result = await msalInstance.acquireTokenSilent({ ...request, forceRefresh: false });
    // Check the token actually has our scopes — if not, force refresh
    accessToken = result.accessToken;
    return accessToken;
  } catch (e) {
    try {
      const result = await msalInstance.acquireTokenSilent({ ...request, forceRefresh: true });
      accessToken = result.accessToken;
      return accessToken;
    } catch (e2) {
      const result = await msalInstance.acquireTokenPopup(request);
      accessToken = result.accessToken;
      return accessToken;
    }
  }
}

// Always force-refresh token for write operations to ensure all scopes present
async function getWriteToken() {
  try {
    const result = await msalInstance.acquireTokenSilent({
      scopes: CONFIG.SCOPES,
      account: currentAccount,
      forceRefresh: true,
    });
    accessToken = result.accessToken;
    return accessToken;
  } catch (e) {
    const result = await msalInstance.acquireTokenPopup({ scopes: CONFIG.SCOPES, account: currentAccount });
    accessToken = result.accessToken;
    currentAccount = result.account;
    return accessToken;
  }
}

// Force a fresh token — call this after adding new scopes to clear cache
async function forceRefreshToken() {
  try {
    msalInstance.clearCache();
  } catch (_) {}
  const request = { scopes: CONFIG.SCOPES, account: currentAccount };
  const result = await msalInstance.acquireTokenPopup(request);
  accessToken = result.accessToken;
  currentAccount = result.account;
  return accessToken;
}

async function login() {
  try {
    // Safety net: clear any stale MSAL interaction locks from sessionStorage.
    // MSAL stores interaction state under keys prefixed with "msal.".
    // A stale lock (e.g. from a crashed session or a missed await on logout)
    // causes the "interaction_in_progress" error and prevents loginPopup from opening.
    Object.keys(localStorage)
      .filter(key => key.startsWith("msal."))
      .forEach(key => localStorage.removeItem(key));

    showLoginLoading(true);
    // Show a locked modal so the user knows a popup is expected and cannot
    // interact with anything behind it while the sign-in handshake completes.
    showProgress("Signing in", "A Microsoft sign-in window has opened — please complete sign-in there.");
    const result = await msalInstance.loginPopup({ scopes: CONFIG.SCOPES });
    currentAccount = result.account;
    accessToken = result.accessToken;
    updateProgress("Signed in. Loading your workspace…");
    await bootApp();
    // bootApp() replaces the entire shell, so no explicit hideProgress() needed.
  } catch (e) {
    hideProgress();
    showLoginLoading(false);
    showToast("error", "Sign-in failed: " + e.message);
  }
}

async function logout() {
  // Show a locked blocking modal immediately so nothing can be clicked while
  // the logout popup is open and MSAL is cleaning up its session state.
  showProgress("Signing out", "A Microsoft sign-out window has opened — please wait…");
  try {
    // await the popup so MSAL finishes its cleanup before we reload.
    // Without await, location.reload() fires immediately and tears down
    // the page while MSAL's interaction lock is still written to localStorage,
    // causing "interaction_in_progress" on the next login attempt.
    await msalInstance.logoutPopup({ account: currentAccount });
  } catch (_) {
    // If the user closes the popup or an error occurs, we still want
    // to reload and clear local state — so we swallow the error.
  }
  hideProgress();
  location.reload();
}