// Form Studio — Authentication (MSAL)

// =============================================================
// MSAL SETUP
// =============================================================
let msalInstance = null;
let currentAccount = null;
let accessToken = null;

function initMsal() {
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
    cache: { cacheLocation: "sessionStorage", storeAuthStateInCookie: false },
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
    showLoginLoading(true);
    const result = await msalInstance.loginPopup({ scopes: CONFIG.SCOPES });
    currentAccount = result.account;
    accessToken = result.accessToken;
    await bootApp();
  } catch (e) {
    showLoginLoading(false);
    showToast("error", "Sign-in failed: " + e.message);
  }
}

function logout() {
  msalInstance.logoutPopup({ account: currentAccount });
  location.reload();
}
