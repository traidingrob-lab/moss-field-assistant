// Microsoft sign-in for OneDrive access (Phase 2). Uses MSAL.js
// (vendored in js/vendor/msal-browser.min.js — no external CDN, so our
// strict script-src 'self' CSP doesn't need to be loosened for it).
//
// "consumers" authority is for personal Microsoft accounts only, matching
// the "Personal accounts only" option chosen when registering the app in
// Azure. (If the app registration is ever changed to also allow
// work/school accounts, this needs to change to "common".)
const MSAL_CONFIG = {
  auth: {
    clientId: MSAL_CLIENT_ID,
    authority: "https://login.microsoftonline.com/consumers",
    // Hardcoded (not derived from location.*) so it always exactly matches
    // the redirect URI registered in Azure, regardless of how the page was
    // reached (with/without a trailing hash route, etc).
    redirectUri: "https://traidingrob-lab.github.io/moss-field-assistant/"
  },
  cache: {
    cacheLocation: "localStorage",
    storeAuthStateInCookie: false
  }
};

const GRAPH_SCOPES = ["Files.ReadWrite", "User.Read"];

let msalInstance = null;
let msalReadyPromise = null;

function msalConfigured() {
  return typeof MSAL_CLIENT_ID === "string" && MSAL_CLIENT_ID && !MSAL_CLIENT_ID.startsWith("REPLACE_");
}

// Idempotent — safe to call from every screen render. Handles the
// redirect-back leg of the login flow (MSAL reads the auth response out
// of the URL/session storage after Microsoft sends the browser back here).
function initMsal() {
  if (!msalConfigured()) return Promise.resolve();
  if (msalReadyPromise) return msalReadyPromise;

  msalInstance = new msal.PublicClientApplication(MSAL_CONFIG);
  msalReadyPromise = msalInstance
    .initialize()
    .then(() => msalInstance.handleRedirectPromise())
    .then((result) => {
      if (result?.account) {
        msalInstance.setActiveAccount(result.account);
      } else if (!msalInstance.getActiveAccount()) {
        const accounts = msalInstance.getAllAccounts();
        if (accounts.length) msalInstance.setActiveAccount(accounts[0]);
      }
    })
    .catch((err) => {
      console.error("MSAL init failed", err);
    });
  return msalReadyPromise;
}

async function msSignIn() {
  await initMsal();
  if (!msalInstance) return;
  await msalInstance.loginRedirect({ scopes: GRAPH_SCOPES });
}

async function msSignOut() {
  await initMsal();
  if (!msalInstance) return;
  await msalInstance.logoutRedirect();
}

function msCurrentAccount() {
  if (!msalInstance) return null;
  return msalInstance.getActiveAccount();
}

// Returns a Graph access token, or null if not signed in. Falls back to
// an interactive redirect if the silent refresh fails (expired session,
// revoked consent, etc.) rather than throwing mid-capture.
async function msGetToken() {
  await initMsal();
  if (!msalInstance) return null;
  const account = msalInstance.getActiveAccount();
  if (!account) return null;
  try {
    const result = await msalInstance.acquireTokenSilent({ scopes: GRAPH_SCOPES, account });
    return result.accessToken;
  } catch (err) {
    console.warn("Silent token refresh failed, redirecting to sign in again", err);
    await msalInstance.acquireTokenRedirect({ scopes: GRAPH_SCOPES });
    return null;
  }
}
