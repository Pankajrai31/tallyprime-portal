import * as msal from 'https://cdn.jsdelivr.net/npm/@azure/msal-browser@3.18.0/+esm';

// Multitenant SPA registered in Microsoft Entra ID ("Tally Prime Portal").
// Tenant: 35ba5bb9-3c7b-4758-931e-11d472045692
const APP_CONFIG = {
  clientId: '7cf27726-5ad8-4cc2-be31-2e30b4b0d5d7',
  // Resource tenant that owns the AVD host pool / workspace.
  // We authenticate DIRECTLY against this tenant (not /common) so the AAD session cookie
  // is set for Pankajrai87outlook. When the user then opens client.wvd.microsoft.com it
  // silently reuses that session and shows the Tally Prime Workspace. Otherwise the AVD
  // client signs the user into their home tenant and shows the wrong (or empty) feed.
  resourceTenantId: '35ba5bb9-3c7b-4758-931e-11d472045692',
  authority: 'https://login.microsoftonline.com/35ba5bb9-3c7b-4758-931e-11d472045692',
  // URL opened in a new tab after successful sign-in: AVD web client. User picks the "Tally Prime" tile.
  avdWebClientUrl: 'https://client.wvd.microsoft.com/arm/webclient/index.html'
};

const msalConfig = {
  auth: {
    clientId: APP_CONFIG.clientId,
    authority: APP_CONFIG.authority,
    knownAuthorities: ['login.microsoftonline.com'],
    redirectUri: window.location.origin + '/'
  },
  cache: { cacheLocation: 'sessionStorage', storeAuthStateInCookie: false }
};

const loginRequest = {
  scopes: ['openid', 'profile', 'email', 'User.Read'],
  // Force Entra to show the account picker so users do not get auto-signed into
  // their home tenant. This is critical for B2B guests (panrai@microsoft.com etc.)
  prompt: 'select_account'
};

const signinBtn = document.getElementById('signinBtn');
const launchBtn = document.getElementById('launchBtn');
const signoutBtn = document.getElementById('signoutBtn');
const statusMsg = document.getElementById('statusMsg');

let msalInstance;
let activeAccount = null;

function setStatus(text) { statusMsg.textContent = text || ''; }

function updateUi() {
  const signedIn = !!activeAccount;
  signinBtn.hidden = signedIn;
  launchBtn.hidden = !signedIn;
  signoutBtn.hidden = !signedIn;
  if (signedIn) {
    setStatus(`Signed in as ${activeAccount.username}. Opening Tally Prime…`);
  }
}

function openTallyRemoteApp() {
  // Pin the AVD client to the resource tenant. Because we authenticated against the
  // same tenant in this browser, AAD silently SSOs the user — no second sign-in,
  // no "wrong tenant" issue.
  const url = `${APP_CONFIG.avdWebClientUrl}?tenantId=${encodeURIComponent(APP_CONFIG.resourceTenantId)}`;
  window.open(url, '_blank', 'noopener,noreferrer');
}

async function init() {
  msalInstance = new msal.PublicClientApplication(msalConfig);
  await msalInstance.initialize();

  try {
    const resp = await msalInstance.handleRedirectPromise();
    if (resp && resp.account) {
      activeAccount = resp.account;
      msalInstance.setActiveAccount(activeAccount);
    }
  } catch (e) {
    console.error(e);
    setStatus('Sign-in failed: ' + (e.message || e));
  }

  if (!activeAccount) {
    const accounts = msalInstance.getAllAccounts();
    if (accounts.length > 0) {
      activeAccount = accounts[0];
      msalInstance.setActiveAccount(activeAccount);
    }
  }

  if (activeAccount && !sessionStorage.getItem('tally_launched')) {
    sessionStorage.setItem('tally_launched', '1');
    openTallyRemoteApp();
  }

  updateUi();
}

signinBtn.addEventListener('click', async () => {
  try {
    setStatus('Redirecting to sign in...');
    await msalInstance.loginRedirect(loginRequest);
  } catch (e) {
    console.error(e);
    setStatus('Sign-in failed: ' + (e.message || e));
  }
});

launchBtn.addEventListener('click', () => openTallyRemoteApp());

signoutBtn.addEventListener('click', async () => {
  sessionStorage.removeItem('tally_launched');
  await msalInstance.logoutRedirect({ postLogoutRedirectUri: window.location.origin + '/' });
});

init();
