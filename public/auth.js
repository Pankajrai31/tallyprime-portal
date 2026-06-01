import * as msal from 'https://cdn.jsdelivr.net/npm/@azure/msal-browser@3.18.0/+esm';

// Multi-tenant: accepts any Microsoft Entra ID work or school account.
// Replace clientId after registering this site as an Entra app (Web platform, redirectUri = current origin).
const APP_CONFIG = {
  // Public client app id placeholder. Replace with real Entra App Registration (Application (client) ID).
  clientId: window.__TALLY_CLIENT_ID__ || '00000000-0000-0000-0000-000000000000',
  authority: 'https://login.microsoftonline.com/common',
  // URL opened in a new tab after successful sign-in. This is the Azure Virtual Desktop web client.
  // The user will then see the "Tally Prime" RemoteApp tile in their AVD workspace.
  avdWebClientUrl: 'https://client.wvd.microsoft.com/arm/webclient/index.html'
};

const msalConfig = {
  auth: {
    clientId: APP_CONFIG.clientId,
    authority: APP_CONFIG.authority,
    redirectUri: window.location.origin + '/'
  },
  cache: { cacheLocation: 'sessionStorage', storeAuthStateInCookie: false }
};

const loginRequest = { scopes: ['openid', 'profile', 'email', 'User.Read'] };

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
    setStatus(`Signed in as ${activeAccount.username}. Click "Open Tally Prime" to launch.`);
  }
}

function openTallyRemoteApp() {
  // Open the AVD web client in a new tab. The user sees the Tally Prime RemoteApp tile there.
  window.open(APP_CONFIG.avdWebClientUrl, '_blank', 'noopener,noreferrer');
}

async function init() {
  if (APP_CONFIG.clientId === '00000000-0000-0000-0000-000000000000') {
    setStatus('Sign-in is not yet configured. Register an Entra App and set its client id.');
  }

  msalInstance = new msal.PublicClientApplication(msalConfig);
  await msalInstance.initialize();

  // Handle redirect response (if any).
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

  // Auto-open AVD client right after a fresh sign-in (one-time per session).
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
