import * as msal from 'https://cdn.jsdelivr.net/npm/@azure/msal-browser@3.18.0/+esm';

// Multitenant SPA registered in Microsoft Entra ID ("Tally Prime Portal").
// Tenant: 35ba5bb9-3c7b-4758-931e-11d472045692
const APP_CONFIG = {
  clientId: '7cf27726-5ad8-4cc2-be31-2e30b4b0d5d7',
  authority: 'https://login.microsoftonline.com/common',
  // URL opened in a new tab after successful sign-in: AVD web client. User picks the "Tally Prime" tile.
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
const registerRequest = {
  scopes: loginRequest.scopes,
  prompt: 'create',
  extraQueryParameters: { signup: '1' }
};

const signinBtn = document.getElementById('signinBtn');
const registerBtn = document.getElementById('registerBtn');
const launchBtn = document.getElementById('launchBtn');
const signoutBtn = document.getElementById('signoutBtn');
const statusMsg = document.getElementById('statusMsg');

let msalInstance;
let activeAccount = null;

function setStatus(text) { statusMsg.textContent = text || ''; }

function updateUi() {
  const signedIn = !!activeAccount;
  signinBtn.hidden = signedIn;
  registerBtn.hidden = signedIn;
  launchBtn.hidden = !signedIn;
  signoutBtn.hidden = !signedIn;
  if (signedIn) {
    setStatus(`Signed in as ${activeAccount.username}. Click "Open Tally Prime" to launch.`);
  }
}

function openTallyRemoteApp() {
  window.open(APP_CONFIG.avdWebClientUrl, '_blank', 'noopener,noreferrer');
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

registerBtn.addEventListener('click', async () => {
  try {
    setStatus('Redirecting to register...');
    await msalInstance.loginRedirect(registerRequest);
  } catch (e) {
    console.error(e);
    setStatus('Register failed: ' + (e.message || e));
  }
});

launchBtn.addEventListener('click', () => openTallyRemoteApp());

signoutBtn.addEventListener('click', async () => {
  sessionStorage.removeItem('tally_launched');
  await msalInstance.logoutRedirect({ postLogoutRedirectUri: window.location.origin + '/' });
});

init();
