# Tally Prime Portal

Public sign-in landing page for **Tally Prime** delivered as an Azure Virtual Desktop RemoteApp.

- Static front-end served by Express on Azure App Service (Linux, Node 20).
- Uses MSAL.js multi-tenant (`/common`) so any Microsoft Entra ID work or school account can sign in.
- After successful sign-in, opens the AVD web client (`https://client.wvd.microsoft.com/arm/webclient/index.html`) in a new tab.
- Users see the **Tally Prime** RemoteApp tile (published from the `tally-remoteapp-ag` application group) and launch it.

## Local dev

```powershell
cd website
npm install
npm start
# open http://localhost:8080
```

## Configure sign-in

1. Register a new Microsoft Entra **App Registration** (Single tenant or Multitenant).
2. Platform: **Single-page application**. Redirect URI = your App Service URL (e.g. `https://tallyprime-portal-xxxx.azurewebsites.net/`).
3. Copy the **Application (client) ID** and inject it as `window.__TALLY_CLIENT_ID__` (via an App Service app setting that emits a script tag) or hard-code it in `public/auth.js`.

## Deploy

Deployment is wired up to GitHub via Azure App Service continuous deployment.
Pushes to `main` trigger an automatic build and release.
