# Claudette-Teams — Azure setup

One-time provisioning to get the bot live in Microsoft Teams. You need an Azure
subscription and Teams admin (or custom-app-upload) rights in the tenant.

## 1. App registration + Azure Bot resource

1. **Azure Portal → Microsoft Entra ID → App registrations → New registration.**
   Single tenant. Note the **Application (client) ID** and **Directory (tenant) ID**.
2. **Certificates & secrets → New client secret.** Copy the secret value now.
3. **Azure Portal → Create a resource → Azure Bot.**
   - Type of App: **Single Tenant**, reuse the app registration above (Microsoft App ID = client ID).
   - After creation, **Channels → add Microsoft Teams**.

## 2. App Service (hosts the Python code)

1. **Create an App Service** — Runtime **Python 3.11+**, OS **Linux**.
2. **Configuration → Application settings**, add:

   ```
   CONNECTIONS__SERVICE_CONNECTION__SETTINGS__CLIENTID=<app client id>
   CONNECTIONS__SERVICE_CONNECTION__SETTINGS__CLIENTSECRET=<client secret>
   CONNECTIONS__SERVICE_CONNECTION__SETTINGS__TENANTID=<tenant id>
   PROJECT_DIR=/home/site/wwwroot           # or wherever analysis runs
   AUTHORIZED_USERS=<your AAD object id>     # NOT display name; see below
   SHAREPOINT_SITE_ID=<optional, for channel file uploads>
   ```

   > Find your AAD object id: Entra ID → Users → your user → **Object ID**. The bot
   > authorizes on `activity.from.aadObjectId`, which equals this value.

3. **Startup command:** `python bot.py` (the app serves `/api/messages` on `$PORT`).
   App Service injects `PORT`; the code reads it.
4. **Deploy** the repo (zip deploy, `az webapp up`, or GitHub Actions).

   ```powershell
   az webapp up --name <app-name> --runtime "PYTHON:3.11"
   ```

5. **Azure Bot resource → Configuration → Messaging endpoint:**
   `https://<app-name>.azurewebsites.net/api/messages`
6. Verify with **Azure Bot → Test in Web Chat**.

## 3. Channel file uploads (optional, full parity)

1:1 chats use the FileConsentCard flow and need nothing extra. For uploads in
**channels/group chats**:

1. **App registration → API permissions →** add **Microsoft Graph → Application →
   `Sites.ReadWrite.All`**, then **Grant admin consent**.
2. Find the target SharePoint **site id** (`GET /sites/{hostname}:/sites/{path}`) and
   set `SHAREPOINT_SITE_ID`.

## 4. Package + install the Teams app

1. Edit `appManifest/manifest.json`: replace `<<AAD_APP_CLIENT_ID>>` (3 places) and
   `<<BOT_DOMAIN>>`. Add `color.png` (192×192) and `outline.png` (32×32).
2. Zip the three files into `manifest.zip` (see `appManifest/README.txt`).
3. **Teams Admin Center → Teams apps → Manage apps → Upload**, or for a quick test:
   **Teams → Apps → Manage your apps → Upload a custom app.**

## 5. Smoke test

DM the bot in Teams. You should see a typing indicator, then a streamed reply.
`@mention` it in a channel it's installed in. Send a file to test inbound download;
ask it to produce a file (it emits `attach:<path>`) to test the consent-card upload.

## Local development (Bot Framework Emulator)

```powershell
pip install -r requirements.txt
# In .env, set ...__ANONYMOUS_ALLOWED=True and leave CLIENTID/SECRET blank.
python bot.py
```

Point the Emulator at `http://localhost:3978/api/messages` with empty App ID/password.
Subprocess streaming works on Windows' default Proactor event loop.
