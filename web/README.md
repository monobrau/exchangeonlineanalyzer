# Exchange Online Analyzer — Web API

FastAPI service for **async bulk export jobs**, Authentik-ready (OIDC JWT), SQLite job store (swap for PostgreSQL later).

## Run (development)

```bash
cd web
python -m venv .venv
. .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -r requirements.txt
uvicorn app.main:app --reload --host 0.0.0.0 --port 8080
python -m pytest tests/ -q
```

- **Browser UI:** `http://127.0.0.1:8080/` — Microsoft 365 **Add client** sign-in, app registration tools, bulk job form, optional **client queue** (one job per tenant GUID), **Run activity** cards, **Run again** (loads `request_payload`)
- Health: `GET http://127.0.0.1:8080/health`
- Readiness: `GET http://127.0.0.1:8080/ready` (checks DB)
- OpenAPI: `http://127.0.0.1:8080/docs`

With Authentik enabled, open the app in the browser and set `sessionStorage.setItem('eoa_bearer', '<jwt>')` in devtools (until interactive OIDC login is added), or leave OIDC off for local use.

## Endpoints (v1)

| Method | Path | Notes |
|--------|------|--------|
| GET | `/health` | Liveness |
| GET | `/api/v1/ui-info` | **Live deploy check:** files under `web/static` this process reads (`index_html.has_ms_graph_outer`, etc.); no auth |
| GET | `/api/v1/export/options-schema` | JSON Schema for `options` on `POST /api/v1/jobs/bulk` (parity with desktop bulk exporter); see `web/docs/bulk-export-parity.md` |
| GET | `/api/v1/me` | Current user `sub` when `EOA_OIDC_ISSUER` is set |
| GET | `/api/v1/connections/status` | Graph/EXO configured (server env) + optional default job tenant GUID |
| GET | `/api/v1/jobs` | List jobs |
| POST | `/api/v1/jobs/bulk` | Create bulk export job (body: `tenant_ids`, `options`) |
| GET | `/api/v1/jobs/{uuid}` | Job status and `request_payload` (tenant_ids + options snapshot for **Run again**) |
| GET | `/api/v1/jobs/{uuid}/artifacts` | List artifact filenames for that job |
| GET | `/api/v1/jobs/{uuid}/artifact?file=summary.json` | Download file from `web/data/artifacts/{uuid}/` (requires `succeeded`) |

Job JSON (`GET` list/detail) includes **`artifact_files`** when status is `succeeded`.

**CORS:** set **`EOA_CORS_ORIGINS`** to a comma-separated allowlist for production; default `*` disables credential cookies in browsers (fine for local dev).

Without `EOA_OIDC_ISSUER`, API routes are open for local development. Set issuer + audience to enforce Bearer tokens.

### Browser sign-in (Authentik / OIDC + PKCE)

1. In Authentik, create an **OAuth2/OpenID** provider and application. Add redirect URI exactly: **`EOA_OIDC_REDIRECT_URI`** (e.g. `http://127.0.0.1:8080/api/v1/auth/oidc/callback` or your public URL).
2. Set **`EOA_OIDC_ISSUER`**, **`EOA_OIDC_CLIENT_ID`**, **`EOA_OIDC_REDIRECT_URI`**, **`EOA_OIDC_AUDIENCE`** (often the client ID). Use **`EOA_OIDC_CLIENT_SECRET`** only if the provider is confidential.
3. Set **`EOA_SESSION_SECRET`** to a long random string (signs the session cookie used for PKCE `state` / verifier).
4. Open the UI: **Sign in** goes to **`/api/v1/auth/oidc/login`**, then **callback** stores the access token in **`sessionStorage`** as **`eoa_bearer`** for API calls.

Use **`EOA_CORS_ORIGINS`** with your real UI origin when not same-host.

## Environment

See `.env.example`. Database defaults to `web/data/eoa_jobs.db`.

### Settings in the browser

The header **Settings** button opens a panel that reads the **effective** configuration (base `web/.env` plus overrides). **Save** writes **`web/data/eoa_gui.env`**, which overrides the same keys from `.env` for this process after save (settings cache is reloaded). The **Microsoft 365 — browser sign-in (MSAL)** section includes **SPA client ID**, **authority tenant**, and **delegated scopes** so you can enable and tune M365 auth without editing files. Secrets are never shown in full; use **Remove override** to drop a gui-only secret and fall back to `.env`. **CORS** and **session secret** still require an **API process restart** to apply to middleware. `EOA_DATABASE_URL` / app name are not exposed in the GUI—edit `web/.env` manually if needed.

### PowerShell stub worker (Linux / webhost)

1. Install **PowerShell 7** (`pwsh`) and ensure it is on `PATH` (or set **`EOA_PWSH_PATH`**).
2. Set **`EOA_USE_PWSH_STUB_WORKER=true`** so each job runs a pwsh script (default **`web/pwsh/WebBulkJobStub.ps1`** — writes **`ReportSelections.json`** + **`summary.json`**). For **Exchange Online + Graph** using the same **`New-SecurityInvestigationReport`** path as the desktop exporter, set **`EOA_PWSH_WORKER_SCRIPT=WebExoLinuxRunner.ps1`** and configure **EXO + Graph app-only** env vars (no secrets in job JSON). See **[`docs/exo-linux-runner.md`](docs/exo-linux-runner.md)**.
3. Set **`EOA_REPO_ROOT`** to the repo root if the API working directory is not the repo (default: parent of `web/`).

If **`EOA_USE_PWSH_STUB_WORKER`** is false (default for local dev), the API uses a short **in-process placeholder** only (unless the Python Graph worker is enabled below).

**Worker order:** by default, **PowerShell runs first** (`web/pwsh/WebBulkJobStub.ps1`); the Python Graph worker runs only if pwsh is disabled or `pwsh` is not on `PATH`. To run **Graph first** when both workers are enabled (so `summary.json` and `report_*.json` come from Graph and the stub writes `pwsh_summary.json`), set **`EOA_PYTHON_GRAPH_BEFORE_PWSH=true`** (requires **`EOA_USE_PYTHON_GRAPH_WORKER=true`** and **`EOA_GRAPH_*`** app credentials). See [`docs/first-report-milestone.md`](docs/first-report-milestone.md).

### Python Graph worker (Linux, optional fallback)

1. Register an Entra app with a **client secret** and grant **application** permissions for the reports you need. **Admin-consent** the app in the target tenant. Examples:
   - **organization** — `Organization.Read.All`
   - **users** — `User.Read.All`
   - **conditional_access** — `Policy.Read.All`
   - **applications** — `Application.Read.All`
   - **sign_in_logs** — `AuditLog.Read.All`
   - **directory_audits** — `AuditLog.Read.All`
   - **security_alerts** / **security_incidents** — `SecurityEvents.Read.All` / `SecurityIncident.Read.All` (or product-specific equivalents)
   - **intune_devices** — `DeviceManagementManagedDevices.Read.All`
   - **mfa_registration** — `Reports.Read.All`
2. Job body **`options`** can list **`reports`** and/or set **`include_*`** booleans (same names as the web schema); the worker merges **`include_*`** into the Graph report list (see `GET /api/v1/export/options-schema`).
3. Set **`EOA_GRAPH_CLIENT_ID`**, **`EOA_GRAPH_CLIENT_SECRET`**, and **`EOA_USE_PYTHON_GRAPH_WORKER=true`** when you want Graph-based jobs **without** pwsh (or as fallback when pwsh is missing).
4. **Smoke-test a report locally** (writes under `web/data/artifacts/<job-id>/`):

```bash
cd web
# Windows PowerShell: $env:EOA_GRAPH_CLIENT_ID='...'; $env:EOA_GRAPH_CLIENT_SECRET='...'
export EOA_GRAPH_CLIENT_ID=...
export EOA_GRAPH_CLIENT_SECRET=...
python tools/run_graph_report.py <YOUR-TENANT-GUID> organization
# optional: python tools/run_graph_report.py <TENANT-GUID> organization users
```

Exit code **0** means every requested report succeeded; **1** means at least one failed (check `report_*.json` and `worker.log`). **2** means credentials were not set.

### Console vs API

The **web console** targets **interactive** Microsoft sign-in for the active directory and **checkbox** options (parity with desktop report toggles). **`POST /api/v1/jobs/bulk`** still accepts arbitrary **`tenant_ids`** for **automated** callers (e.g. Python Graph worker with app-only creds). See **[`docs/multi-tenant-scaling.md`](docs/multi-tenant-scaling.md)**.

The API **injects** `window.__EOA_MSAL_BOOTSTRAP__` into **`/`** and **`/app`** HTML (same resolved `EOA_MS_GRAPH_SPA_CLIENT_ID` as `/api/v1/auth/msal-config`) so **Add client / Sign in** works after you configure the SPA id and **reload**, without depending on the static file path alone.

### Microsoft sign-in (browser) — app registrations + tenant for jobs

This is **separate** from Authentik/API auth: the header **Sign in** still controls access to `/api/v1/jobs` when `EOA_OIDC_ISSUER` is set. The **Microsoft 365** panel uses **MSAL** in the browser to sign in with a work account and call **Microsoft Graph** directly (delegated). **Creating app registrations** uses `POST https://graph.microsoft.com/v1.0/applications` from the browser with the signed-in user’s token.

**OAuth always needs an Entra app registration** (a public “client ID”). You can supply it in any of these ways:

1. **Same app as the Graph worker (default)** — **`EOA_MS_GRAPH_SPA_USE_GRAPH_APP_ID`** defaults to **true**. When **`EOA_GRAPH_CLIENT_ID`** is set and that Entra app has a **Single-page application** platform with redirect URIs for this host (`/` and `/app`), browser **Sign in with Microsoft** works **without** a separate **`EOA_MS_GRAPH_SPA_CLIENT_ID`** or Settings change. Add the SPA platform and redirects in Entra if you only had a confidential / worker registration before. Set **`EOA_MS_GRAPH_SPA_USE_GRAPH_APP_ID=false`** only if the Graph app cannot be used as an SPA.  
2. **`EOA_MS_GRAPH_SPA_CLIENT_ID`** in `web/.env` (or systemd) — use when the interactive SPA is a different registration than the Python Graph worker.  
3. **Bundled default** — set **`BUNDLED_MS_GRAPH_SPA_CLIENT_ID`** in **`web/app/bundled_ms_graph.py`** (or run **`web/tools/Register-EoaMsalSpaApp.ps1`** once and paste the printed ID) so the API enables MSAL without env.  
4. **Browser only** — paste the **Application (client) ID** once in the Microsoft 365 panel; it is stored in **`localStorage`** (`eoa_ms_graph_spa_client_id`). Use **Clear browser-stored client ID** to reset.

Then:

1. In **Entra ID** → **App registrations** → **New registration** (or use an existing app).  
   - **Supported account types**: use **this directory only** (single-tenant) for one org or your own deployment; use **any organizational directory** (multitenant) only if you ship one client ID to many customer tenants (ISV). The helper script **`Register-EoaMsalSpaApp.ps1`** defaults to single-tenant; pass **`-Multitenant`** for `AzureADMultipleOrgs`.  
   - **Redirect URI**: platform **Single-page application (SPA)**. Add every URL users will open, exactly (path matters):  
     - `https://<your-host>/`  
     - `https://<your-host>/app`  
     - For local dev: `http://127.0.0.1:8080/`, `http://127.0.0.1:8080/app`  
2. **API permissions** → **Microsoft Graph** → **Delegated permissions**: add **`User.Read`**, **`Organization.Read.All`**, **`Application.ReadWrite.All`**.  
   - **`Application.ReadWrite.All`** almost always requires **Grant admin consent for \<tenant\>** (or an admin consent workflow). Without it, list/create/delete app registration calls return **403**.  
3. Optionally set **`EOA_MS_GRAPH_TENANT`** (default **`organizations`** for work accounts) when using a server-side client ID.  
4. Optional: **`EOA_MS_GRAPH_DELEGATED_SCOPES`** — comma-separated delegated Graph scopes for the browser (empty = default `User.Read`, `Organization.Read.All`, `Application.ReadWrite.All`). Set this in **Settings** or `.env` / `eoa_gui.env` when you need different or extra permissions (admin consent may be required).  
5. Open **`/`** or **`/app`**, expand **Microsoft 365**, **Sign in with Microsoft** (after pasting client ID if prompted). After an admin has consented permissions, use **Re-consent permissions** once if you still see consent or authorization errors.  
6. **App registrations** table: **Create** (display name), **Refresh list**, **Rename**, **Delete** — all call Graph interactively; no server-side Graph secret is used for this panel.  
7. **Bulk jobs**: **Microsoft 365 interactive sign-in** supplies the Entra **tenant** for jobs when the queue is empty (no manual default GUID required). The **client queue** creates **one job per tenant**; optional **`EOA_JOB_DEFAULT_TENANT_ID`** is a fallback only. **GET** `/api/v1/jobs/{id}` includes **`request_payload`** for **Run again**.

## Deploy (Linux)

After **`git pull`**, install Python deps in the same venv systemd uses:

```bash
cd web && .venv/bin/pip install -r requirements.txt && sudo systemctl restart eoa-api
```

HTML/CSS/JS responses include **`CDN-Cache-Control`** and asset URLs use **`?v=<api version>`** so each deploy changes filenames. If the public URL still looks stale, **purge Cloudflare cache** for the hostname once (browser hard refresh is not enough when the edge cached HTML).

Examples (adjust paths/users):

- **[`deploy/eoa-api.service.example`](deploy/eoa-api.service.example)** — systemd unit for `uvicorn` on `127.0.0.1:8080`.
- **[`deploy/nginx-snippet.conf.example`](deploy/nginx-snippet.conf.example)** — reverse proxy snippet for nginx in front of the API.
