# Exchange Online Analyzer — Web API

FastAPI service for **async bulk export jobs**, Authentik-ready (OIDC JWT), SQLite job store (swap for PostgreSQL later).

## Run (development)

```bash
cd web
python -m venv .venv
. .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -r requirements.txt
uvicorn app.main:app --reload --host 0.0.0.0 --port 8080
```

- **Browser UI:** `http://127.0.0.1:8080/` — bulk job form and recent jobs table
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
| GET | `/api/v1/jobs` | List jobs |
| POST | `/api/v1/jobs/bulk` | Create bulk export job (body: `tenant_ids`, `options`) |
| GET | `/api/v1/jobs/{uuid}` | Job status |
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

### PowerShell stub worker (Linux / webhost)

1. Install **PowerShell 7** (`pwsh`) and ensure it is on `PATH` (or set **`EOA_PWSH_PATH`**).
2. Set **`EOA_USE_PWSH_STUB_WORKER=true`** so each job runs **`web/pwsh/WebBulkJobStub.ps1`** (writes **`web/data/artifacts/<job_id>/summary.json`**). This proves the API → `pwsh` → disk pipeline; replace the script with real **`BulkExportWorker.ps1`** orchestration later.
3. Set **`EOA_REPO_ROOT`** to the repo root if the API working directory is not the repo (default: parent of `web/`).

If **`EOA_USE_PWSH_STUB_WORKER`** is false (default for local dev), the API uses a short **in-process placeholder** only (unless the Python Graph worker is enabled below).

**Worker order:** when both are enabled, **PowerShell runs first** (`web/pwsh/WebBulkJobStub.ps1`); the Python Graph worker runs only if pwsh is disabled or `pwsh` is not on `PATH`.

### Python Graph worker (Linux, optional fallback)

1. Register an Entra app with a **client secret** and grant **application** permissions for the reports you need (e.g. `Organization.Read.All` for `organization`; add `User.Read.All`, `Policy.Read.All`, `Application.Read.All` as you enable more reports). **Admin-consent** the app in the target tenant.
2. Set **`EOA_GRAPH_CLIENT_ID`**, **`EOA_GRAPH_CLIENT_SECRET`**, and **`EOA_USE_PYTHON_GRAPH_WORKER=true`** when you want Graph-based jobs **without** pwsh (or as fallback when pwsh is missing).
3. **Smoke-test a report locally** (writes under `web/data/artifacts/<job-id>/`):

```bash
cd web
# Windows PowerShell: $env:EOA_GRAPH_CLIENT_ID='...'; $env:EOA_GRAPH_CLIENT_SECRET='...'
export EOA_GRAPH_CLIENT_ID=...
export EOA_GRAPH_CLIENT_SECRET=...
python tools/run_graph_report.py <YOUR-TENANT-GUID> organization
# optional: python tools/run_graph_report.py <TENANT-GUID> organization users
```

Exit code **0** means every requested report succeeded; **1** means at least one failed (check `report_*.json` and `worker.log`). **2** means credentials were not set.

### Microsoft sign-in (browser) — tenant without pasting GUIDs + app registrations

This is **separate** from Authentik/API auth: the header **Sign in** still controls access to `/api/v1/jobs` when `EOA_OIDC_ISSUER` is set. The **Microsoft 365** panel uses **MSAL** in the browser to sign in with a work account and call **Microsoft Graph** directly (delegated).

1. In **Entra ID**, register a **single-page application** (public client).  
   - **Redirect URIs** must include the exact origins you use, e.g. `http://127.0.0.1:8080/`, `http://127.0.0.1:8080/app`, and HTTPS equivalents.  
2. **API permissions** (delegated): `User.Read`, `Organization.Read.All`, `Application.ReadWrite.All` — **admin consent** in the tenant for app CRUD.  
3. Set **`EOA_MS_GRAPH_SPA_CLIENT_ID`** (and optionally **`EOA_MS_GRAPH_TENANT`**, default `organizations`).  
4. Open the console: tenant **display name** and **directory (tenant) ID** appear after **Sign in with Microsoft**. Use **Use this tenant for bulk job** to fill the job field, or submit with an empty tenant field — the signed-in tenant ID is used automatically.  
5. **App registrations**: create, rename, or delete apps from the table (Graph calls from the browser).

## Deploy (Linux)

After **`git pull`**, install Python deps in the same venv systemd uses:

```bash
cd web && .venv/bin/pip install -r requirements.txt && sudo systemctl restart eoa-api
```

HTML/CSS/JS responses include **`CDN-Cache-Control`** and asset URLs use **`?v=<api version>`** so each deploy changes filenames. If the public URL still looks stale, **purge Cloudflare cache** for the hostname once (browser hard refresh is not enough when the edge cached HTML).

Examples (adjust paths/users):

- **[`deploy/eoa-api.service.example`](deploy/eoa-api.service.example)** — systemd unit for `uvicorn` on `127.0.0.1:8080`.
- **[`deploy/nginx-snippet.conf.example`](deploy/nginx-snippet.conf.example)** — reverse proxy snippet for nginx in front of the API.
