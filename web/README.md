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

If **`EOA_USE_PWSH_STUB_WORKER`** is false (default for local dev), the API uses a short **in-process placeholder** only.

## Deploy (Linux)

Examples (adjust paths/users):

- **[`deploy/eoa-api.service.example`](deploy/eoa-api.service.example)** — systemd unit for `uvicorn` on `127.0.0.1:8080`.
- **[`deploy/nginx-snippet.conf.example`](deploy/nginx-snippet.conf.example)** — reverse proxy snippet for nginx in front of the API.
