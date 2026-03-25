# Linux EXO runner (`WebExoLinuxRunner.ps1`)

The web API can run a **local `pwsh` process** on the same host as `uvicorn` and execute the same **`New-SecurityInvestigationReport`** pipeline as the desktop app (`Modules/ExportUtils.psm1` in the repo root), so bulk jobs can produce **Exchange Online + Microsoft Graph** artifacts without a Windows desktop session.

## Security model

- **Job JSON** (`POST /api/v1/jobs/bulk`) may only contain **`tenant_ids`** (validated GUIDs) and **`options`** (report flags). **No secrets** are accepted in the payload.
- **Authentication** uses **environment variables only** (or `web/data/eoa_gui.env` via **Settings**):
  - **Graph (application permissions):** `EOA_GRAPH_CLIENT_ID`, `EOA_GRAPH_CLIENT_SECRET` — same app used by the Python Graph worker, or a dedicated app with Graph permissions for the collectors you need.
  - **Exchange Online (app-only certificate):** `EOA_EXO_APP_ID`, `EOA_EXO_ORGANIZATION` (e.g. `contoso.onmicrosoft.com`), `EOA_EXO_CERT_THUMBPRINT` — certificate must be installed where `pwsh` runs (see Microsoft’s **app-only authentication** docs for Exchange Online PowerShell).
- **`EOA_REPO_ROOT`** must point at this repository so the worker can import `Modules/ExportUtils.psm1`.

Optional: **`EOA_EXO_SKIP_CONNECT=true`** runs **Graph-only** (no `Connect-ExchangeOnline`); EXO-specific slices will be empty.

## Enable on the webhost

1. Install **PowerShell 7** and prerequisites for **Exchange Online Management on Linux** ([Microsoft Learn](https://learn.microsoft.com/powershell/exchange/exchange-online-powershell-v2?view=exchange-ps)).
2. Install **Microsoft Graph PowerShell modules** required by `ExportUtils` (e.g. `Microsoft.Graph.Authentication` and others your selected reports need).
3. Set **`EOA_USE_PWSH_STUB_WORKER=true`** and point the worker script at the EXO runner:

   ```bash
   EOA_PWSH_WORKER_SCRIPT=WebExoLinuxRunner.ps1
   ```

4. Configure **EXO** and **Graph** env vars (or use the in-app **Settings** panel for the same keys).

## Operational notes

- **First tenant only:** if `tenant_ids` has multiple entries, the runner uses the **first** GUID for token and EXO organization scoping; narrow jobs to one tenant per job until multi-org is designed.
- **`EOA_PWSH_NONINTERACTIVE`:** default **`true`** (subprocess uses `pwsh -NonInteractive`). Use **certificate + client secret** auth. Set **`false`** only if you intentionally use **device code** and a human completes sign-in during the job (not typical for production).
- **Timeouts:** the API subprocess timeout is **600s** today; large audits may need a higher timeout in `job_runner.py` later.

## Parity with `WebBulkJobStub.ps1`

The runner still writes **`ReportSelections.json`** in the same shape as the stub for desktop compatibility. **`summary.json`** identifies **`workerBackend`: `pwsh-exo-linux`**.
