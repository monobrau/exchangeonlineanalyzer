# Bulk Web Runner (Phase A)

Browser UI + local API for bulk tenant export. Requires a **Windows runner** on the analyst machine (Option B).

## Start

```powershell
cd C:\Git\exchangeonlineanalyzer\exchangeonlineanalyzer
.\web-runner\Start-BulkWebRunner.ps1
```

Opens http://127.0.0.1:8765/

## Flow

1. **New session** — creates temp dir + report selections (same shape as WinForms bulk exporter).
2. **Add tenant** — launches `Scripts/BulkExportWorker.ps1` in a visible PowerShell window.
3. **Load app registrations** — lists WCM tenants (`GraphAppCredential.psm1`).
4. Per tenant: **Graph Auth** → **Exchange Auth** (same commands as `BulkTenantExporter.ps1`).

Auth popups appear on **this PC**, not in the browser.

## Architecture

See [docs/BulkWebRunnerArchitecture.md](../docs/BulkWebRunnerArchitecture.md).

## Not yet in web UI

- Full report checkbox matrix
- Create / delete Graph App buttons (still use `BulkTenantExporter.ps1` or CLI)
- Generate Reports + download artifacts
- User validation / filter UI

These reuse existing PowerShell; wire them through `/api/tenants/{n}/command` when ready.
