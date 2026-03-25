# Web bulk export — parity with BulkTenantExporter / ExportUtils

This document maps the **Windows** bulk flow (`BulkTenantExporter.ps1` → `BulkExportWorker.ps1` → `ExportUtils.psm1` → `New-SecurityInvestigationReport`) to the **web** stack and defines a phased path to **feature parity** where technically possible.

## What the PowerShell stack does today

1. **Per-tenant context** — `Connect-ExchangeOnline` + `Connect-MgGraph` (interactive or token), optional **selected users** (UPNs), **days back** or **absolute date range**.
2. **`New-SecurityInvestigationReport`** builds a single structured report object with toggles (`Include*`) for:
   - **Exchange Online (remote PS)**: message trace, inbox rules, transport rules, mail flow connectors, mailbox forwarding, unified audit (EXO).
   - **Microsoft Graph** (parallel/sequential): directory, audit/sign-in, CA policies, app registrations, SharePoint/OneDrive/Teams activity, security alerts/incidents, MFA coverage, Intune devices, DLP, sharing links, etc.
3. **Output** — JSON + optional files under a tenant-scoped folder; **BulkExportWorker** orchestrates worker processes and **report selections** JSON.

Full parity on **Linux-only** workers is **not** possible for **Exchange-only cmdlets** without either: (a) a **Windows** or **Exchange Online–capable** execution host, or (b) replacing those slices with **Graph / REST / export APIs** where Microsoft exposes them (often partial).

## Feature classification

| Area | Examples (Include* / params) | Web today | Full parity path |
|------|------------------------------|-----------|------------------|
| **Graph — directory / identity** | users, CA policies, app registrations, MFA | Python `graph_worker` reports or MSAL browser | Extend `graph_worker` + permissions |
| **Graph — sign-in / audit** | audit logs, sign-in logs | Same | Graph audit query APIs; paging + filters |
| **Graph — security** | alerts, incidents | Same | Security Graph endpoints |
| **Graph — workload activity** | SharePoint, OneDrive, Teams | Same | Reports / Graph (app permissions) |
| **Exchange Online (remote PS)** | message trace, transport rules, connectors, inbox rules (EXO path), unified audit via EXO | **Not** on Linux pwsh alone | **pwsh on Windows** + `ExchangeOnlineManagement`, or Microsoft’s REST where available (limited) |
| **Orchestration** | `DaysBack`, `StartDate`/`EndDate`, `SelectedUsers`, tickets | Partial (`options` JSON) | Same fields in job payload; worker respects them |

## Target web architecture

```
Browser (OIDC + optional MSAL)
    → POST /api/v1/jobs/bulk { tenant_ids, options }
    → SQLite job row
    → Background worker:
        A) **pwsh** → `web/pwsh/WebBulkExport.ps1` (evolution of stub) → calls **BulkExportWorker.ps1** on **Windows** host with repo modules, OR
        B) **Python** → `graph_worker` for Graph-only slices on **Linux**, OR
        C) **Hybrid** — queue EXO slice to Windows worker, Graph slice on Linux (future).
```

**Single artifact root** per job: `web/data/artifacts/<job_id>/` with `summary.json`, `report_*.json`, `worker.log` — same idea as desktop output folders.

## Phased roadmap

### Phase 1 (now — contract & docs)
- Document **all** desktop toggles in **`GET /api/v1/export/options-schema`** (see `export_options.py`).
- Clients send **`options`** using **`snake_case`** keys aligned with this schema (e.g. `include_message_trace`, `days_back`).
- **`WebBulkJobStub.ps1`** writes **`ReportSelections.json`** (same shape as the desktop exporter) plus **`summary.json`**.
- **Python Graph worker** implements additional **`reports`**: sign-in logs, directory audits, security alerts/incidents, Intune devices, MFA registration report; **`include_*`** flags are merged into the report list automatically.

### Phase 2 — PowerShell worker = real bulk export (Windows host)
- Replace stub with a script that:
  - Reads `tenant_ids[0]` and **`options`** from payload JSON.
  - Writes **`ReportSelections`** JSON compatible with **`BulkExportWorker.ps1`** (or calls **`New-SecurityInvestigationReport`** with equivalent parameters).
  - Requires **Windows**, **ExchangeOnlineManagement**, **Microsoft.Graph.*`, repo **`Modules`** on **`EOA_REPO_ROOT`**.
- **Linux** servers: keep stub or Python graph only; document “full EXO parity requires Windows worker.”

### Phase 3 — Python graph parity
- For each **Graph-backed** `Include*` flag, add a **report** implementation in **`graph_worker.py`** (or split modules) with required **application permissions** and admin consent documentation.

### Phase 4 — UI
- Replace free-form JSON with **checkboxes** / **date range** / **user search** in `static/` that POST the structured **`options`** object.

## Payload shape (target)

```json
{
  "tenant_ids": ["<guid>"],
  "options": {
    "days_back": 10,
    "message_trace_days_back": 10,
    "sign_in_logs_days_back": 7,
    "start_date": null,
    "end_date": null,
    "selected_users": ["user@contoso.com"],
    "include_message_trace": true,
    "include_inbox_rules": true,
    "include_transport_rules": true,
    "include_audit_logs": true,
    "reports": ["organization", "users"]
  }
}
```

Use **`include_*`** booleans for parity with **`Include*`** in PowerShell; use **`reports`** for the **Python graph worker** short names where applicable.

## Summary

- **“Web version with all the same features”** = **shared job contract** + **Windows pwsh worker** for EXO + **Graph workers** (Python or pwsh) for Graph.
- **One** API (`/api/v1/jobs/bulk`) can feed **multiple** worker implementations; parity is **progressive**, not a single runtime.
