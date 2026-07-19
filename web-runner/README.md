# Bulk Web Runner

Browser UI + local API for bulk tenant security investigation exports. Requires a **Windows runner** on the analyst machine.

## Start

```powershell
cd C:\Git\exchangeonlineanalyzer\exchangeonlineanalyzer
.\web-runner\Start-BulkWebRunner.ps1
```

Opens http://127.0.0.1:8765/

Workers run **hidden** by default (no PowerShell console popups). For troubleshooting:

```powershell
.\web-runner\Start-BulkWebRunner.ps1 -ShowWorkers
```

**LAN access** (other PCs on your network can open the UI; auth popups still run on this machine):

```powershell
.\web-runner\Start-BulkWebRunner.ps1 -ListenLan -NoBrowser
```

First run may require two one-time **elevated** PowerShell steps (the script prints exact commands if needed):

1. URL reservation if bind fails: `netsh http add urlacl url=http://+:8765/ user="DOMAIN\user" listen=yes`
2. Firewall allow inbound TCP 8765: `netsh advfirewall firewall add rule name="EOA Bulk Web Runner (TCP 8765)" dir=in action=allow protocol=TCP localport=8765 profile=domain,private,public enable=yes`

Use your PC's real LAN IP (e.g. `http://192.168.1.50:8765/`), not Hyper-V/WSL virtual addresses like `172.19.x.x`.

Per-tenant **Show console** restarts a worker with a visible window.

## Flow

1. **New session** — temp dir, session timeframe, report selections (same shape as WinForms).
2. **Add tenant** — launches `Scripts/BulkExportWorker.ps1` (hidden by default).
3. **Load app registrations** — WCM tenants for Graph auth dropdown.
4. **App registrations (WCM)** — Create / delete Graph app, export / import `.eoa-creds`, clear local WCM.
5. Per tenant: Manage ticket fetch or paste → **Graph Auth** → **Exchange Auth** → **Generate Reports**.
6. **Activity** + **Client N** log tabs poll worker status during auth and generate.

Auth popups appear on **this PC**, not in the browser.

## Main app launcher

The **Bulk Tenant Report Exporter** button in the main app opens the web runner by default. Choose **No** on the prompt to use legacy WinForms `BulkTenantExporter.ps1` for one release.

## Architecture

See [docs/BulkWebRunnerArchitecture.md](../docs/BulkWebRunnerArchitecture.md).

## Security integrations (Liongard · Huntress · SentinelOne)

After **Fetch from Manage**, the tenant card **Security integrations** panel resolves the client in Liongard and suggests Huntress / SentinelOne pulls from ticket alert type and product entitlements.

### Credentials (EOA Settings JSON — server-side only)

| Setting | Purpose |
|---------|---------|
| `LiongardInstance`, `LiongardAccessKey`, `LiongardAccessSecret` | Liongard ROAR API (`X-ROAR-API-KEY`) |
| `HuntressApiKey`, `HuntressApiSecret` | Huntress REST API (Basic auth) |
| `SentinelOneConnectWiseInstanceId`, `SentinelOneConnectWiseApiToken` | ConnectWise MSP S1 console (full read) |
| `SentinelOneBarracudaInstanceId`, `SentinelOneBarracudaApiToken` | Barracuda XDR S1 site (read-only Viewer token) |

### API routes

- `GET /api/liongard/status`, `POST /api/liongard/resolve-client`, `POST /api/liongard/export-context`
- `GET /api/huntress/status`, `GET /api/huntress/organizations`, `POST /api/huntress/preview`, `POST /api/huntress/export`
- `GET /api/sentinelone/status`, `POST /api/sentinelone/resolve-site`, `POST /api/sentinelone/preview`, `POST /api/sentinelone/export`

Exports land in the tenant's report folder (or a new timestamp folder under `Documents\ExchangeOnlineAnalyzer\SecurityInvestigation\{Company}\`).

### Barracuda XDR S1 — API assumption and fallback

Barracuda XDR does not expose separate endpoint telemetry APIs; EOA uses the **SentinelOne Management API** with a **Viewer** service user scoped to the client's site.

**Assumption:** You can create (or Barracuda provisions) a read-only S1 API token for Barracuda-managed sites.

**If Barracuda token is unavailable:**

- The UI shows: *Barracuda S1 API unavailable — use Barracuda portal and ticket IOCs. No ConnectWise profile fallback.*
- Parsed ticket IOCs from `Filter-TicketContent` remain available; EOA does **not** silently fall back to the ConnectWise S1 profile (wrong tenant scope risk).

**Validation steps:**

1. In the S1 console tied to Barracuda-managed endpoints, create a service user with **Viewer** role and site scope.
2. Set `SentinelOneBarracudaInstanceId` (hostname prefix or full URL) and `SentinelOneBarracudaApiToken` in settings.
3. `GET /api/sentinelone/status` should show `barracuda_xdr.configured: true`.
4. On a Barracuda XDR ticket, **Resolve client** should auto-select the Barracuda profile and site ID from Liongard when mapped.

Restart the web runner after changing modules or settings (`Ctrl+C`, then re-run `Start-BulkWebRunner.ps1`).
