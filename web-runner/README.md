# Bulk Web Runner

Browser UI + local API for bulk tenant security investigation exports. Requires a **Windows runner** on the analyst machine.

## Start

```powershell
cd C:\Git\exchangeonlineanalyzer\exchangeonlineanalyzer
.\web-runner\Start-BulkWebRunner.ps1
```

Opens http://127.0.0.1:8765/

If you previously used `-ListenLan`, a `http://+:8765/` URL reservation already exists. Plain start now auto-uses that reservation (avoids HTTP.sys **503 Service Unavailable** from a localhost-only bind). Starting again on the same port stops the existing EOA web runner first (localhost shutdown when it is healthy, then force-stop leftovers).

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
4. **App registrations (WCM)** — Create / **Update Graph App scopes** / delete Graph app, export / import `.eoa-creds`, clear local WCM.
5. Per tenant: Manage ticket fetch or paste → **Graph Auth** → **Exchange Auth** → **Generate Reports**. Graph/EXO stay connected so you can **Generate another report pack** on the same tenant (each run is a new timestamped folder).
6. **Containment** (per tenant card): **Validate users**, then work the BEC playbook top to bottom — lock the account (revoke / password reset / block / hold+audit), identity footholds (MFA, OAuth consents, Entra devices, ActiveSync, Intune), mailbox persistence (inbox rules, forwarding, delegates, folder ACL, auto-reply, junk lists, rights on other mailboxes, Restricted Users), tenant-wide persistence (transport rules, connectors, org auto-forward, journaling, apps, secrets/owners, roles/groups/RBAC), then restore (unblock). **Save containment zips** writes `Containment_<user>.zip` (with `actions.csv` for password reset, revoke, block, and other account changes) plus folder-level `Remediation.csv`. **Clear user pulls** drops per-user list results so you can pull the next user; the account-change log, tenant-wide lists, and saved zips stay. Status and list buttons run immediately; writes use a confirm popup. Actions run in that tenant’s worker — the browser never holds tokens. Passwords are never written to the log.
7. **Activity** + **Client N** log tabs poll worker status during auth and generate.
8. After reports exist: **Analyze reports** (findings) and **Curate logs** (include/exclude facet values → `Curated_*` CSV set beside the pack; originals untouched).

### Graph write scopes (containment)

Interactive **Graph Auth** now requests `User.RevokeSessions.All`, `User.EnableDisableAccount.All`, `UserAuthenticationMethod.ReadWrite.All`, `Application.ReadWrite.All`, `User-PasswordProfile.ReadWrite.All`, `DelegatedPermissionGrant.ReadWrite.All`, `RoleManagement.ReadWrite.Directory`, `GroupMember.ReadWrite.All`, `DeviceManagementManagedDevices.ReadWrite.All`, `DeviceManagementManagedDevices.PrivilegedOperations.All`, and `Directory.AccessAsUser.All` (device delete / password reset / consent revoke while signed in as you) in addition to the existing read scopes. Sign in again after a worker restart so consent can include them.

WCM / app-only tokens gain those writes after **Update Graph App scopes** (App registrations, or the button under the Containment permission hint). That patches the existing River Run Security Investigator app and grants missing admin consent; the client secret and WCM entry stay the same. Then run **Graph Auth** again. Use **Create Graph App** only when the app is missing. App-only device delete uses `Device.ReadWrite.All`. App-only password reset uses `User-PasswordProfile.ReadWrite.All` and also needs the **User Administrator** directory role on that app. Until then, Graph containment buttons stay disabled or fail with a missing-permission message. Restricted Users and inbox rules use the Exchange session and do not need those Graph app roles.

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
