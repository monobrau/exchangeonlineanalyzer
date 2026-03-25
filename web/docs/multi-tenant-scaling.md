# Many tenants (e.g. 300+) — no per-tenant interactive login

Interactive **browser** sign-in (MSAL delegated) does **not** scale to hundreds of tenants: you cannot manually sign in once per customer directory.

For bulk operations across many Microsoft Entra tenants, use **one** backend app registration with **application (client) credentials** and **admin consent** in each customer tenant.

## Recommended architecture

1. **Register one multi-tenant app** (or a single-tenant app if you only ever run in your own tenant) in **your** Entra tenant.
2. Grant **application permissions** for the Graph reports you need (see `web/README.md` Python Graph worker section). Examples: `Organization.Read.All`, `User.Read.All`, `Policy.Read.All`, `Application.Read.All`, `AuditLog.Read.All`, etc.
3. For **each customer tenant** that you must access, an administrator must **consent** that app **once** (same pattern as Microsoft 365 Lighthouse, CSP partners, or any multi-tenant SaaS):
   - Use the **admin consent** URL for your app and the customer’s tenant ID, **or**
   - **Partner** scenarios: GDAP, CSP, Delegated Admin — your product onboarding should capture **directory (tenant) ID** and store it (database, CSV, CRM), not rely on a human copying IDs from the portal for each run.

4. **Run jobs** with **`tenant_ids`** containing many GUIDs (comma-separated or JSON array). The **Python Graph worker** acquires an **app-only** token **per tenant** and writes artifacts under `tenants/<tenant-guid>/` when more than one tenant is in the job.

5. **Cap**: `EOA_GRAPH_MAX_TENANTS_PER_JOB` (default **300**). Optional job `options.max_tenants` (cannot exceed the server cap).

## What not to rely on at scale

- **Delegated** Graph in the browser for each tenant — only suitable for **admin/dev** tools, not 300 tenants.
- **Manually** pasting tenant IDs for every run — automate ingestion (API, file upload, partner sync).

## API shape

```json
{
  "tenant_ids": [
    "aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa",
    "bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb"
  ],
  "options": {
    "reports": ["organization", "users"],
    "max_tenants": 300
  }
}
```

Server env: **`EOA_GRAPH_CLIENT_ID`**, **`EOA_GRAPH_CLIENT_SECRET`**, **`EOA_USE_PYTHON_GRAPH_WORKER=true`**.

## Artifact layout (multi-tenant)

When `len(tenant_ids) > 1`:

- `summary.json` — `perTenant` results
- `graph.json` — `perTenant` map
- `tenants/<tenant-guid>/report_*.json` — per-directory exports

Single-tenant jobs keep `report_*.json` at the artifact root (backward compatible).
