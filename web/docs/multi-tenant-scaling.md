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

## Partner Center, CSP, and “partner apps”

**Partner Center** ([partner.microsoft.com](https://partner.microsoft.com/dashboard/home)) is the **partner program / dashboard** (CSP, NCE, customer of record, etc.). It is **not** a drop-in replacement for an Entra app registration, but it changes **how** you discover tenants and **how** consent works.

### Two different “API surfaces”

| Surface | Role |
|--------|------|
| **Partner Center REST APIs** | List customers, subscriptions, invoices, relationships — **not** Microsoft Graph. Use these to **discover customer tenant IDs** and drive your job queue. See [Partner Center developer documentation](https://learn.microsoft.com/partner-center/develop). |
| **Microsoft Graph** | Directory, security, Intune, etc. Your **EOA** worker uses Graph with **client credentials** per **tenant ID** (see above). |

You can combine them: **Partner Center** (or your CRM) → **tenant IDs** → **POST /api/v1/jobs/bulk** with those IDs.

### CSP partner–managed applications (preconsent)

If you are in the **Cloud Solution Provider (CSP)** program and build a **partner-managed** app as Microsoft describes, customers can be **preconsented** for that app in certain **legacy** flows. Microsoft is moving from broad **DAP** to **GDAP** (granular delegated admin privileges). With **GDAP**, customers explicitly approve **least-privileged**, **time-bound** access; partners use Graph APIs such as [delegated admin relationships](https://learn.microsoft.com/graph/api/resources/delegatedadminrelationships-api-overview) and Partner Center automation.

**Important:** Microsoft’s article [Call Microsoft Graph from a Cloud Solution Provider application](https://learn.microsoft.com/graph/auth-cloudsolutionprovider) applies **only** to CSP developers and describes constraints (e.g. regional partner tenants, propagation delay after customer creation, which Graph workloads are in scope). Read it before assuming **every** Graph permission EOA uses is available in a given CSP scenario.

### GDAP vs “just” a multi-tenant app

- **Standalone multi-tenant app** + **admin consent** per customer tenant: works for **any** customer who can consent (typical ISV).
- **CSP / GDAP**: Use when you are the customer’s **partner of record** and you follow Microsoft’s **secure application model** and GDAP relationship APIs — **not** the same onboarding as a random ISV app.

### Practical takeaway for EOA

1. **Not a CSP?** Use one multi-tenant (or per-customer) Entra app, **application permissions**, **admin consent** per tenant, and automate **`tenant_ids`** — as in the sections above.
2. **CSP partner?** Add **Partner Center APIs** to **enumerate customers** and tenant IDs; align your Entra app with Microsoft’s **CSP + Graph** guidance and **GDAP** requirements. The **Python Graph worker** stays the same shape: **app-only token per `tenant_id`** in the job payload.

For official detail: [Introduction to GDAP](https://learn.microsoft.com/partner-center/gdap-introduction), [GDAP and secure application model](https://learn.microsoft.com/partner-center/developer/gdap-and-secure-application-model#comparison-of-gdap-with-dap), [CSP Graph authentication](https://learn.microsoft.com/graph/auth-cloudsolutionprovider).
