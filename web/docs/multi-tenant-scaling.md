# Multi-tenant bulk jobs (API / automation only)

The **browser console** is designed for **one signed-in Microsoft directory at a time** (interactive MSAL) and **checkbox** export options — **no** pasted tenant IDs or JSON.

The **`POST /api/v1/jobs/bulk`** API can still accept **`tenant_ids`** arrays for **server-side automation** (e.g. Python Graph worker with app-only credentials). That path is for **integrations**, not the default operator workflow.

For product direction: **Graph** app registration and permission work is **delegated + interactive** in the Microsoft 365 panel; **Exchange Online** follows **`BulkTenantExporter.ps1`** with **interactive EXO** on a Windows workstation.
