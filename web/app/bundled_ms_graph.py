"""Bundled Microsoft Graph SPA (public) client ID — no secret.

OAuth requires *some* Entra app registration. When set here, the M365 browser panel works
without EOA_MS_GRAPH_SPA_CLIENT_ID. The same ID must be registered as a multi-tenant SPA
with redirect URIs for each deployment origin (see README).

Override per deployment: set EOA_MS_GRAPH_SPA_CLIENT_ID (wins over bundled).
"""

from __future__ import annotations

# Paste the Application (client) ID after registering the official EOA multi-tenant SPA app in Entra.
# Leave empty to rely on env or browser localStorage (eoa_ms_graph_spa_client_id).
BUNDLED_MS_GRAPH_SPA_CLIENT_ID = ""
