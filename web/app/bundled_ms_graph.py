"""Bundled Microsoft Graph SPA (public) client ID — no secret.

OAuth requires *some* Entra app registration. When set here, the M365 browser panel works
without EOA_MS_GRAPH_SPA_CLIENT_ID. Register as an SPA with redirect URIs for each deployment origin (see README).
Single-tenant vs multitenant is chosen in Entra (app registration account types), not in this constant.

By default the API also reuses EOA_GRAPH_CLIENT_ID for browser MSAL when this is empty
(see EOA_MS_GRAPH_SPA_USE_GRAPH_APP_ID). Override per deployment: set EOA_MS_GRAPH_SPA_CLIENT_ID (wins over bundled).
"""

from __future__ import annotations

# Paste the Application (client) ID after registering the SPA app in Entra (see web/tools/Register-EoaMsalSpaApp.ps1).
# Leave empty to rely on env or browser localStorage (eoa_ms_graph_spa_client_id).
BUNDLED_MS_GRAPH_SPA_CLIENT_ID = ""
