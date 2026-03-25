"""Microsoft Graph delegated (browser MSAL) — resolved client ID and scopes."""

from __future__ import annotations

from app.bundled_ms_graph import BUNDLED_MS_GRAPH_SPA_CLIENT_ID
from app.config import Settings

# Same list as auth_oidc msal-config; exposed when enabled=false so the browser can use localStorage client ID.
DELEGATED_GRAPH_SCOPES = [
    "User.Read",
    "Organization.Read.All",
    "Application.ReadWrite.All",
]


def resolve_ms_graph_spa_client_id(settings: Settings) -> str:
    """EOA_MS_GRAPH_SPA_CLIENT_ID wins; optional reuse of Graph worker app id; else bundled (may be empty)."""
    explicit = (settings.ms_graph_spa_client_id or "").strip()
    if explicit:
        return explicit
    if settings.ms_graph_spa_use_graph_app_id:
        g = (settings.graph_client_id or "").strip()
        if g:
            return g
    return BUNDLED_MS_GRAPH_SPA_CLIENT_ID.strip()


def resolve_delegated_graph_scopes(settings: Settings) -> list[str]:
    """EOA_MS_GRAPH_DELEGATED_SCOPES (comma/semicolon) or default DELEGATED_GRAPH_SCOPES."""
    raw = (settings.ms_graph_delegated_scopes or "").strip()
    if not raw:
        return list(DELEGATED_GRAPH_SCOPES)
    parts = [p.strip() for p in raw.replace(";", ",").split(",")]
    out = [p for p in parts if p]
    return out if out else list(DELEGATED_GRAPH_SCOPES)
