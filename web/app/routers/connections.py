"""Server-side connection status for Graph app-only and Exchange Online worker (no browser MSAL)."""

from __future__ import annotations

import re
from uuid import UUID

from fastapi import APIRouter, Depends

from app.auth import require_user
from app.config import get_settings

router = APIRouter(prefix="/connections", tags=["connections"])

_GUID_RE = re.compile(r"^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$", re.I)


def _norm_guid(s: str) -> str | None:
    t = (s or "").strip()
    if not t or not _GUID_RE.match(t):
        return None
    try:
        return str(UUID(t))
    except ValueError:
        return None


@router.get("/status")
def connections_status(_: str | None = Depends(require_user)) -> dict:
    """Whether Graph/EXO are configured via env (app-only); optional default tenant for jobs."""
    s = get_settings()
    graph_ok = bool(s.graph_client_id.strip() and s.graph_client_secret.strip())
    exo_skip = s.exo_skip_connect
    exo_ready = bool(
        s.exo_app_id.strip() and s.exo_organization.strip() and s.exo_certificate_thumbprint.strip()
    )
    if exo_skip:
        exo_state = "skipped"
    elif exo_ready:
        exo_state = "ready"
    else:
        exo_state = "not_configured"

    default_tid = _norm_guid(s.job_default_tenant_id or "")

    return {
        "graph_app_configured": graph_ok,
        "use_python_graph_worker": s.use_python_graph_worker,
        "use_pwsh_worker": s.use_pwsh_stub_worker,
        "exo": exo_state,
        "exo_organization": (s.exo_organization or "").strip() or None,
        "job_default_tenant_id": default_tid,
    }
