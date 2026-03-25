"""Read/update GUI-managed env overrides (web/data/eoa_gui.env)."""

from __future__ import annotations

from fastapi import APIRouter, Depends, HTTPException
from pydantic import BaseModel, Field

from app.auth import require_user
from app.services.runtime_env_settings import apply_runtime_patch, build_runtime_payload

router = APIRouter(prefix="/settings", tags=["settings"])


class RuntimeEnvPatchBody(BaseModel):
    patch: dict[str, object] = Field(default_factory=dict)


@router.get("/runtime-env")
def get_runtime_env(_sub: str | None = Depends(require_user)) -> dict[str, object]:
    """Effective settings for the browser settings panel (secrets masked)."""
    return build_runtime_payload()


@router.put("/runtime-env")
def put_runtime_env(
    body: RuntimeEnvPatchBody,
    _sub: str | None = Depends(require_user),
) -> dict[str, object]:
    """Merge into web/data/eoa_gui.env and reload process settings cache."""
    try:
        return apply_runtime_patch(dict(body.patch))
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e)) from e
