from datetime import datetime
from typing import Any

from pydantic import BaseModel, Field


class BulkJobCreate(BaseModel):
    """Body for POST /api/v1/jobs/bulk — mirrors BulkTenantExporter options over time."""

    tenant_ids: list[str] = Field(
        default_factory=list,
        description="Entra directory id(s). Web UI sends the signed-in tenant only; integrations may send more.",
    )
    options: dict[str, Any] = Field(
        default_factory=dict,
        description=(
            "Exporter options (snake_case): include_* booleans, days_back, sign_in_logs_days_back, etc. "
            "Web UI builds this from checkboxes; GET /api/v1/export/options-schema for the full schema. "
            "Python Graph worker merges include_* into Graph reports; EXO slices need Windows + interactive EXO (BulkTenantExporter)."
        ),
    )


class JobOut(BaseModel):
    id: str
    status: str
    kind: str
    created_at: datetime
    updated_at: datetime
    created_by_sub: str | None = None
    error_message: str | None = None
    artifact_uri: str | None = None
    artifact_files: list[str] | None = None
    request_payload: dict[str, Any] | None = Field(
        default=None,
        description="Original POST body snapshot (tenant_ids + options) for Run again in the UI.",
    )

    model_config = {"from_attributes": True}


class JobArtifactNamesOut(BaseModel):
    job_id: str
    files: list[str]


class JobListOut(BaseModel):
    jobs: list[JobOut]
    total: int
