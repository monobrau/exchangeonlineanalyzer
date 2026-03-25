from datetime import datetime
from typing import Any

from pydantic import BaseModel, Field


class BulkJobCreate(BaseModel):
    """Body for POST /api/v1/jobs/bulk — mirrors BulkTenantExporter options over time."""

    tenant_ids: list[str] = Field(default_factory=list, description="Entra tenant id(s) to process")
    options: dict[str, Any] = Field(
        default_factory=dict,
        description=(
            "Exporter options (snake_case). GET /api/v1/export/options-schema returns JSON Schema. "
            "Python Graph worker: key 'reports' — organization, users, conditional_access, applications (aliases org, ca, apps). "
            "include_* booleans mirror BulkTenantExporter Include*; EXO-heavy slices need a Windows + ExchangeOnlineManagement worker."
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

    model_config = {"from_attributes": True}


class JobArtifactNamesOut(BaseModel):
    job_id: str
    files: list[str]


class JobListOut(BaseModel):
    jobs: list[JobOut]
    total: int
