import uuid
from datetime import datetime, timezone
from uuid import UUID

from fastapi import APIRouter, BackgroundTasks, Depends, HTTPException, Query
from fastapi.responses import FileResponse
from sqlalchemy import func, select
from sqlalchemy.orm import Session

from app.artifact_utils import job_artifact_file, list_artifact_filenames
from app.auth import require_user
from app.config import get_settings
from app.db import get_db
from app.models import Job, JobStatus
from app.schemas.job import BulkJobCreate, JobArtifactNamesOut, JobListOut, JobOut
from app.services.job_runner import run_placeholder_job

router = APIRouter(prefix="/jobs", tags=["jobs"])


def _utcnow() -> datetime:
    return datetime.now(timezone.utc)


def _job_to_out(job: Job) -> JobOut:
    out = JobOut.model_validate(job)
    if job.status != JobStatus.succeeded.value:
        return out.model_copy(update={"artifact_files": None})
    files = list_artifact_filenames(get_settings().repo_root, job.id)
    return out.model_copy(update={"artifact_files": files})


@router.get("", response_model=JobListOut)
def list_jobs(
    db: Session = Depends(get_db),
    _: str | None = Depends(require_user),
    limit: int = Query(50, ge=1, le=200),
    offset: int = Query(0, ge=0),
) -> JobListOut:
    total = db.scalar(select(func.count()).select_from(Job)) or 0
    rows = db.scalars(select(Job).order_by(Job.created_at.desc()).offset(offset).limit(limit)).all()
    return JobListOut(jobs=[_job_to_out(r) for r in rows], total=int(total))


@router.post("/bulk", response_model=JobOut, status_code=201)
def create_bulk_job(
    body: BulkJobCreate,
    background_tasks: BackgroundTasks,
    db: Session = Depends(get_db),
    created_by_sub: str | None = Depends(require_user),
) -> JobOut:
    job = Job(
        id=str(uuid.uuid4()),
        status=JobStatus.queued.value,
        kind="bulk_export",
        created_at=_utcnow(),
        updated_at=_utcnow(),
        created_by_sub=created_by_sub,
        request_payload={
            "tenant_ids": body.tenant_ids,
            "options": body.options,
        },
    )
    db.add(job)
    db.commit()
    db.refresh(job)
    background_tasks.add_task(run_placeholder_job, job.id)
    return _job_to_out(job)


@router.get("/{job_id}/artifacts", response_model=JobArtifactNamesOut)
def list_job_artifacts(
    job_id: UUID,
    db: Session = Depends(get_db),
    _: str | None = Depends(require_user),
) -> JobArtifactNamesOut:
    """List file names in web/data/artifacts/<job_id>/ (succeeded jobs only)."""
    job = db.get(Job, str(job_id))
    if not job:
        raise HTTPException(status_code=404, detail="Job not found")
    if job.status != JobStatus.succeeded.value:
        raise HTTPException(status_code=409, detail="Job has no artifacts yet (not succeeded)")
    files = list_artifact_filenames(get_settings().repo_root, str(job_id))
    return JobArtifactNamesOut(job_id=str(job_id), files=files)


@router.get("/{job_id}/artifact")
def download_job_artifact(
    job_id: UUID,
    file: str = Query(
        "summary.json",
        description="File name inside the job artifact directory (e.g. summary.json, placeholder.txt)",
    ),
    db: Session = Depends(get_db),
    _: str | None = Depends(require_user),
) -> FileResponse:
    """Download a file from web/data/artifacts/<job_id>/ (after job has written artifacts)."""
    job = db.get(Job, str(job_id))
    if not job:
        raise HTTPException(status_code=404, detail="Job not found")
    if job.status != JobStatus.succeeded.value:
        raise HTTPException(status_code=409, detail="Job has no artifacts yet (not succeeded)")

    path = job_artifact_file(get_settings().repo_root, job_id, file)
    return FileResponse(
        path,
        filename=path.name,
        media_type="application/octet-stream",
    )


@router.get("/{job_id}", response_model=JobOut)
def get_job_by_id(
    job_id: UUID,
    db: Session = Depends(get_db),
    _: str | None = Depends(require_user),
) -> JobOut:
    job = db.get(Job, str(job_id))
    if not job:
        raise HTTPException(status_code=404, detail="Job not found")
    return _job_to_out(job)
