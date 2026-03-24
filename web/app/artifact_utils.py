"""Resolve files under web/data/artifacts/<job_id>/ (path traversal safe)."""

import re
from pathlib import Path
from uuid import UUID

from fastapi import HTTPException

_SAFE = re.compile(r"^[a-zA-Z0-9._\-]{1,160}$")


def artifacts_root(repo_root: Path) -> Path:
    return (repo_root / "web" / "data" / "artifacts").resolve()


def job_artifact_file(repo_root: Path, job_id: UUID, filename: str) -> Path:
    if not _SAFE.match(filename):
        raise HTTPException(status_code=400, detail="Invalid filename")
    root = artifacts_root(repo_root)
    job_dir = (root / str(job_id)).resolve()
    try:
        job_dir.relative_to(root)
    except ValueError as e:
        raise HTTPException(status_code=400, detail="Invalid job path") from e
    path = (job_dir / filename).resolve()
    try:
        path.relative_to(job_dir)
    except ValueError as e:
        raise HTTPException(status_code=400, detail="Invalid artifact path") from e
    if not path.is_file():
        raise HTTPException(status_code=404, detail="Artifact file not found")
    return path


def list_artifact_filenames(repo_root: Path, job_id: str) -> list[str]:
    """Non-hidden file names under web/data/artifacts/<job_id>/ (safe names only)."""
    root = artifacts_root(repo_root)
    job_dir = (root / job_id).resolve()
    try:
        job_dir.relative_to(root)
    except ValueError:
        return []
    if not job_dir.is_dir():
        return []
    out: list[str] = []
    for p in sorted(job_dir.iterdir()):
        if p.is_file() and not p.name.startswith(".") and _SAFE.match(p.name):
            out.append(p.name)
    return out
