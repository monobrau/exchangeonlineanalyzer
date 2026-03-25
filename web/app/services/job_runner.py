"""Job execution: PowerShell (WebBulkJobStub.ps1) first, then optional Python Graph worker, else placeholder."""

from __future__ import annotations

import json
import logging
import shutil
import subprocess
import time
from datetime import datetime, timezone
from pathlib import Path

from sqlalchemy.orm import Session

from app.config import get_settings
from app.db import SessionLocal
from app.models import Job, JobStatus

logger = logging.getLogger(__name__)


def _utcnow() -> datetime:
    return datetime.now(timezone.utc)


def _artifact_dir(job_id: str) -> Path:
    root = get_settings().repo_root
    d = root / "web" / "data" / "artifacts" / job_id
    d.mkdir(parents=True, exist_ok=True)
    return d


def _payload_path(job_id: str) -> Path:
    root = get_settings().repo_root
    d = root / "web" / "data" / "job_payloads"
    d.mkdir(parents=True, exist_ok=True)
    return d / f"{job_id}.json"


def _pwsh_executable_exists(pwsh: str) -> bool:
    return shutil.which(pwsh) is not None or Path(pwsh).is_file()


def _write_worker_log(out_dir: Path, text: str) -> None:
    try:
        (out_dir / "worker.log").write_text(text, encoding="utf-8")
    except OSError:
        logger.warning("Could not write worker.log under %s", out_dir)


def _run_pwsh_stub(job_id: str, job: Job) -> tuple[bool, str, str | None]:
    """Returns (ok, log_text, artifact_uri)."""
    settings = get_settings()
    script = settings.repo_root / "web" / "pwsh" / "WebBulkJobStub.ps1"
    if not script.is_file():
        return False, f"Missing worker script: {script}", None

    pwsh = settings.pwsh_path
    if not _pwsh_executable_exists(pwsh):
        return False, f"Executable not found: {pwsh}", None

    out_dir = _artifact_dir(job_id)
    payload = _payload_path(job_id)
    body = job.request_payload or {}
    payload.write_text(json.dumps(body, indent=2), encoding="utf-8")

    cmd = [
        pwsh,
        "-NoProfile",
        "-NonInteractive",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        str(script),
        "-PayloadJsonPath",
        str(payload),
        "-JobId",
        job_id,
        "-OutputDir",
        str(out_dir),
    ]
    logger.info("Running %s", " ".join(cmd))
    try:
        proc = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=600,
            cwd=str(settings.repo_root),
        )
    except subprocess.TimeoutExpired:
        return False, "pwsh subprocess timeout (600s)", None

    log = (proc.stdout or "") + ("\n--- stderr ---\n" + proc.stderr if proc.stderr else "")
    if proc.returncode != 0:
        _write_worker_log(out_dir, log)
        return False, log or f"exit {proc.returncode}", None

    _write_worker_log(out_dir, log)
    summary = out_dir / "summary.json"
    uri = f"file://{summary.resolve()}" if summary.is_file() else f"file://{out_dir.resolve()}/"
    return True, log, uri


def _run_placeholder_only(job_id: str, *, note: str = "") -> None:
    time.sleep(1.5)
    out = _artifact_dir(job_id)
    p = out / "placeholder.txt"
    lines = [
        "Placeholder worker (Python): no worker ran.",
        "Set EOA_USE_PWSH_STUB_WORKER=true and install PowerShell 7 (pwsh) on PATH to run web/pwsh/WebBulkJobStub.ps1 for each job.",
        "Optional: EOA_USE_PYTHON_GRAPH_WORKER=true plus EOA_GRAPH_* for the Python Graph worker (runs only if pwsh is not used or unavailable).",
    ]
    if note:
        lines.insert(0, note)
    p.write_text("\n".join(lines) + "\n", encoding="utf-8")
    logger.info("Placeholder finished %s", job_id)


def run_job(job_id: str) -> None:
    """Mark job running, execute pwsh stub or placeholder, then terminal status."""
    db: Session = SessionLocal()
    try:
        job = db.get(Job, job_id)
        if not job:
            logger.warning("Job %s not found for worker", job_id)
            return
        job.status = JobStatus.running.value
        job.updated_at = _utcnow()
        db.commit()

        settings = get_settings()
        ok = False
        log_tail = ""
        artifact_uri: str | None = None

        # 1) PowerShell worker first when enabled (typical production: WebBulkJobStub.ps1 or future real exporter).
        if settings.use_pwsh_stub_worker and _pwsh_executable_exists(settings.pwsh_path):
            ok, log_tail, artifact_uri = _run_pwsh_stub(job_id, job)
            if not ok:
                job = db.get(Job, job_id)
                if job:
                    job.status = JobStatus.failed.value
                    job.updated_at = _utcnow()
                    job.error_message = (log_tail or "pwsh failed")[:8000]
                    job.artifact_uri = artifact_uri
                    db.commit()
                logger.error("Job %s pwsh failed: %s", job_id, log_tail[:500])
                return

            job = db.get(Job, job_id)
            if job:
                job.status = JobStatus.succeeded.value
                job.updated_at = _utcnow()
                job.artifact_uri = artifact_uri
                if len(log_tail) < 4000:
                    job.error_message = None
                db.commit()
            logger.info("Job %s pwsh ok", job_id)
            return

        if settings.use_pwsh_stub_worker and not _pwsh_executable_exists(settings.pwsh_path):
            logger.warning(
                "EOA_USE_PWSH_STUB_WORKER is true but %s not found; trying python graph or placeholder",
                settings.pwsh_path,
            )

        # 2) Python Graph worker (optional) when pwsh not used or pwsh missing.
        if settings.use_python_graph_worker:
            _gcid = (settings.graph_client_id or "").strip()
            _gsec = (settings.graph_client_secret or "").strip()
            if _gcid and _gsec:
                from app.services.graph_worker import run_graph_bulk_job

                ok, log_tail, artifact_uri = run_graph_bulk_job(job_id, job)
                if not ok:
                    job = db.get(Job, job_id)
                    if job:
                        job.status = JobStatus.failed.value
                        job.updated_at = _utcnow()
                        job.error_message = (log_tail or "graph worker failed")[:8000]
                        job.artifact_uri = artifact_uri
                        db.commit()
                    logger.error("Job %s python-graph failed: %s", job_id, log_tail[:500])
                    return

                job = db.get(Job, job_id)
                if job:
                    job.status = JobStatus.succeeded.value
                    job.updated_at = _utcnow()
                    job.artifact_uri = artifact_uri
                    if len(log_tail) < 4000:
                        job.error_message = None
                    db.commit()
                logger.info("Job %s python-graph ok", job_id)
                return

        if settings.use_pwsh_stub_worker and not _pwsh_executable_exists(settings.pwsh_path):
            _run_placeholder_only(
                job_id,
                note=f"Skipped pwsh worker: executable not found ({settings.pwsh_path}).",
            )
        else:
            _run_placeholder_only(job_id)
        job = db.get(Job, job_id)
        if not job:
            return
        job.status = JobStatus.succeeded.value
        job.updated_at = _utcnow()
        job.artifact_uri = f"file://{_artifact_dir(job_id).resolve()}/placeholder.txt"
        db.commit()
        logger.info("Placeholder job %s finished", job_id)
    except Exception as e:
        logger.exception("Job %s failed", job_id)
        try:
            job = db.get(Job, job_id)
            if job:
                job.status = JobStatus.failed.value
                job.updated_at = _utcnow()
                job.error_message = str(e)[:8000]
                db.commit()
        except Exception:
            db.rollback()
    finally:
        db.close()


# Backwards-compatible name for router
run_placeholder_job = run_job
