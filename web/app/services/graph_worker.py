"""Cross-platform bulk job worker: Microsoft Graph via MSAL (client credentials) + httpx.

No Windows, PowerShell, or Exchange Online remote session required. Requires an Entra app
registration with application permissions, consented per tenant (see options.reports).

Typical application permissions by report:
  organization — Organization.Read.All
  users — User.Read.All
  conditional_access — Policy.Read.All (or Policy.Read.ConditionalAccess)
  applications — Application.Read.All
"""

from __future__ import annotations

import json
import logging
import platform
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

import httpx
import msal

from app.config import Settings, get_settings
from app.models import Job

logger = logging.getLogger(__name__)

GRAPH_SCOPE = ["https://graph.microsoft.com/.default"]
GRAPH_BASE = "https://graph.microsoft.com/v1.0"

# Canonical report keys (options.reports uses strings; aliases map here).
REPORT_KEYS = frozenset({"organization", "users", "conditional_access", "applications"})
REPORT_ALIASES: dict[str, str] = {
    "organization": "organization",
    "org": "organization",
    "tenant": "organization",
    "users": "users",
    "user": "users",
    "directory": "users",
    "conditional_access": "conditional_access",
    "ca": "conditional_access",
    "capolicies": "conditional_access",
    "applications": "applications",
    "apps": "applications",
    "app_registrations": "applications",
    "apps_registrations": "applications",
}
# "rules" (inbox rules) is Exchange/Graph mail — not implemented in this worker yet; skipped with warning.

USER_SELECT = "id,displayName,userPrincipalName,accountEnabled,assignedLicenses,userType"
APPLICATION_SELECT = "id,appId,displayName,createdDateTime,signInAudience,publisherDomain"


def graph_worker_configured(settings: Settings) -> bool:
    cid = (settings.graph_client_id or "").strip()
    csec = (settings.graph_client_secret or "").strip()
    return bool(cid and csec)


def parse_requested_reports(options: dict[str, Any] | None) -> list[str]:
    """Normalize options.reports to a list of canonical report keys. Default: organization only."""
    if not options:
        return ["organization"]
    raw = options.get("reports")
    if raw is None:
        return ["organization"]
    if isinstance(raw, str):
        raw = [raw]
    if not isinstance(raw, list) or len(raw) == 0:
        return ["organization"]
    out: list[str] = []
    for x in raw:
        key = str(x).strip().lower()
        canon = REPORT_ALIASES.get(key)
        if canon is None:
            if key in REPORT_KEYS:
                canon = key
            else:
                logger.warning("Unknown report key skipped: %s", x)
                continue
        if canon not in out:
            out.append(canon)
    return out if out else ["organization"]


def _artifact_dir(job_id: str) -> Path:
    root = get_settings().repo_root
    d = root / "web" / "data" / "artifacts" / job_id
    d.mkdir(parents=True, exist_ok=True)
    return d


def _write_worker_log(out_dir: Path, text: str) -> None:
    try:
        (out_dir / "worker.log").write_text(text, encoding="utf-8")
    except OSError:
        logger.warning("Could not write worker.log under %s", out_dir)


def acquire_graph_token(tenant_id: str, client_id: str, client_secret: str) -> tuple[str | None, str | None]:
    """Client-credentials token for the given tenant. Returns (access_token, error_message)."""
    authority = f"https://login.microsoftonline.com/{tenant_id.strip()}"
    app = msal.ConfidentialClientApplication(
        client_id.strip(),
        authority=authority,
        client_credential=client_secret.strip(),
    )
    result = app.acquire_token_for_client(scopes=GRAPH_SCOPE)
    if result and result.get("access_token"):
        return result["access_token"], None
    err = result.get("error_description") or result.get("error") or str(result) if result else "unknown"
    return None, err


def _graph_get_json(url: str, token: str, timeout: float = 120.0) -> tuple[dict[str, Any] | None, int, str | None]:
    headers = {"Authorization": f"Bearer {token}"}
    try:
        with httpx.Client(timeout=timeout) as client:
            r = client.get(url, headers=headers)
            text = r.text
            if r.status_code >= 400:
                return None, r.status_code, text[:4000]
            try:
                return r.json(), r.status_code, None
            except json.JSONDecodeError:
                return None, r.status_code, text[:4000]
    except httpx.HTTPError as e:
        return None, 0, str(e)


def graph_get_all_pages(
    first_url: str,
    token: str,
    *,
    max_items: int = 100_000,
    max_pages: int = 500,
) -> tuple[list[Any], str | None]:
    """Follow @odata.nextLink. Returns (items, error_message)."""
    items: list[Any] = []
    url: str | None = first_url
    page = 0
    while url and page < max_pages:
        data, status, err = _graph_get_json(url, token)
        if err:
            return items, err
        if not data:
            return items, f"empty response status={status}"
        chunk = data.get("value")
        if chunk is None:
            # Single-object endpoints sometimes return no "value"
            if page == 0 and isinstance(data, dict) and "id" in data:
                return [data], None
            return items, err or f"unexpected shape status={status}"
        items.extend(chunk)
        if len(items) >= max_items:
            break
        url = data.get("@odata.nextLink")
        page += 1
        if not url:
            break
    return items, None


def _write_report_json(out_dir: Path, name: str, payload: dict[str, Any]) -> None:
    (out_dir / f"report_{name}.json").write_text(json.dumps(payload, indent=2), encoding="utf-8")


def _collect_organization(
    out_dir: Path, token: str, tid: str
) -> tuple[bool, str | None]:
    url = f"{GRAPH_BASE}/organization"
    data, status, err = _graph_get_json(url, token)
    if err:
        _write_report_json(
            out_dir,
            "organization",
            {"ok": False, "tenantId": tid, "httpStatus": status, "error": err},
        )
        return False, err
    values = (data or {}).get("value") or []
    first = values[0] if values else {}
    payload = {
        "ok": True,
        "tenantId": tid,
        "displayName": first.get("displayName"),
        "id": first.get("id"),
        "verifiedDomains": [
            (d or {}).get("name") for d in (first.get("verifiedDomains") or [])[:20]
        ],
    }
    _write_report_json(out_dir, "organization", payload)
    return True, None


def _collect_users(out_dir: Path, token: str, tid: str) -> tuple[bool, str | None]:
    sel = USER_SELECT.replace(",", "%2C")
    first = f"{GRAPH_BASE}/users?$select={sel}&$top=999"
    users, err = graph_get_all_pages(first, token, max_items=100_000)
    if err:
        _write_report_json(
            out_dir,
            "users",
            {"ok": False, "tenantId": tid, "error": err, "users": []},
        )
        return False, err
    _write_report_json(
        out_dir,
        "users",
        {
            "ok": True,
            "tenantId": tid,
            "count": len(users),
            "users": users,
        },
    )
    return True, None


def _collect_conditional_access(out_dir: Path, token: str, tid: str) -> tuple[bool, str | None]:
    first = f"{GRAPH_BASE}/identity/conditionalAccess/policies"
    policies, err = graph_get_all_pages(first, token, max_items=10_000)
    if err:
        _write_report_json(
            out_dir,
            "conditional_access",
            {"ok": False, "tenantId": tid, "error": err, "policies": []},
        )
        return False, err
    _write_report_json(
        out_dir,
        "conditional_access",
        {
            "ok": True,
            "tenantId": tid,
            "count": len(policies),
            "policies": policies,
        },
    )
    return True, None


def _collect_applications(out_dir: Path, token: str, tid: str) -> tuple[bool, str | None]:
    sel = APPLICATION_SELECT.replace(",", "%2C")
    first = f"{GRAPH_BASE}/applications?$select={sel}&$top=999"
    apps, err = graph_get_all_pages(first, token, max_items=50_000)
    if err:
        _write_report_json(
            out_dir,
            "applications",
            {"ok": False, "tenantId": tid, "error": err, "applications": []},
        )
        return False, err
    _write_report_json(
        out_dir,
        "applications",
        {
            "ok": True,
            "tenantId": tid,
            "count": len(apps),
            "applications": apps,
        },
    )
    return True, None


def run_graph_bulk_job(job_id: str, job: Job) -> tuple[bool, str, str | None]:
    """Returns (ok, log_text, artifact_uri). Writes summary.json, graph.json, report_*.json."""
    settings = get_settings()
    out_dir = _artifact_dir(job_id)
    body = job.request_payload or {}
    tenant_ids = body.get("tenant_ids") or []
    options = body.get("options") if isinstance(body.get("options"), dict) else {}
    reports = parse_requested_reports(options)

    if not tenant_ids:
        msg = "No tenant_ids in job payload; cannot acquire tenant-scoped token."
        _write_worker_log(out_dir, msg)
        summary = {
            "workerVersion": "3",
            "workerBackend": "python-graph",
            "jobId": job_id,
            "ok": False,
            "error": msg,
            "reportsRequested": reports,
            "at": datetime.now(timezone.utc).isoformat(),
        }
        (out_dir / "summary.json").write_text(json.dumps(summary, indent=2), encoding="utf-8")
        return False, msg, f"file://{out_dir.resolve()}/"

    tid = str(tenant_ids[0]).strip()
    client_id = settings.graph_client_id.strip()
    client_secret = settings.graph_client_secret.strip()

    log_lines: list[str] = [
        f"python-graph worker job={job_id} tenant={tid}",
        f"reports={reports}",
        f"platform={platform.system()} python={sys.version.split()[0]}",
    ]

    token, terr = acquire_graph_token(tid, client_id, client_secret)
    if not token:
        log_lines.append(f"token_error: {terr}")
        _write_worker_log(out_dir, "\n".join(log_lines))
        summary = {
            "workerVersion": "3",
            "workerBackend": "python-graph",
            "jobId": job_id,
            "ok": False,
            "tenantId": tid,
            "error": terr,
            "reportsRequested": reports,
            "at": datetime.now(timezone.utc).isoformat(),
        }
        (out_dir / "summary.json").write_text(json.dumps(summary, indent=2), encoding="utf-8")
        return False, "\n".join(log_lines), f"file://{out_dir.resolve()}/"

    reports_failed: dict[str, str] = {}
    reports_ok: list[str] = []

    collectors: dict[str, Any] = {
        "organization": _collect_organization,
        "users": _collect_users,
        "conditional_access": _collect_conditional_access,
        "applications": _collect_applications,
    }

    for name in reports:
        fn = collectors.get(name)
        if not fn:
            log_lines.append(f"skip: no collector for {name!r}")
            continue
        log_lines.append(f"collect: {name}...")
        ok, err = fn(out_dir, token, tid)
        if ok:
            reports_ok.append(name)
            log_lines.append(f"ok: {name}")
        else:
            reports_failed[name] = err or "unknown error"
            log_lines.append(f"failed: {name}: {err}")

    graph_artifact = {
        "tenantId": tid,
        "tenantIdsInPayload": tenant_ids[:20],
        "options": options,
        "reportsRequested": reports,
        "reportsCompleted": reports_ok,
        "reportsFailed": reports_failed,
    }
    (out_dir / "graph.json").write_text(json.dumps(graph_artifact, indent=2), encoding="utf-8")

    summary = {
        "workerVersion": "3",
        "workerBackend": "python-graph",
        "jobId": job_id,
        "tenantCount": len(tenant_ids),
        "tenantIdsSample": tenant_ids[:5],
        "options": options,
        "reportsRequested": reports,
        "reportsCompleted": reports_ok,
        "reportsFailed": reports_failed,
        "repoRootEnv": str(settings.repo_root) if settings.repo_root else None,
        "message": (
            "Python Graph worker: MSAL app-only token; per-report JSON under report_*.json."
        ),
        "python": sys.version.split()[0],
        "os": platform.platform(),
        "ok": len(reports_failed) == 0,
        "at": datetime.now(timezone.utc).isoformat(),
    }
    (out_dir / "summary.json").write_text(json.dumps(summary, indent=2), encoding="utf-8")

    log_text = "\n".join(log_lines)
    _write_worker_log(out_dir, log_text)
    ok_job = len(reports_failed) == 0
    summary_path = out_dir / "summary.json"
    uri = f"file://{summary_path.resolve()}"
    return ok_job, log_text, uri
