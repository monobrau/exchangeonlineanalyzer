#!/usr/bin/env python3
"""Live test of the Python Graph worker (MSAL + Graph REST).

From repo: cd web

  set EOA_GRAPH_CLIENT_ID=your-app-id
  set EOA_GRAPH_CLIENT_SECRET=your-secret
  python tools/run_graph_report.py <tenant-guid> [report ...]

Reports default to "organization" if none given. Examples:

  python tools/run_graph_report.py 00000000-0000-0000-0000-000000000000
  python tools/run_graph_report.py 00000000-0000-0000-0000-000000000000 organization users
  python tools/run_graph_report.py 00000000-0000-0000-0000-000000000000 sign_in_logs security_alerts
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

_WEB_ROOT = Path(__file__).resolve().parent.parent
if str(_WEB_ROOT) not in sys.path:
    sys.path.insert(0, str(_WEB_ROOT))


def main() -> int:
    parser = argparse.ArgumentParser(description="Test-run graph_worker reports against a tenant.")
    parser.add_argument("tenant_id", help="Entra tenant ID (directory ID)")
    parser.add_argument(
        "reports",
        nargs="*",
        default=["organization"],
        help="Report keys (default: organization). E.g. users conditional_access applications",
    )
    parser.add_argument(
        "--job-id",
        default="cli-test-run",
        help="Artifact folder name under web/data/artifacts/",
    )
    args = parser.parse_args()

    from app.config import get_settings
    from app.services.graph_worker import graph_worker_configured, run_graph_bulk_job

    get_settings.cache_clear()
    settings = get_settings()

    if not graph_worker_configured(settings):
        print(
            "Missing Graph app credentials. Set:\n"
            "  EOA_GRAPH_CLIENT_ID\n"
            "  EOA_GRAPH_CLIENT_SECRET\n"
            f"(optional: load {_WEB_ROOT / '.env'})",
            file=sys.stderr,
        )
        return 2

    class _Job:
        request_payload = {
            "tenant_ids": [args.tenant_id.strip()],
            "options": {"reports": list(args.reports)},
        }

    print(f"Job id: {args.job_id}")
    print(f"Tenant: {args.tenant_id.strip()}")
    print(f"Reports: {args.reports}")
    print("Running...")

    ok, log, uri = run_graph_bulk_job(args.job_id, _Job())  # type: ignore[arg-type]
    print(log)
    print(f"artifact: {uri}")
    print(f"ok: {ok}")

    out = _WEB_ROOT / "data" / "artifacts" / args.job_id
    if out.is_dir():
        print(f"files: {sorted(p.name for p in out.iterdir())}")

    get_settings.cache_clear()
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
