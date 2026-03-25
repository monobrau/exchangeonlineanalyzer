"""Metadata for web ↔ desktop bulk export parity (no secrets)."""

from typing import Any

from fastapi import APIRouter

from app.schemas.export_options import WebBulkExportOptions

router = APIRouter(prefix="/export", tags=["export"])


@router.get("/options-schema")
def export_options_schema() -> dict[str, Any]:
    """
    JSON Schema for `options` on POST /api/v1/jobs/bulk (snake_case, mirrors BulkTenantExporter / ExportUtils).
    Extra keys are allowed on the wire; workers document which they honor.
    """
    return {
        "schema_version": 1,
        "json_schema": WebBulkExportOptions.options_schema(),
        "parity_doc": "web/docs/bulk-export-parity.md",
    }
