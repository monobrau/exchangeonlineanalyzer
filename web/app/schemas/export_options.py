"""Structured options for POST /api/v1/jobs/bulk — mirrors ExportUtils New-SecurityInvestigationReport / BulkTenantExporter."""

from typing import Any

from pydantic import BaseModel, ConfigDict, Field


class WebBulkExportOptions(BaseModel):
    """
    Target shape for job `options` (snake_case JSON). Workers may ignore unknown keys.
    PowerShell uses Include* booleans; map 1:1 to include_* here.
    """

    model_config = ConfigDict(extra="allow")

    investigator_name: str | None = Field(
        default=None,
        description="Investigator label on exported report metadata (desktop default: Security Administrator).",
    )
    company_name: str | None = Field(
        default=None,
        description="Organization label on exported report metadata.",
    )
    days_back: int | None = Field(
        default=None,
        ge=1,
        description="Relative window: analyze last N days when start/end not set (desktop default 10).",
    )
    message_trace_days_back: int | None = Field(
        default=None,
        ge=1,
        description="Message trace lookback (Exchange); may differ from days_back.",
    )
    sign_in_logs_days_back: int | None = Field(
        default=None,
        ge=1,
        description="Sign-in log lookback (Graph / Entra).",
    )
    start_date: str | None = Field(
        default=None,
        description="ISO 8601 start of absolute range (optional; use with end_date).",
    )
    end_date: str | None = Field(
        default=None,
        description="ISO 8601 end of absolute range (optional).",
    )
    selected_users: list[str] = Field(
        default_factory=list,
        description="UPNs to scope user-centric slices (inbox rules, forwarding, etc.).",
    )
    ticket_numbers: list[str] = Field(default_factory=list, description="Ticket IDs for report header/metadata.")
    ticket_content: str | None = Field(default=None, description="Free-text ticket context for metadata.")

    include_message_trace: bool | None = Field(
        default=None,
        description="Exchange message trace — requires EXO on worker host for full parity.",
    )
    include_inbox_rules: bool | None = Field(default=None, description="Per-mailbox inbox rules (EXO / Graph hybrid in PS).")
    include_transport_rules: bool | None = Field(default=None, description="Transport rules (EXO).")
    include_mail_flow_connectors: bool | None = Field(default=None, description="Connectors (EXO).")
    include_mailbox_forwarding: bool | None = Field(default=None, description="Forwarding settings (EXO/Graph).")
    include_audit_logs: bool | None = Field(default=None, description="Directory/Entra audit where applicable.")
    include_sign_in_logs: bool | None = Field(default=None, description="Sign-in logs (Graph).")
    include_intune_devices: bool | None = Field(default=None, description="Intune-managed devices (Graph).")
    include_mfa_coverage: bool | None = Field(default=None, description="MFA / auth methods coverage (Graph).")
    include_conditional_access_policies: bool | None = Field(
        default=None,
        description="Conditional Access policies (Graph identity).",
    )
    include_app_registrations: bool | None = Field(default=None, description="App registrations (Graph).")
    include_share_point_activity: bool | None = Field(default=None, description="SharePoint activity reports.")
    include_one_drive_activity: bool | None = Field(default=None, description="OneDrive activity reports.")
    include_teams_activity: bool | None = Field(default=None, description="Teams activity reports.")
    include_share_point_sharing: bool | None = Field(default=None, description="SharePoint sharing activity.")
    include_anonymous_share_point_sharing: bool | None = Field(
        default=None,
        description="Anonymous SharePoint sharing where collected in desktop exporter.",
    )
    include_share_point_file_sharing_links: bool | None = Field(
        default=None,
        description="File sharing links inventory (Graph/SharePoint APIs).",
    )
    include_share_point_one_drive_file_actions: bool | None = Field(
        default=None,
        description="SharePoint/OneDrive file actions audit slice.",
    )
    include_security_alerts: bool | None = Field(default=None, description="Microsoft 365 Defender / Graph security alerts.")
    include_security_incidents: bool | None = Field(default=None, description="Security incidents (Graph).")
    include_unified_audit_logs: bool | None = Field(default=None, description="Unified audit (EXO/Graph depending on implementation).")
    unified_audit_log_record_types: list[str] | None = Field(
        default=None,
        description="Optional filter for unified audit record types (desktop supports list).",
    )
    include_dlp_violations: bool | None = Field(default=None, description="DLP policy matches / violations.")

    reports: list[str] = Field(
        default_factory=list,
        description=(
            "Python Graph worker: report keys — organization, users, conditional_access, applications, "
            "sign_in_logs, directory_audits, security_alerts, security_incidents, intune_devices, mfa_registration "
            "(aliases: org, ca, signins, …). Booleans include_* also add matching Graph reports when set."
        ),
    )

    @classmethod
    def options_schema(cls) -> dict[str, Any]:
        """JSON Schema for OpenAPI and UI generators."""
        return cls.model_json_schema()

