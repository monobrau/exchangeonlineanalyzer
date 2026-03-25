"""GUI-editable runtime settings persisted to web/data/eoa_gui.env (overrides web/.env)."""

from __future__ import annotations

from collections import OrderedDict
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Literal

from app.config import GUI_ENV_FILE, Settings, get_settings, reload_settings

Kind = Literal["bool", "int", "str", "secret", "path"]


@dataclass(frozen=True, slots=True)
class RuntimeField:
    env_key: str
    attr: str
    kind: Kind
    section: str
    label: str
    description: str = ""
    restart_required: bool = False


# Order defines UI and write order. Maps to Settings fields (EOA_ prefix applied by pydantic).
RUNTIME_FIELDS: tuple[RuntimeField, ...] = (
    RuntimeField("EOA_DEBUG", "debug", "bool", "general", "Debug mode", "More verbose API logging."),
    RuntimeField(
        "EOA_CORS_ORIGINS",
        "cors_origins",
        "str",
        "general",
        "CORS allowed origins",
        'Comma-separated origins, or * (default). Requires API restart to take effect.',
        restart_required=True,
    ),
    RuntimeField("EOA_OIDC_ISSUER", "oidc_issuer", "str", "oidc", "OIDC issuer URL", "Authentik / IdP issuer base URL."),
    RuntimeField("EOA_OIDC_CLIENT_ID", "oidc_client_id", "str", "oidc", "OIDC client ID", "OAuth2 application client ID."),
    RuntimeField(
        "EOA_OIDC_CLIENT_SECRET",
        "oidc_client_secret",
        "secret",
        "oidc",
        "OIDC client secret",
        "Only for confidential clients / HS256 tokens.",
    ),
    RuntimeField(
        "EOA_OIDC_AUDIENCE",
        "oidc_audience",
        "str",
        "oidc",
        "OIDC audience",
        "JWT aud claim (often same as client ID).",
    ),
    RuntimeField(
        "EOA_OIDC_AUDIENCES",
        "oidc_audiences",
        "str",
        "oidc",
        "OIDC audiences (comma-separated)",
        "Optional; multiple acceptable aud values.",
    ),
    RuntimeField(
        "EOA_OIDC_REDIRECT_URI",
        "oidc_redirect_uri",
        "str",
        "oidc",
        "OIDC redirect URI",
        "Must match IdP app registration exactly.",
    ),
    RuntimeField("EOA_OIDC_SCOPE", "oidc_scope", "str", "oidc", "OIDC scope", "Default: openid profile email."),
    RuntimeField(
        "EOA_SESSION_SECRET",
        "session_secret",
        "secret",
        "oidc",
        "Session / cookie signing secret",
        "Long random string. Requires API restart after change.",
        restart_required=True,
    ),
    RuntimeField(
        "EOA_REPO_ROOT",
        "repo_root",
        "path",
        "workers",
        "Repository root",
        "Repo root for pwsh worker and artifacts (default: parent of web/).",
    ),
    RuntimeField("EOA_PWSH_PATH", "pwsh_path", "str", "workers", "pwsh executable", 'Default: pwsh on PATH.'),
    RuntimeField(
        "EOA_PWSH_WORKER_SCRIPT",
        "pwsh_worker_script",
        "str",
        "workers",
        "PowerShell worker script",
        "File name under web/pwsh/.",
    ),
    RuntimeField(
        "EOA_USE_PWSH_STUB_WORKER",
        "use_pwsh_stub_worker",
        "bool",
        "workers",
        "Run PowerShell stub worker",
        "When pwsh exists, runs web/pwsh/<script> per job.",
    ),
    RuntimeField(
        "EOA_PYTHON_GRAPH_BEFORE_PWSH",
        "python_graph_before_pwsh",
        "bool",
        "workers",
        "Run Python Graph before pwsh",
        "When both workers on, Graph first; stub writes pwsh_summary.json.",
    ),
    RuntimeField(
        "EOA_USE_PYTHON_GRAPH_WORKER",
        "use_python_graph_worker",
        "bool",
        "workers",
        "Use Python Graph worker",
        "App-only Graph on Linux when pwsh is off or after pwsh (see order flag).",
    ),
    RuntimeField(
        "EOA_PWSH_NONINTERACTIVE",
        "pwsh_noninteractive",
        "bool",
        "workers",
        "pwsh -NonInteractive",
        "When true (default), worker cannot prompt (use EXO cert + Graph secret). Set false only for device-code testing.",
    ),
    RuntimeField(
        "EOA_EXO_APP_ID",
        "exo_app_id",
        "str",
        "exo_pwsh",
        "EXO app-only — client ID",
        "Entra app registered for Exchange PowerShell app-only (certificate). Used by WebExoLinuxRunner.ps1.",
    ),
    RuntimeField(
        "EOA_EXO_ORGANIZATION",
        "exo_organization",
        "str",
        "exo_pwsh",
        "EXO organization",
        "Initial domain, e.g. contoso.onmicrosoft.com (Connect-ExchangeOnline -Organization).",
    ),
    RuntimeField(
        "EOA_EXO_CERT_THUMBPRINT",
        "exo_certificate_thumbprint",
        "str",
        "exo_pwsh",
        "EXO certificate thumbprint",
        "Certificate installed on the webhost for the EXO app (not the Graph secret).",
    ),
    RuntimeField(
        "EOA_EXO_SKIP_CONNECT",
        "exo_skip_connect",
        "bool",
        "exo_pwsh",
        "Skip Exchange Online connection",
        "If true, only Graph runs (New-SecurityInvestigationReport); EXO cmdlets see no session.",
    ),
    RuntimeField(
        "EOA_GRAPH_CLIENT_ID",
        "graph_client_id",
        "str",
        "graph_server",
        "Graph app client ID (server)",
        "Entra app for client-credentials Graph worker.",
    ),
    RuntimeField(
        "EOA_GRAPH_CLIENT_SECRET",
        "graph_client_secret",
        "secret",
        "graph_server",
        "Graph app client secret (server)",
        "Application permission app; keep secret.",
    ),
    RuntimeField(
        "EOA_GRAPH_MAX_TENANTS_PER_JOB",
        "graph_max_tenants_per_job",
        "int",
        "graph_server",
        "Max tenants per Graph job",
        "Cap for tenant_ids per job (1–10000).",
    ),
    RuntimeField(
        "EOA_JOB_DEFAULT_TENANT_ID",
        "job_default_tenant_id",
        "str",
        "graph_server",
        "Default tenant ID for jobs (GUID)",
        "When the queue is empty, jobs use this Entra directory id. Set in Settings instead of browser sign-in.",
    ),
    RuntimeField(
        "EOA_MS_GRAPH_SPA_USE_GRAPH_APP_ID",
        "ms_graph_spa_use_graph_app_id",
        "bool",
        "microsoft_spa",
        "Browser MSAL: reuse Graph worker app ID",
        "Default on: when SPA client ID is empty, use EOA_GRAPH_CLIENT_ID so M365 sign-in works without a separate SPA env var if the same Entra app has a Single-page application platform with this site redirect (e.g. https://your-host/app). Set false only if Graph is a confidential-only app.",
    ),
    RuntimeField(
        "EOA_MS_GRAPH_SPA_CLIENT_ID",
        "ms_graph_spa_client_id",
        "str",
        "microsoft_spa",
        "M365 browser sign-in — Entra SPA client ID",
        "Enables MSAL in the UI when set (or use bundled ID in code). Users can still paste a client ID in the browser if this is empty.",
    ),
    RuntimeField(
        "EOA_MS_GRAPH_TENANT",
        "ms_graph_tenant",
        "str",
        "microsoft_spa",
        "M365 sign-in — authority tenant",
        "login.microsoftonline.com/<this>: organizations (work accounts), common, consumers, or a directory GUID.",
    ),
    RuntimeField(
        "EOA_MS_GRAPH_DELEGATED_SCOPES",
        "ms_graph_delegated_scopes",
        "str",
        "microsoft_spa",
        "M365 Graph — delegated scopes (browser)",
        "Comma-separated Microsoft Graph delegated scopes for MSAL (e.g. User.Read, Organization.Read.All, Application.ReadWrite.All). Leave empty for the app default list.",
    ),
)

ALLOWED_ENV_KEYS = frozenset(f.env_key for f in RUNTIME_FIELDS)
SECRET_ENV_KEYS = frozenset(f.env_key for f in RUNTIME_FIELDS if f.kind == "secret")
FIELD_BY_ENV = {f.env_key: f for f in RUNTIME_FIELDS}


def _quote_env_value(val: str) -> str:
    if val == "":
        return '""'
    if any(c in val for c in ' #"\n\r\t\\'):
        esc = val.replace("\\", "\\\\").replace('"', '\\"')
        return f'"{esc}"'
    return val


def parse_env_file(path: Path) -> OrderedDict[str, str]:
    if not path.is_file():
        return OrderedDict()
    raw = path.read_text(encoding="utf-8")
    out: OrderedDict[str, str] = OrderedDict()
    for line in raw.splitlines():
        s = line.strip()
        if not s or s.startswith("#"):
            continue
        if "=" not in s:
            continue
        k, _, v = s.partition("=")
        key = k.strip()
        val = v.strip()
        if len(val) >= 2 and val[0] == val[-1] and val[0] in "'\"":
            val = val[1:-1]
        out[key] = val
    return out


def format_env_file(data: OrderedDict[str, str]) -> str:
    if not data:
        return ""
    lines = [f"{k}={_quote_env_value(v)}" for k, v in data.items()]
    return "\n".join(lines) + "\n"


def atomic_write_text(path: Path, text: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(path.suffix + ".tmp")
    tmp.write_text(text, encoding="utf-8")
    tmp.replace(path)


def _raw_setting_value(s: Settings, f: RuntimeField) -> Any:
    return getattr(s, f.attr)


def build_runtime_payload() -> dict[str, Any]:
    s = get_settings()
    items: list[dict[str, Any]] = []
    for f in RUNTIME_FIELDS:
        raw = _raw_setting_value(s, f)
        entry: dict[str, Any] = {
            "env_key": f.env_key,
            "attr": f.attr,
            "kind": f.kind,
            "section": f.section,
            "label": f.label,
            "description": f.description,
            "restart_required": f.restart_required,
        }
        if f.kind == "secret":
            text = (raw or "").strip() if isinstance(raw, str) else ""
            entry["value"] = None
            entry["has_value"] = bool(text)
        elif f.kind == "bool":
            entry["value"] = bool(raw)
        elif f.kind == "int":
            entry["value"] = int(raw)
        elif f.kind == "path":
            try:
                entry["value"] = str(Path(raw).resolve()) if raw else ""
            except (OSError, TypeError, ValueError):
                entry["value"] = str(raw) if raw is not None else ""
        else:
            entry["value"] = str(raw) if raw is not None else ""
        items.append(entry)
    return {
        "items": items,
        "gui_env_file": str(GUI_ENV_FILE.resolve()),
        "note": (
            "Values shown are effective settings (web/.env plus overrides in web/data/eoa_gui.env). "
            "Saving writes only to eoa_gui.env. Restart the API process after changing CORS or session secret."
        ),
    }


def _coerce_patch_value(f: RuntimeField, val: Any) -> str:
    if f.kind == "bool":
        if isinstance(val, bool):
            return "true" if val else "false"
        if isinstance(val, str):
            return "true" if val.strip().lower() in ("1", "true", "yes", "on") else "false"
        raise ValueError(f"{f.env_key}: expected boolean")
    if f.kind == "int":
        n = int(val)
        if f.attr == "graph_max_tenants_per_job" and not (1 <= n <= 10000):
            raise ValueError(f"{f.env_key}: must be between 1 and 10000")
        return str(n)
    if f.kind in ("str", "secret", "path"):
        return str(val).strip() if val is not None else ""
    raise ValueError(f"{f.env_key}: unsupported kind {f.kind}")


def apply_runtime_patch(patch: dict[str, Any]) -> dict[str, Any]:
    """
    Merge patch into GUI env file. Secret keys: empty string removes override from gui file.
    str/path: empty string removes gui override. Returns updated_keys and restart_recommended.
    """
    unknown = [k for k in patch if k not in ALLOWED_ENV_KEYS]
    if unknown:
        raise ValueError(f"Unknown keys: {', '.join(sorted(unknown))}")
    if not patch:
        return {"updated_keys": [], "restart_recommended": False}

    current = parse_env_file(GUI_ENV_FILE)
    restart_touch = False
    updated: list[str] = []

    for env_key, raw_val in patch.items():
        f = FIELD_BY_ENV[env_key]
        if raw_val is None:
            continue
        if f.kind == "secret":
            sval = str(raw_val).strip()
            if sval == "":
                if env_key in current:
                    del current[env_key]
                    updated.append(env_key)
                continue
            current[env_key] = sval
            updated.append(env_key)
            if f.restart_required:
                restart_touch = True
            continue
        if f.kind in ("str", "path") and str(raw_val).strip() == "":
            if env_key in current:
                del current[env_key]
                updated.append(env_key)
            continue
        coerced = _coerce_patch_value(f, raw_val)
        current[env_key] = coerced
        updated.append(env_key)
        if f.restart_required:
            restart_touch = True

    # Stable write order: definition order first, then any stray keys
    ordered = OrderedDict()
    for f in RUNTIME_FIELDS:
        if f.env_key in current:
            ordered[f.env_key] = current[f.env_key]
    for k, v in current.items():
        if k not in ordered:
            ordered[k] = v

    text = format_env_file(ordered)
    if text.strip():
        atomic_write_text(GUI_ENV_FILE, text)
    elif GUI_ENV_FILE.is_file():
        GUI_ENV_FILE.unlink()

    reload_settings()

    rr = restart_touch or any(
        FIELD_BY_ENV[k].restart_required for k in updated if k in FIELD_BY_ENV
    )
    return {"updated_keys": updated, "restart_recommended": rr}
