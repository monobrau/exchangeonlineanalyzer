from functools import lru_cache
from pathlib import Path

from pydantic import field_validator
from pydantic_settings import BaseSettings, SettingsConfigDict

# Always load web/.env next to the app package (not CWD-dependent when started by systemd).
_WEB_DIR = Path(__file__).resolve().parent.parent
_ENV_FILE = _WEB_DIR / ".env"


class Settings(BaseSettings):
    model_config = SettingsConfigDict(
        env_prefix="EOA_",
        env_file=_ENV_FILE,
        env_file_encoding="utf-8",
        extra="ignore",
    )

    app_name: str = "ExchangeOnlineAnalyzer API"
    debug: bool = False

    # SQLite database file
    database_url: str = ""

    # Optional: Authentik / OIDC — when set, /api/v1/* requires Bearer JWT
    oidc_issuer: str = ""
    oidc_audience: str = ""
    # Comma-separated if multiple acceptable audiences
    oidc_audiences: str = ""
    # Browser login (Authorization Code + PKCE) — must match Authentik provider settings
    oidc_client_id: str = ""
    oidc_client_secret: str = ""
    oidc_redirect_uri: str = ""
    oidc_scope: str = "openid profile email"

    # Starlette session cookie signing (required for OIDC PKCE state). Use a long random string in production.
    session_secret: str = ""

    # Repo root (for worker: pwsh scripts under repo). Override on server (EOA_REPO_ROOT).
    repo_root: Path = Path(__file__).resolve().parent.parent.parent

    # PowerShell 7+ executable (Linux/macOS/Windows)
    pwsh_path: str = "pwsh"

    # Script under web/pwsh/ (default WebBulkJobStub.ps1). Override for a custom worker entrypoint.
    pwsh_worker_script: str = "WebBulkJobStub.ps1"

    # If true, run web/pwsh/<pwsh_worker_script> when pwsh is available. Default false for dev/CI; set true on webhost.
    use_pwsh_stub_worker: bool = False

    # Linux-friendly worker: MSAL + Graph REST. Runs only if pwsh worker did not run (see job_runner order).
    use_python_graph_worker: bool = False
    graph_client_id: str = ""
    graph_client_secret: str = ""
    # Cap tenant_ids processed per job (Python Graph worker). Override with options.max_tenants (cannot exceed this cap).
    graph_max_tenants_per_job: int = 300

    # Microsoft Graph SPA (browser MSAL): Entra app registration — "Single-page application", public client.
    # Used for sign-in with Microsoft (tenant context + delegated Graph for app registration CRUD).
    # Register redirect URIs: e.g. http://127.0.0.1:8080/ and http://127.0.0.1:8080/app (and HTTPS equivalents).
    ms_graph_spa_client_id: str = ""
    # Authority tenant: organizations | common | consumers | or a directory (tenant) GUID
    ms_graph_tenant: str = "organizations"

    # CORS: "*" or comma-separated origins (e.g. https://app.example.com,http://localhost:5173)
    cors_origins: str = "*"

    @field_validator(
        "oidc_issuer",
        "oidc_audience",
        "oidc_audiences",
        "oidc_client_id",
        "oidc_client_secret",
        "oidc_redirect_uri",
        mode="before",
    )
    @classmethod
    def _strip_oidc_strings(cls, v: object) -> object:
        if isinstance(v, str):
            return v.strip()
        return v

    @field_validator("ms_graph_spa_client_id", "ms_graph_tenant", mode="before")
    @classmethod
    def _strip_ms_graph_strings(cls, v: object) -> object:
        if isinstance(v, str):
            return v.strip()
        return v

    def model_post_init(self, __context: object) -> None:
        if not self.database_url:
            data_dir = Path(__file__).resolve().parent.parent / "data"
            data_dir.mkdir(parents=True, exist_ok=True)
            self.database_url = f"sqlite:///{data_dir / 'eoa_jobs.db'}"


@lru_cache
def get_settings() -> Settings:
    return Settings()
