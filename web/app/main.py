from contextlib import asynccontextmanager
from pathlib import Path

from fastapi import Depends, FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from sqlalchemy import text
from starlette.middleware.sessions import SessionMiddleware

from app.auth import require_user
from app.config import get_settings
from app.db import engine, init_db
from app.routers import auth_oidc, jobs

STATIC_DIR = Path(__file__).resolve().parent.parent / "static"


@asynccontextmanager
async def lifespan(_: FastAPI):
    init_db()
    yield


settings = get_settings()
app = FastAPI(
    title=settings.app_name,
    description="Bulk export jobs API and browser console. OIDC optional (EOA_OIDC_ISSUER).",
    version="0.6.0",
    lifespan=lifespan,
)


_s = get_settings()
_session_key = _s.session_secret or "dev-insecure-change-EOA_SESSION_SECRET"
# PKCE session cookie: Secure when public URL is HTTPS (e.g. eoa.knospe.org)
_redirect = (_s.oidc_redirect_uri or "").strip()
_session_https = _redirect.casefold().startswith("https://")
app.add_middleware(
    SessionMiddleware,
    secret_key=_session_key,
    max_age=3600,
    same_site="lax",
    https_only=_session_https,
)


def _cors_origins() -> list[str]:
    s = get_settings().cors_origins.strip()
    if s == "*":
        return ["*"]
    return [x.strip() for x in s.split(",") if x.strip()]


_origins = _cors_origins()
app.add_middleware(
    CORSMiddleware,
    allow_origins=_origins,
    # Browsers forbid credentials with wildcard origin
    allow_credentials=_origins != ["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

app.include_router(jobs.router, prefix="/api/v1")
app.include_router(auth_oidc.router, prefix="/api/v1")


@app.get("/")
def index_page() -> FileResponse:
    return FileResponse(STATIC_DIR / "index.html")


app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")


@app.get("/health")
def health() -> dict[str, str]:
    return {"status": "ok"}


@app.get("/ready")
def ready() -> dict[str, str]:
    """Readiness: database reachable (use behind load balancers / k8s)."""
    try:
        with engine.connect() as conn:
            conn.execute(text("SELECT 1"))
    except Exception as e:
        raise HTTPException(status_code=503, detail=f"not_ready: {e!s}") from e
    return {"status": "ready", "db": "ok"}


@app.get("/api/v1/me")
def me(sub: str | None = Depends(require_user)) -> dict[str, str | None]:
    """Return current subject when OIDC is configured; otherwise dev mode."""
    return {"sub": sub, "auth": "oidc" if get_settings().oidc_issuer else "disabled"}
