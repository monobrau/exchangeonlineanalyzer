import json
from contextlib import asynccontextmanager
from pathlib import Path

from fastapi import Depends, FastAPI, HTTPException, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import HTMLResponse, RedirectResponse
from fastapi.staticfiles import StaticFiles
from sqlalchemy import text
from starlette.middleware.sessions import SessionMiddleware

from app.auth import require_user
from app.config import get_settings
from app.ms_graph_spa import resolve_ms_graph_spa_client_id
from app.db import engine, init_db
from app.routers import auth_oidc, connections, export_meta, jobs, settings_env

STATIC_DIR = Path(__file__).resolve().parent.parent / "static"

_HTML_NO_STORE = {
    "Cache-Control": "no-store, no-cache, must-revalidate",
    "Pragma": "no-cache",
    "CDN-Cache-Control": "no-store",
}


def _inject_asset_version(html: str, version: str) -> str:
    """Bust CDN/browser caches: CSS/JS URLs change when API version changes."""
    return html.replace("{{EOA_ASSET_V}}", version)


def _inject_msal_bootstrap(html: str) -> str:
    """Expose resolved SPA client id in HTML so MSAL can start without relying on /auth/msal-config alone."""
    s = get_settings()
    cid = resolve_ms_graph_spa_client_id(s)
    ten = (s.ms_graph_tenant or "organizations").strip() or "organizations"
    authority = f"https://login.microsoftonline.com/{ten}"
    payload = {
        "clientId": cid if cid else None,
        "authority": authority,
        "ms_graph_tenant": ten,
    }
    script = f'<script id="eoa-msal-bootstrap">window.__EOA_MSAL_BOOTSTRAP__={json.dumps(payload)};</script>'
    if "</head>" in html:
        return html.replace("</head>", f"{script}\n</head>", 1)
    return script + html


def _oidc_browser_ready() -> bool:
    s = get_settings()
    return bool(s.oidc_issuer and s.oidc_client_id and s.oidc_redirect_uri)


@asynccontextmanager
async def lifespan(_: FastAPI):
    init_db()
    yield


settings = get_settings()
app = FastAPI(
    title=settings.app_name,
    description="Bulk export jobs API and browser console. OIDC optional (EOA_OIDC_ISSUER).",
    version="0.9.0",
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
app.include_router(export_meta.router, prefix="/api/v1")
app.include_router(auth_oidc.router, prefix="/api/v1")
app.include_router(settings_env.router, prefix="/api/v1")
app.include_router(connections.router, prefix="/api/v1")


@app.middleware("http")
async def _cache_control_headers(request: Request, call_next):
    """Discourage edge caches (Cloudflare) from serving stale HTML/CSS/JS."""
    response = await call_next(request)
    p = request.url.path
    if p.startswith("/static/") and (p.endswith(".js") or p.endswith(".css")):
        response.headers["Cache-Control"] = "no-cache, must-revalidate"
        response.headers["CDN-Cache-Control"] = "no-cache"
    ct = (response.headers.get("content-type") or "").lower()
    if "text/html" in ct:
        for k, v in _HTML_NO_STORE.items():
            response.headers[k] = v
    return response


def _oidc_locked_console_html() -> HTMLResponse:
    """Full console HTML with locked shell until sessionStorage has a bearer token."""
    path = STATIC_DIR / "index.html"
    html = path.read_text(encoding="utf-8")
    if '<body class="auth-locked">' not in html:
        html = html.replace("<body>", '<body class="auth-locked">', 1)
    html = html.replace(
        'id="auth-gate" class="auth-gate panel" hidden',
        'id="auth-gate" class="auth-gate panel"',
        1,
    )
    html = html.replace(
        '<main class="layout" id="app-main">',
        '<main class="layout" id="app-main" hidden>',
        1,
    )
    # View source: confirms which API build served this page (compare to git deploy).
    html = f"<!-- eoa-console build=api-{app.version} template=index.html has-ms-graph-outer -->\n" + html
    html = _inject_asset_version(html, app.version)
    # Same as / when OIDC is off: expose resolved SPA client id for MSAL (public data only).
    html = _inject_msal_bootstrap(html)
    return HTMLResponse(
        content=html,
        media_type="text/html; charset=utf-8",
        headers=dict(_HTML_NO_STORE),
    )


@app.get("/", response_model=None)
def root_page() -> HTMLResponse:
    """OIDC: minimal landing only. No OIDC: full console at / (local dev)."""
    if _oidc_browser_ready():
        raw = (STATIC_DIR / "landing.html").read_text(encoding="utf-8")
        body = _inject_asset_version(raw, app.version)
        return HTMLResponse(content=body, media_type="text/html; charset=utf-8", headers=dict(_HTML_NO_STORE))
    raw = (STATIC_DIR / "index.html").read_text(encoding="utf-8")
    body = _inject_asset_version(raw, app.version)
    body = _inject_msal_bootstrap(body)
    return HTMLResponse(content=body, media_type="text/html; charset=utf-8", headers=dict(_HTML_NO_STORE))


@app.get("/app", response_model=None)
def app_console_page() -> HTMLResponse | RedirectResponse:
    """Full bulk-export console. OIDC: locked shell until token in sessionStorage. No OIDC: use / instead."""
    if not _oidc_browser_ready():
        return RedirectResponse("/", status_code=302)
    return _oidc_locked_console_html()


@app.get("/health")
def health() -> dict[str, str]:
    return {"status": "ok"}


@app.get("/api/v1/ui-info")
def ui_info() -> dict[str, object]:
    """Which UI files this process reads from disk — curl the live server to verify deploy (no auth)."""
    s = get_settings()
    index = STATIC_DIR / "index.html"
    app_js = STATIC_DIR / "app.js"
    ms_js = STATIC_DIR / "ms-graph.js"
    index_html: dict[str, bool] = {
        "exists": index.is_file(),
        "has_ms_graph_outer": False,
        "has_ms_graph_mount": False,
    }
    if index.is_file():
        t = index.read_text(encoding="utf-8")
        index_html["has_ms_graph_outer"] = "ms-graph-outer" in t
        index_html["has_ms_graph_mount"] = "ms-graph-mount" in t
    app_js_info: dict[str, bool] = {"exists": app_js.is_file(), "has_dynamic_ms_graph_import": False}
    if app_js.is_file():
        aj = app_js.read_text(encoding="utf-8")
        app_js_info["has_dynamic_ms_graph_import"] = "ms-graph.js" in aj and "import(" in aj
    ms_graph_js_info: dict[str, bool] = {"exists": ms_js.is_file(), "has_dynamic_msal_loader": False}
    if ms_js.is_file():
        mj = ms_js.read_text(encoding="utf-8")
        ms_graph_js_info["has_dynamic_msal_loader"] = "importMsalBrowser" in mj
    return {
        "api_version": app.version,
        "static_dir": str(STATIC_DIR.resolve()),
        "repo_root_config": str(s.repo_root.resolve()),
        "ms_graph_spa_client_id_configured": bool(resolve_ms_graph_spa_client_id(s)),
        "index_html": index_html,
        "app_js": app_js_info,
        "ms_graph_js": ms_graph_js_info,
    }


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


# Mount static last so /api/v1/* and /health are never shadowed by the static app.
app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")
