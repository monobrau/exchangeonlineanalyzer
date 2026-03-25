"""Browser OIDC login (Authorization Code + PKCE). Configure EOA_OIDC_* + EOA_SESSION_SECRET."""

from __future__ import annotations

import hashlib
import html
import json
import secrets
import base64
from urllib.parse import urlencode

import httpx
from fastapi import APIRouter, HTTPException, Query, Request
from fastapi.responses import HTMLResponse, RedirectResponse

from app.config import get_settings
from app.oidc_metadata import get_oidc_metadata

router = APIRouter(prefix="/auth", tags=["auth"])


def _looks_like_jwt(token: str) -> bool:
    """Three base64url segments — opaque OAuth tokens are usually not shaped like this."""
    if not token or not isinstance(token, str):
        return False
    parts = token.split(".")
    return len(parts) == 3 and all(len(p) > 0 for p in parts)


def _pick_browser_bearer_token(body: dict) -> str | None:
    """
    Prefer a JWT the API can validate with JWKS (RS256/ES256).

    Many providers (including some Authentik setups) return an opaque access_token; the
    id_token from openid scope is always a JWT. Storing opaque tokens causes 401 on /api/v1/*.
    """
    access = body.get("access_token")
    id_token = body.get("id_token")
    if access and _looks_like_jwt(access):
        return access
    if id_token and _looks_like_jwt(id_token):
        return id_token
    if access:
        return access
    if id_token:
        return id_token
    return None


def _pkce_pair() -> tuple[str, str]:
    verifier = secrets.token_urlsafe(48)[:128]
    digest = hashlib.sha256(verifier.encode("ascii")).digest()
    challenge = base64.urlsafe_b64encode(digest).decode("ascii").rstrip("=")
    return verifier, challenge


def _oidc_ready() -> bool:
    s = get_settings()
    return bool(s.oidc_issuer and s.oidc_client_id and s.oidc_redirect_uri)


@router.get("/status")
def auth_status() -> dict:
    s = get_settings()
    return {
        "oidc_login_enabled": _oidc_ready(),
        "issuer": s.oidc_issuer or None,
    }


@router.get("/oidc/login")
def oidc_login(request: Request) -> RedirectResponse:
    if not _oidc_ready():
        raise HTTPException(
            status_code=501,
            detail="OIDC browser login not configured (EOA_OIDC_ISSUER, EOA_OIDC_CLIENT_ID, EOA_OIDC_REDIRECT_URI)",
        )
    s = get_settings()
    meta = get_oidc_metadata(s.oidc_issuer)
    auth_ep = meta.get("authorization_endpoint")
    if not auth_ep:
        raise HTTPException(status_code=500, detail="OIDC metadata missing authorization_endpoint")

    verifier, challenge = _pkce_pair()
    state = secrets.token_urlsafe(32)
    request.session["oidc_pkce_verifier"] = verifier
    request.session["oidc_state"] = state

    params = {
        "client_id": s.oidc_client_id,
        "redirect_uri": s.oidc_redirect_uri,
        "response_type": "code",
        "scope": s.oidc_scope,
        "state": state,
        "code_challenge": challenge,
        "code_challenge_method": "S256",
    }
    sep = "&" if "?" in auth_ep else "?"
    url = f"{auth_ep}{sep}{urlencode(params)}"
    return RedirectResponse(url, status_code=302)


@router.get("/oidc/callback")
def oidc_callback(
    request: Request,
    code: str | None = Query(None),
    state: str | None = Query(None),
    error: str | None = Query(None),
    error_description: str | None = Query(None),
) -> HTMLResponse:
    if error:
        msg = html.escape(error_description or error)
        return HTMLResponse(
            f'<!DOCTYPE html><html><body><p>Sign-in error: {msg}</p><a href="/">Home</a></body></html>',
            status_code=400,
        )
    if not code or not state:
        raise HTTPException(status_code=400, detail="Missing code or state")

    saved = request.session.get("oidc_state")
    verifier = request.session.get("oidc_pkce_verifier")
    if not saved or not verifier:
        raise HTTPException(status_code=400, detail="Invalid OAuth session (retry login)")
    try:
        ok = secrets.compare_digest(str(saved), str(state))
    except ValueError:
        ok = False
    if not ok:
        raise HTTPException(status_code=400, detail="Invalid OAuth state (retry login)")

    if not _oidc_ready():
        raise HTTPException(status_code=501, detail="OIDC not configured")

    s = get_settings()
    meta = get_oidc_metadata(s.oidc_issuer)
    token_ep = meta.get("token_endpoint")
    if not token_ep:
        raise HTTPException(status_code=500, detail="OIDC metadata missing token_endpoint")

    data = {
        "grant_type": "authorization_code",
        "code": code,
        "redirect_uri": s.oidc_redirect_uri,
        "client_id": s.oidc_client_id,
        "code_verifier": verifier,
    }
    if s.oidc_client_secret:
        data["client_secret"] = s.oidc_client_secret

    with httpx.Client(timeout=30.0) as client:
        tr = client.post(
            token_ep,
            data=data,
            headers={"Accept": "application/json"},
        )
    request.session.pop("oidc_state", None)
    request.session.pop("oidc_pkce_verifier", None)

    if tr.status_code >= 400:
        return HTMLResponse(
            "<!DOCTYPE html><html><body><p>Token exchange failed.</p><pre>"
            + tr.text[:2000]
            + "</pre><a href=\"/\">Home</a></body></html>",
            status_code=400,
        )

    try:
        body = tr.json()
    except Exception:
        raise HTTPException(status_code=502, detail="Invalid token response") from None

    bearer = _pick_browser_bearer_token(body)
    if not bearer:
        raise HTTPException(status_code=502, detail="No access_token or id_token in token response")

    # One-page handoff: store bearer for existing fetch() + /api/v1 calls
    js_token = json.dumps(bearer)
    html = f"""<!DOCTYPE html>
<html lang="en"><head><meta charset="utf-8"/><title>Signed in</title></head>
<body>
<script>
sessionStorage.setItem("eoa_bearer", {js_token});
location.replace("/app");
</script>
<p>Signed in. <a href="/app">Continue</a></p>
</body></html>"""
    return HTMLResponse(content=html, media_type="text/html; charset=utf-8")
