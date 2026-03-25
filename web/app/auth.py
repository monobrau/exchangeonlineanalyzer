"""Optional OIDC JWT validation (Authentik). When EOA_OIDC_ISSUER is unset, auth is disabled for local dev."""

from typing import Annotated

import jwt
from fastapi import Depends, HTTPException, Request, status
from fastapi.security import HTTPAuthorizationCredentials, HTTPBearer
from jwt import PyJWKClient

from app.config import get_settings
from app.oidc_metadata import get_oidc_metadata

# HttpOnly cookie set on OIDC callback — survives when reverse proxies strip Authorization headers.
ACCESS_TOKEN_COOKIE_NAME = "eoa_access_token"
# Set on successful OAuth callback (same signed session as PKCE); used when JWT cookie is missing/stripped.
OIDC_SUB_SESSION_KEY = "oidc_sub"

security = HTTPBearer(auto_error=False)
_jwks_by_uri: dict[str, PyJWKClient] = {}


def _jwks_client_for(jwks_uri: str) -> PyJWKClient:
    if jwks_uri not in _jwks_by_uri:
        _jwks_by_uri[jwks_uri] = PyJWKClient(jwks_uri)
    return _jwks_by_uri[jwks_uri]


def _audiences(settings) -> list[str]:
    if settings.oidc_audiences:
        return [a.strip() for a in settings.oidc_audiences.split(",") if a.strip()]
    if settings.oidc_audience:
        return [settings.oidc_audience]
    # Authentik default access-token aud is the OAuth2 client id
    if settings.oidc_client_id:
        return [settings.oidc_client_id]
    return []


def validate_oidc_jwt_token(token: str) -> str:
    """Validate a Bearer/cookie JWT and return `sub`. Raises HTTPException(401) on failure."""
    settings = get_settings()
    if not settings.oidc_issuer:
        raise HTTPException(status_code=401, detail="OIDC not configured")
    token = (token or "").strip()
    if not token:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Missing bearer token or session cookie (re-sign in)",
            headers={"WWW-Authenticate": "Bearer"},
        )
    try:
        unverified = jwt.decode(token, options={"verify_signature": False})
        header = jwt.get_unverified_header(token)
    except Exception as e:
        raise HTTPException(
            status_code=401,
            detail=f"Invalid token: not a JWT (opaque access token?). Re-sign in. ({e!s})",
        ) from e

    try:
        meta = get_oidc_metadata(settings.oidc_issuer)
    except Exception as e:
        raise HTTPException(
            status_code=503,
            detail=f"OIDC discovery failed (cannot reach issuer metadata): {e!s}",
        ) from e

    iss_disc = str(meta.get("issuer") or "").strip()
    if not iss_disc:
        raise HTTPException(status_code=503, detail="OIDC discovery document missing issuer")
    token_iss = unverified.get("iss")
    if not token_iss:
        raise HTTPException(status_code=401, detail="Invalid token: missing iss claim")
    if token_iss.rstrip("/") != iss_disc.rstrip("/"):
        raise HTTPException(
            status_code=401,
            detail=(
                f"Invalid token: issuer mismatch (token iss={token_iss!r}, "
                f"discovery issuer={iss_disc!r})"
            ),
        )

    audiences = _audiences(settings)
    aud_kw: dict = {}
    if audiences:
        aud_kw["audience"] = audiences if len(audiences) > 1 else audiences[0]

    alg = (header or {}).get("alg") or ""

    try:
        if alg == "HS256":
            if not settings.oidc_client_secret:
                raise HTTPException(
                    status_code=401,
                    detail=(
                        "Invalid token: HS256 JWT requires EOA_OIDC_CLIENT_SECRET on the API "
                        "(confidential client) to verify."
                    ),
                )
            payload = jwt.decode(
                token,
                settings.oidc_client_secret,
                algorithms=["HS256"],
                issuer=token_iss,
                **aud_kw,
            )
        else:
            jwks_uri = str(meta.get("jwks_uri") or "").strip()
            if not jwks_uri:
                raise HTTPException(status_code=503, detail="OIDC discovery document missing jwks_uri")
            jwks = _jwks_client_for(jwks_uri)
            signing_key = jwks.get_signing_key_from_jwt(token)
            payload = jwt.decode(
                token,
                signing_key.key,
                algorithms=["RS256", "ES256"],
                issuer=token_iss,
                **aud_kw,
            )
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=401, detail=f"Invalid token: {e!s}") from e

    sub = payload.get("sub")
    if not sub:
        raise HTTPException(status_code=401, detail="Token missing sub")
    return str(sub)


def extract_sub_after_token_exchange(token: str) -> str:
    """
    Get `sub` from a JWT returned by the IdP token endpoint without JWKS (no signature verification).

    Only safe immediately after a successful authorization_code → token exchange: the token was
    just received over HTTPS from the IdP. Use when JWKS fetch is blocked (403 from Cloudflare, etc.).
    """
    settings = get_settings()
    if not settings.oidc_issuer:
        raise HTTPException(status_code=401, detail="OIDC not configured")
    token = (token or "").strip()
    if not token:
        raise HTTPException(status_code=401, detail="Empty token from IdP")
    try:
        payload = jwt.decode(token, options={"verify_signature": False})
    except Exception as e:
        raise HTTPException(status_code=401, detail=f"Not a valid JWT: {e!s}") from e

    try:
        meta = get_oidc_metadata(settings.oidc_issuer)
    except Exception as e:
        raise HTTPException(
            status_code=503,
            detail=f"OIDC discovery failed: {e!s}",
        ) from e

    iss_disc = str(meta.get("issuer") or "").strip()
    token_iss = payload.get("iss")
    if not token_iss:
        raise HTTPException(status_code=401, detail="Token missing iss claim")
    if token_iss.rstrip("/") != iss_disc.rstrip("/"):
        raise HTTPException(
            status_code=401,
            detail="Token issuer does not match OIDC discovery issuer",
        )

    audiences = _audiences(settings)
    if audiences:
        aud_claim = payload.get("aud")
        if aud_claim is not None:
            aud_list = aud_claim if isinstance(aud_claim, list) else [aud_claim]
            if not any(str(x) in audiences for x in aud_list):
                raise HTTPException(status_code=401, detail="Token audience does not match client")

    sub = payload.get("sub")
    if not sub:
        raise HTTPException(status_code=401, detail="Token missing sub")
    return str(sub)


async def require_user(
    request: Request,
    creds: Annotated[HTTPAuthorizationCredentials | None, Depends(security)],
) -> str | None:
    """Return Authentik `sub` if JWT is valid; if OIDC not configured, return None (open API)."""
    settings = get_settings()
    if not settings.oidc_issuer:
        return None
    # Prefer HttpOnly cookie over Authorization (stale sessionStorage Bearer loses to cookie order).
    token = (request.cookies.get(ACCESS_TOKEN_COOKIE_NAME) or "").strip()
    if not token and creds and creds.credentials:
        token = (creds.credentials or "").strip()

    if token:
        try:
            return validate_oidc_jwt_token(token)
        except HTTPException:
            # JWKS blocked (403) or invalid cookie — use server session from OAuth callback if present.
            sub_session = request.session.get(OIDC_SUB_SESSION_KEY)
            if sub_session:
                return str(sub_session)
            raise

    sub_session = request.session.get(OIDC_SUB_SESSION_KEY)
    if sub_session:
        return str(sub_session)

    raise HTTPException(
        status_code=status.HTTP_401_UNAUTHORIZED,
        detail="Missing bearer token or session cookie (re-sign in)",
        headers={"WWW-Authenticate": "Bearer"},
    )
