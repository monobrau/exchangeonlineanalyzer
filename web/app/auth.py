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


async def require_user(
    request: Request,
    creds: Annotated[HTTPAuthorizationCredentials | None, Depends(security)],
) -> str | None:
    """Return Authentik `sub` if JWT is valid; if OIDC not configured, return None (open API)."""
    settings = get_settings()
    if not settings.oidc_issuer:
        return None
    # Prefer HttpOnly cookie (set on OIDC callback) over Authorization header.
    # Stale opaque tokens left in sessionStorage would otherwise win and always 401.
    token = (request.cookies.get(ACCESS_TOKEN_COOKIE_NAME) or "").strip()
    if not token and creds and creds.credentials:
        token = (creds.credentials or "").strip()
    if not token:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Missing bearer token or session cookie (re-sign in)",
            headers={"WWW-Authenticate": "Bearer"},
        )
    try:
        # Opaque OAuth tokens are not JWTs — fail clearly (callback should store id_token or JWT access).
        unverified = jwt.decode(token, options={"verify_signature": False})
        header = jwt.get_unverified_header(token)
    except Exception as e:
        raise HTTPException(
            status_code=401,
            detail=f"Invalid token: not a JWT (opaque access token?). Re-sign in. ({e!s})",
        ) from e

    meta = get_oidc_metadata(settings.oidc_issuer)
    iss_disc = str(meta["issuer"]).strip()
    token_iss = unverified.get("iss")
    if not token_iss:
        raise HTTPException(status_code=401, detail="Invalid token: missing iss claim")
    # Discovery issuer and JWT iss often differ only by trailing slash — compare normalized.
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
        # Authentik / some IdPs issue HS256 access tokens signed with the OAuth client secret.
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
            jwks_uri = str(meta["jwks_uri"])
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
