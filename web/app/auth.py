"""Optional OIDC JWT validation (Authentik). When EOA_OIDC_ISSUER is unset, auth is disabled for local dev."""

from typing import Annotated

import httpx
import jwt
from fastapi import Depends, HTTPException, status
from fastapi.security import HTTPAuthorizationCredentials, HTTPBearer
from jwt import PyJWKClient

from app.config import get_settings

security = HTTPBearer(auto_error=False)
_jwks_client: PyJWKClient | None = None


def _get_jwks_client(issuer: str) -> PyJWKClient:
    global _jwks_client
    if _jwks_client is None:
        well_known = issuer.rstrip("/") + "/.well-known/openid-configuration"
        with httpx.Client(timeout=10.0) as client:
            r = client.get(well_known)
            r.raise_for_status()
            jwks_uri = r.json()["jwks_uri"]
        _jwks_client = PyJWKClient(jwks_uri)
    return _jwks_client


def _audiences(settings) -> list[str]:
    if settings.oidc_audiences:
        return [a.strip() for a in settings.oidc_audiences.split(",") if a.strip()]
    if settings.oidc_audience:
        return [settings.oidc_audience]
    return []


async def require_user(
    creds: Annotated[HTTPAuthorizationCredentials | None, Depends(security)],
) -> str | None:
    """Return Authentik `sub` if JWT is valid; if OIDC not configured, return None (open API)."""
    settings = get_settings()
    if not settings.oidc_issuer:
        return None
    if creds is None or not creds.credentials:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Missing bearer token",
            headers={"WWW-Authenticate": "Bearer"},
        )
    token = creds.credentials
    issuer = settings.oidc_issuer.rstrip("/")
    try:
        jwks = _get_jwks_client(issuer)
        signing_key = jwks.get_signing_key_from_jwt(token)
        audiences = _audiences(settings)
        decode_kw: dict = {
            "issuer": issuer,
            "algorithms": ["RS256", "ES256"],
        }
        if audiences:
            aud = audiences if len(audiences) > 1 else audiences[0]
            decode_kw["audience"] = aud
        payload = jwt.decode(
            token,
            signing_key.key,
            **decode_kw,
        )
        sub = payload.get("sub")
        if not sub:
            raise HTTPException(status_code=401, detail="Token missing sub")
        return str(sub)
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=401, detail=f"Invalid token: {e!s}") from e
