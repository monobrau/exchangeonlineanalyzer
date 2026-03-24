"""Optional OIDC JWT validation (Authentik). When EOA_OIDC_ISSUER is unset, auth is disabled for local dev."""

from typing import Annotated

import jwt
from fastapi import Depends, HTTPException, status
from fastapi.security import HTTPAuthorizationCredentials, HTTPBearer
from jwt import PyJWKClient

from app.config import get_settings
from app.oidc_metadata import get_oidc_metadata

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
    try:
        meta = get_oidc_metadata(settings.oidc_issuer)
        jwks_uri = str(meta["jwks_uri"])
        # Must match JWT iss claim (Authentik matches discovery document issuer string)
        iss = str(meta["issuer"])
        jwks = _jwks_client_for(jwks_uri)
        signing_key = jwks.get_signing_key_from_jwt(token)
        audiences = _audiences(settings)
        decode_kw: dict = {
            "issuer": iss,
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
