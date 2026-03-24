"""Fetch and cache OIDC discovery document."""

from functools import lru_cache

import httpx


@lru_cache(maxsize=8)
def get_oidc_metadata(issuer: str) -> dict[str, str]:
    url = issuer.rstrip("/") + "/.well-known/openid-configuration"
    with httpx.Client(timeout=20.0) as client:
        r = client.get(url)
        r.raise_for_status()
        return r.json()
