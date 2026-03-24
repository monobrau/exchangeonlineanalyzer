# Authentik for Exchange Online Analyzer (EOA)

The web app uses **OpenID Connect (OIDC)** with **Authorization Code + PKCE** for the browser login, and validates **Bearer JWT access tokens** on `/api/v1/*` when `EOA_OIDC_ISSUER` is set.

**Public URL:** `https://eoa.knospe.org` (adjust if your hostname differs).

---

## 1. Create an Application in Authentik

1. **Applications** → **Applications** → **Create with Provider**.
2. **Application name:** e.g. `Exchange Online Analyzer`.
3. **Provider type:** **OAuth2/OpenID Provider** (or **OIDC** if you only see that option).
4. Choose the **Authorization flow** (e.g. default authorization flow) and **Signing key** (default is fine).

---

## 2. Provider settings (OAuth2 / OIDC)

Use these as a baseline; exact labels depend on your Authentik version.

| Setting | Value |
|--------|--------|
| **Client type** | **Public** (recommended). PKCE is used; `EOA_OIDC_CLIENT_SECRET` can stay empty. |
| **Redirect URIs** | `https://eoa.knospe.org/api/v1/auth/oidc/callback` — **exact** match, including `https`, path, and no trailing slash. |
| **Signing Key** | Default RSA key (access tokens must be **JWT** signed with RS256/ES256 for the API). |

**Scopes:** The app requests `openid profile email` (default `EOA_OIDC_SCOPE`). Ensure these scopes exist on the provider (defaults usually do).

**Save** the provider and application.

---

## 3. Values to copy from Authentik

After the provider exists, open the **provider** (not only the application).

1. **Issuer** / OpenID configuration  
   - Often shown as **`https://<your-authentik-host>/application/o/<provider-slug>/`**  
   - **This deployment:** `https://auth.knospe.org/application/o/eoa/`  
   - Copy this **exact** string (including trailing `/` if Authentik shows it).  
   - Sanity check: open  
     `https://auth.knospe.org/application/o/eoa/.well-known/openid-configuration`  
     in a browser — it must return JSON.

2. **Client ID**  
   - Copy the OAuth2 **Client ID** (UUID string).

3. **Audience for API validation**  
   - The app **defaults** JWT audience validation to **`EOA_OIDC_CLIENT_ID`** when `EOA_OIDC_AUDIENCE` is unset (Authentik’s default `aud` on access tokens).  
   - Set **`EOA_OIDC_AUDIENCE`** (or **`EOA_OIDC_AUDIENCES`**) only if you customize `aud` via scope mappings.

4. **Client secret**  
   - **Public client:** leave empty in `.env`.  
   - **Confidential client:** if you enable a secret, set `EOA_OIDC_CLIENT_SECRET` to match.

---

## 4. Server `web/.env` on webhost

Copy [`env.eoa.knospe.example`](env.eoa.knospe.example) to `web/.env` and set:

```env
EOA_OIDC_ISSUER=https://auth.knospe.org/application/o/eoa/
EOA_OIDC_CLIENT_ID=<client-id-from-authentik>
# EOA_OIDC_AUDIENCE=  # optional; defaults to client id
EOA_OIDC_REDIRECT_URI=https://eoa.knospe.org/api/v1/auth/oidc/callback
EOA_CORS_ORIGINS=https://eoa.knospe.org
EOA_SESSION_SECRET=<long random string>
```

Discovery URL (sanity check): `https://auth.knospe.org/application/o/eoa/.well-known/openid-configuration`

Optional:

```env
# EOA_OIDC_CLIENT_SECRET=   # only confidential clients
# EOA_OIDC_SCOPE=openid profile email
```

Restart the API:

```bash
sudo systemctl restart eoa-api
```

---

## 5. Smoke tests

| Check | URL / action |
|--------|----------------|
| OIDC enabled | `GET https://eoa.knospe.org/api/v1/auth/status` → `"oidc_login_enabled": true` |
| OIDC discovery | `GET https://eoa.knospe.org/api/v1/me` without `Authorization` → **401** when issuer is set |
| Browser login | Open `/`, **Sign in** → Authentik → back to app with token in `sessionStorage` (`eoa_bearer`) |
| API | `GET /api/v1/jobs` with `Authorization: Bearer <access_token>` → **200** |

---

## 6. Troubleshooting

| Symptom | What to check |
|--------|----------------|
| **501** on `/api/v1/auth/oidc/login` | `EOA_OIDC_ISSUER`, `EOA_OIDC_CLIENT_ID`, `EOA_OIDC_REDIRECT_URI` set and service restarted. |
| **Redirect URI mismatch** | Authentik redirect URI must match **character-for-character** (scheme, host, path). |
| **401** `Invalid token` after login | **`EOA_OIDC_AUDIENCE`** must match JWT `aud` (usually client id). **Issuer** must match token `iss` exactly (including trailing slash). |
| **Token exchange failed** (HTML error page) | Public client: no client secret in Authentik; leave `EOA_OIDC_CLIENT_SECRET` empty. Confidential: secret must match. |
| PKCE / session errors | **`EOA_SESSION_SECRET`** stable across restarts; browser must accept cookies on `eoa.knospe.org` (SameSite `lax`). |

---

## 7. Local development

Use a separate Authentik redirect URI, e.g.:

`http://localhost:8000/api/v1/auth/oidc/callback`

Add that URI in the Authentik provider. Run uvicorn locally with the same `EOA_*` values, but `EOA_OIDC_REDIRECT_URI` and `EOA_CORS_ORIGINS` pointing at your dev URL.
