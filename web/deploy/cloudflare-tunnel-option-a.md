# Cloudflare Tunnel — Option A (same tunnel, new hostname)

Add the EOA web API to your **existing** locally configured tunnel (e.g. `blingus`) with a **new** hostname (e.g. `eoa.knospe.org`). No tunnel migration required.

**Tunnel name vs hostnames:** In Zero Trust, the tunnel may still appear as a single node (e.g. **blingus**) while **Published application** routes list **multiple** URLs (`eoa.knospe.org`, `blingus.knospe.org`, …). That is expected — the tunnel name is only a label; each hostname is routed separately. You do **not** need to rename the tunnel or create a second tunnel for `eoa` when using Option A.

## 1. DNS (Cloudflare dashboard → `knospe.org` → DNS)

Create a record for the new hostname, **same pattern as `blingus.knospe.org`**:

- **Type:** usually **CNAME**
- **Name:** `eoa` (or your chosen subdomain)
- **Target:** same as `blingus` (often `<tunnel-id>.cfargotunnel.com` — copy from the working `blingus` row)
- **Proxy:** Proxied (orange cloud) if that matches `blingus`

## 2. Tunnel ingress (machine running `cloudflared`)

**Webhost (`webhost` / `lan-30-100`):** config is **`/etc/cloudflared/config.yml`** (managed as root). Service: **`cloudflared.service`**. Credentials: **`/root/.cloudflared/<tunnel-id>.json`**.

Committed example matching that layout: [`cloudflared-config.knospe.example.yml`](cloudflared-config.knospe.example.yml).

Add a **hostname block for EOA** and keep **more specific rules before the catch-all**:

```yaml
ingress:
  - hostname: eoa.knospe.org
    service: http://127.0.0.1:18080
  - hostname: blingus.knospe.org
    service: http://localhost:80
  - service: http_status:404
```

**Webhost:** `nginx` already binds **`127.0.0.1:8080`** for the blingus backend, so run **`uvicorn` on `127.0.0.1:18080`** and point the tunnel there (see [`eoa-api.service.example`](eoa-api.service.example)). On a host with nothing on 8080, you can use **8080** for both instead.

Restart after saving:

```bash
sudo systemctl restart cloudflared
sudo systemctl status cloudflared --no-pager
```

## 3. App environment (`web/.env` on the server)

Use the **public HTTPS URL** (no port):

```env
EOA_OIDC_REDIRECT_URI=https://eoa.knospe.org/api/v1/auth/oidc/callback
EOA_CORS_ORIGINS=https://eoa.knospe.org
EOA_SESSION_SECRET=<long random string>
```

Authentik: register the **same** redirect URI for your OIDC application (step-by-step: [`authentik-eoa.md`](authentik-eoa.md)).

## 4. Smoke test

- `https://eoa.knospe.org/health`
- `https://eoa.knospe.org/docs`
- Sign-in flow: **Sign in** → Authentik → back to `/` with token in `sessionStorage`

If you get **502** / **connection refused**, the tunnel is up but nothing listens on the configured `service` URL — fix `uvicorn`/nginx bind and port.
