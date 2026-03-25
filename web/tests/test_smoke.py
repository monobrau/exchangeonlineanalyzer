from fastapi.testclient import TestClient

from app.main import app


def test_export_options_schema() -> None:
    with TestClient(app) as client:
        r = client.get("/api/v1/export/options-schema")
    assert r.status_code == 200
    body = r.json()
    assert body.get("schema_version") == 1
    assert "json_schema" in body
    assert body["json_schema"].get("title") == "WebBulkExportOptions"


def test_health() -> None:
    with TestClient(app) as client:
        r = client.get("/health")
    assert r.status_code == 200
    assert r.json()["status"] == "ok"


def test_auth_status() -> None:
    with TestClient(app) as client:
        r = client.get("/api/v1/auth/status")
    assert r.status_code == 200
    body = r.json()
    assert "oidc_login_enabled" in body
    assert body["oidc_login_enabled"] is False
    assert body.get("ms_graph_spa_enabled") is False


def test_ui_info() -> None:
    with TestClient(app) as client:
        r = client.get("/api/v1/ui-info")
    assert r.status_code == 200
    body = r.json()
    assert body.get("api_version")
    assert body["index_html"]["exists"] is True
    assert body["index_html"]["has_ms_graph_outer"] is True
    assert body["app_js"]["has_dynamic_ms_graph_import"] is True
    assert body["ms_graph_js"]["has_dynamic_msal_loader"] is True


def test_connections_status() -> None:
    with TestClient(app) as client:
        r = client.get("/api/v1/connections/status")
    assert r.status_code == 200
    body = r.json()
    assert "graph_app_configured" in body
    assert body["graph_app_configured"] is False
    assert body["exo"] in ("skipped", "ready", "not_configured")
    assert body.get("job_default_tenant_id") is None


def test_msal_config_when_disabled() -> None:
    with TestClient(app) as client:
        r = client.get("/api/v1/auth/msal-config")
    assert r.status_code == 200
    body = r.json()
    assert body.get("enabled") is False
    assert len(body.get("scopes") or []) == 3
    assert body.get("ms_graph_tenant") == "organizations"
    assert "login.microsoftonline.com" in (body.get("suggestedAuthority") or "")
    assert "fallback_client_id" in body


def test_oidc_login_not_configured() -> None:
    with TestClient(app) as client:
        r = client.get("/api/v1/auth/oidc/login")
    assert r.status_code == 501


def test_ready() -> None:
    with TestClient(app) as client:
        r = client.get("/ready")
    assert r.status_code == 200
    assert r.json().get("status") == "ready"


def test_create_and_get_job() -> None:
    with TestClient(app) as client:
        r = client.post(
            "/api/v1/jobs/bulk",
            json={"tenant_ids": ["11111111-1111-1111-1111-111111111111"], "options": {"reports": ["rules"]}},
        )
        assert r.status_code == 201
        jid = r.json()["id"]
        g = client.get(f"/api/v1/jobs/{jid}")
    assert g.status_code == 200
    body = g.json()
    assert body["id"] == jid
    # Placeholder worker runs inline under TestClient
    assert body["status"] == "succeeded"
    assert "placeholder.txt" in (body.get("artifact_files") or [])
    rp = body.get("request_payload")
    assert rp is not None
    assert rp.get("tenant_ids") == ["11111111-1111-1111-1111-111111111111"]


def test_index_and_static() -> None:
    with TestClient(app) as client:
        r = client.get("/")
        assert r.status_code == 200
        assert b"Bulk export jobs" in r.content
        assert b"eoa-msal-bootstrap" in r.content
        assert b"__EOA_MSAL_BOOTSTRAP__" in r.content
        assert b"Connections (server)" in r.content
        s = client.get("/static/styles.css")
        assert s.status_code == 200


def test_list_job_artifacts() -> None:
    with TestClient(app) as client:
        r = client.post(
            "/api/v1/jobs/bulk",
            json={"tenant_ids": ["33333333-3333-3333-3333-333333333333"], "options": {}},
        )
        jid = r.json()["id"]
        lst = client.get(f"/api/v1/jobs/{jid}/artifacts")
    assert lst.status_code == 200
    assert "placeholder.txt" in lst.json()["files"]


def test_download_placeholder_artifact() -> None:
    with TestClient(app) as client:
        r = client.post(
            "/api/v1/jobs/bulk",
            json={"tenant_ids": ["22222222-2222-2222-2222-222222222222"], "options": {}},
        )
        assert r.status_code == 201
        jid = r.json()["id"]
        dl = client.get(f"/api/v1/jobs/{jid}/artifact", params={"file": "placeholder.txt"})
    assert dl.status_code == 200
    assert b"Placeholder worker" in dl.content or len(dl.content) > 0
