"""Unit tests for graph_worker (no live Microsoft identity calls)."""

from __future__ import annotations

from unittest.mock import patch

from app.services.graph_worker import graph_worker_configured, run_graph_bulk_job


def test_graph_worker_configured_requires_id_and_secret() -> None:
    from app.config import Settings

    assert graph_worker_configured(Settings(graph_client_id="", graph_client_secret="")) is False
    assert graph_worker_configured(Settings(graph_client_id="a", graph_client_secret="")) is False
    assert graph_worker_configured(Settings(graph_client_id="", graph_client_secret="b")) is False
    assert graph_worker_configured(Settings(graph_client_id="a", graph_client_secret="b")) is True


def test_run_graph_bulk_job_fails_without_tenant_ids() -> None:
    class _Job:
        request_payload: dict | None = {"tenant_ids": []}

    ok, log, _uri = run_graph_bulk_job("no-tenant-job", _Job())  # type: ignore[arg-type]
    assert ok is False
    assert "tenant_ids" in log.lower() or "tenant" in log.lower()


@patch("app.services.graph_worker.acquire_graph_token")
@patch("app.services.graph_worker._graph_get_json")
def test_run_graph_bulk_job_success_path(mock_get, mock_token, monkeypatch) -> None:
    from app.config import Settings, get_settings

    monkeypatch.setenv("EOA_GRAPH_CLIENT_ID", "cid")
    monkeypatch.setenv("EOA_GRAPH_CLIENT_SECRET", "sec")
    get_settings.cache_clear()

    mock_token.return_value = ("fake-token", None)
    mock_get.return_value = (
        {"value": [{"displayName": "Contoso", "id": "org-guid", "verifiedDomains": []}]},
        200,
        None,
    )

    class _Job:
        request_payload = {
            "tenant_ids": ["aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa"],
            "options": {"reports": ["rules"]},
        }

    ok, _log, uri = run_graph_bulk_job("ok-job", _Job())  # type: ignore[arg-type]
    assert ok is True
    assert "summary.json" in uri
    mock_token.assert_called_once()
    mock_get.assert_called_once()

    get_settings.cache_clear()


@patch("app.services.graph_worker.acquire_graph_token")
@patch("app.services.graph_worker._graph_get_json")
def test_run_graph_bulk_job_two_tenants(mock_get, mock_token, monkeypatch) -> None:
    from app.config import get_settings

    monkeypatch.setenv("EOA_GRAPH_CLIENT_ID", "cid")
    monkeypatch.setenv("EOA_GRAPH_CLIENT_SECRET", "sec")
    get_settings.cache_clear()

    mock_token.return_value = ("fake-token", None)
    mock_get.return_value = (
        {"value": [{"displayName": "Contoso", "id": "org-guid", "verifiedDomains": []}]},
        200,
        None,
    )

    class _Job:
        request_payload = {
            "tenant_ids": [
                "aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa",
                "bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb",
            ],
            "options": {"reports": ["organization"]},
        }

    ok, log, _uri = run_graph_bulk_job("multi-job", _Job())  # type: ignore[arg-type]
    assert ok is True
    assert mock_token.call_count == 2
    assert mock_get.call_count == 2
    assert "multi=True" in log

    get_settings.cache_clear()
