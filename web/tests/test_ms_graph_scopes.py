"""Tests for Microsoft Graph SPA scope resolution."""

from __future__ import annotations

from app.config import Settings
from app.ms_graph_spa import DELEGATED_GRAPH_SCOPES, resolve_delegated_graph_scopes, resolve_ms_graph_spa_client_id


def test_resolve_delegated_graph_scopes_default() -> None:
    s = Settings(ms_graph_delegated_scopes="")
    assert resolve_delegated_graph_scopes(s) == list(DELEGATED_GRAPH_SCOPES)


def test_resolve_delegated_graph_scopes_custom() -> None:
    s = Settings(ms_graph_delegated_scopes="User.Read, Group.Read.All")
    assert resolve_delegated_graph_scopes(s) == ["User.Read", "Group.Read.All"]


def test_resolve_delegated_graph_scopes_semicolon() -> None:
    s = Settings(ms_graph_delegated_scopes="User.Read; Organization.Read.All")
    assert resolve_delegated_graph_scopes(s) == ["User.Read", "Organization.Read.All"]


def test_resolve_ms_graph_spa_client_id_explicit_wins() -> None:
    s = Settings(
        ms_graph_spa_client_id="aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa",
        ms_graph_spa_use_graph_app_id=True,
        graph_client_id="bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb",
    )
    assert resolve_ms_graph_spa_client_id(s) == "aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa"


def test_resolve_ms_graph_spa_client_id_fallback_to_graph() -> None:
    s = Settings(
        ms_graph_spa_client_id="",
        ms_graph_spa_use_graph_app_id=True,
        graph_client_id="bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb",
    )
    assert resolve_ms_graph_spa_client_id(s) == "bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb"


def test_resolve_ms_graph_spa_client_id_no_fallback_when_flag_off() -> None:
    s = Settings(
        ms_graph_spa_client_id="",
        ms_graph_spa_use_graph_app_id=False,
        graph_client_id="bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb",
    )
    assert resolve_ms_graph_spa_client_id(s) == ""


def test_resolve_ms_graph_spa_client_id_defaults_to_graph_app_when_flag_default() -> None:
    s = Settings(
        ms_graph_spa_client_id="",
        graph_client_id="bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb",
    )
    assert resolve_ms_graph_spa_client_id(s) == "bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb"
