"""Tests for Microsoft Graph SPA scope resolution."""

from __future__ import annotations

from app.config import Settings
from app.ms_graph_spa import DELEGATED_GRAPH_SCOPES, resolve_delegated_graph_scopes


def test_resolve_delegated_graph_scopes_default() -> None:
    s = Settings(ms_graph_delegated_scopes="")
    assert resolve_delegated_graph_scopes(s) == list(DELEGATED_GRAPH_SCOPES)


def test_resolve_delegated_graph_scopes_custom() -> None:
    s = Settings(ms_graph_delegated_scopes="User.Read, Group.Read.All")
    assert resolve_delegated_graph_scopes(s) == ["User.Read", "Group.Read.All"]


def test_resolve_delegated_graph_scopes_semicolon() -> None:
    s = Settings(ms_graph_delegated_scopes="User.Read; Organization.Read.All")
    assert resolve_delegated_graph_scopes(s) == ["User.Read", "Organization.Read.All"]
