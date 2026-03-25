"""Tests for options.reports parsing and graph paging helpers."""

from __future__ import annotations

from unittest.mock import patch

from app.services.graph_worker import graph_get_all_pages, parse_requested_reports


def test_parse_requested_reports_defaults() -> None:
    assert parse_requested_reports(None) == ["organization"]
    assert parse_requested_reports({}) == ["organization"]
    assert parse_requested_reports({"reports": None}) == ["organization"]
    assert parse_requested_reports({"reports": []}) == ["organization"]


def test_parse_requested_reports_aliases_and_order() -> None:
    assert parse_requested_reports({"reports": ["org", "users"]}) == ["organization", "users"]
    assert parse_requested_reports({"reports": ["ca", "applications"]}) == [
        "conditional_access",
        "applications",
    ]


def test_parse_requested_reports_unknown_falls_back_to_organization() -> None:
    assert parse_requested_reports({"reports": ["rules", "nope"]}) == ["organization"]


def test_parse_requested_reports_dedupes() -> None:
    assert parse_requested_reports({"reports": ["users", "user", "users"]}) == ["users"]


@patch("app.services.graph_worker._graph_get_json")
def test_graph_get_all_pages_follows_next_link(mock_get) -> None:
    mock_get.side_effect = [
        (
            {
                "value": [{"id": "1"}],
                "@odata.nextLink": "https://graph.microsoft.com/v1.0/users?$skiptoken=abc",
            },
            200,
            None,
        ),
        ({"value": [{"id": "2"}]}, 200, None),
    ]
    items, err = graph_get_all_pages("https://graph.microsoft.com/v1.0/users", "tok")
    assert err is None
    assert len(items) == 2
    assert mock_get.call_count == 2
