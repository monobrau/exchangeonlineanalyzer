"""Tests for GUI runtime env file (eoa_gui.env)."""

from __future__ import annotations

import pytest

from app.config import get_settings
from app.services import runtime_env_settings as res


@pytest.fixture
def gui_env_isolated(tmp_path, monkeypatch):
    """Redirect GUI env path and skip Settings reload (model_config still points at real path)."""
    p = tmp_path / "eoa_gui.env"
    monkeypatch.setattr(res, "GUI_ENV_FILE", p)
    monkeypatch.setattr(res, "reload_settings", lambda: None)
    yield p
    get_settings.cache_clear()


def test_apply_runtime_patch_bool_and_int(gui_env_isolated):
    out = res.apply_runtime_patch({"EOA_DEBUG": True, "EOA_GRAPH_MAX_TENANTS_PER_JOB": 42})
    assert "EOA_DEBUG" in out["updated_keys"]
    assert gui_env_isolated.is_file()
    text = gui_env_isolated.read_text(encoding="utf-8")
    assert "EOA_DEBUG=true" in text
    assert "EOA_GRAPH_MAX_TENANTS_PER_JOB=42" in text


def test_apply_runtime_patch_secret_clear(gui_env_isolated):
    gui_env_isolated.write_text("EOA_GRAPH_CLIENT_SECRET=old\n", encoding="utf-8")
    out = res.apply_runtime_patch({"EOA_GRAPH_CLIENT_SECRET": ""})
    assert out["updated_keys"]
    assert not gui_env_isolated.exists()


def test_apply_runtime_patch_unknown_key():
    with pytest.raises(ValueError, match="Unknown"):
        res.apply_runtime_patch({"EOA_NOT_A_REAL_KEY": "x"})


def test_parse_format_roundtrip(tmp_path):
    p = tmp_path / "t.env"
    p.write_text('A=1\nB="x y"\n', encoding="utf-8")
    d = res.parse_env_file(p)
    assert d["A"] == "1"
    assert d["B"] == "x y"
    p2 = tmp_path / "out.env"
    res.atomic_write_text(p2, res.format_env_file(d))
    d2 = res.parse_env_file(p2)
    assert d2 == d


def test_build_runtime_payload_shape():
    payload = res.build_runtime_payload()
    assert "items" in payload
    assert "gui_env_file" in payload
    keys = {x["env_key"] for x in payload["items"]}
    assert "EOA_USE_PWSH_STUB_WORKER" in keys
    assert "EOA_GRAPH_CLIENT_SECRET" in keys
    assert "EOA_MS_GRAPH_DELEGATED_SCOPES" in keys
    assert "EOA_EXO_APP_ID" in keys
    assert "EOA_JOB_DEFAULT_TENANT_ID" in keys
    secret = next(x for x in payload["items"] if x["env_key"] == "EOA_GRAPH_CLIENT_SECRET")
    assert secret["kind"] == "secret"
    assert secret["value"] is None
    assert "has_value" in secret
