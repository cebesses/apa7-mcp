"""CLI e instalación: configuración de agentes sin tocar la real del usuario (HOME y APPDATA temporales)."""
import json
import subprocess
import sys

import pytest

from apa7_mcp import install as I


@pytest.fixture()
def fake_home(tmp_path, monkeypatch):
    monkeypatch.setenv("HOME", str(tmp_path))
    monkeypatch.setenv("USERPROFILE", str(tmp_path))
    monkeypatch.setenv("APPDATA", str(tmp_path / "AppData"))
    monkeypatch.setattr(I.Path, "home", classmethod(lambda cls: tmp_path))
    return tmp_path


def test_opencode_merge_preserves_other_entries_and_is_idempotent(fake_home):
    cfg = fake_home / ".config" / "opencode" / "opencode.json"
    cfg.parent.mkdir(parents=True)
    cfg.write_text(json.dumps({"theme": "x", "mcp": {"otro": {"type": "local", "command": ["x"]}}}), encoding="utf-8")
    I.install("opencode")
    I.install("opencode")
    data = json.loads(cfg.read_text(encoding="utf-8"))
    assert data["theme"] == "x" and "otro" in data["mcp"]
    e = data["mcp"]["apa7-tools"]
    assert e["type"] == "local" and e["command"][-2:] == ["-m", "apa7_mcp"] and e["enabled"] is True
    assert e["environment"]["PYTHONUTF8"] == "1"
    assert (fake_home / ".config" / "opencode" / "opencode.json.bak").exists()


@pytest.mark.parametrize("agent", ["claude-desktop", "cursor"])
def test_other_agents_write_mcpservers(fake_home, agent):
    msg = I.install(agent)
    path, _ = I.target(agent)
    assert json.loads(path.read_text(encoding="utf-8"))["mcpServers"]["apa7-tools"]["args"] == ["-m", "apa7_mcp"]
    assert agent in msg


def test_claude_without_cli_writes_project_mcp_json(fake_home, monkeypatch, tmp_path):
    monkeypatch.setattr(I.shutil, "which", lambda _n: None)
    monkeypatch.chdir(tmp_path)
    I.install("claude")
    data = json.loads((tmp_path / ".mcp.json").read_text(encoding="utf-8"))
    assert data["mcpServers"]["apa7-tools"]["type"] == "stdio"


def test_print_mode_writes_nothing(fake_home):
    out = I.install("opencode", print_only=True)
    assert "apa7-tools" in out and not (fake_home / ".config").exists()


def test_invalid_json_and_unknown_agent(fake_home):
    cfg = fake_home / ".config" / "opencode" / "opencode.json"
    cfg.parent.mkdir(parents=True)
    cfg.write_text("{ esto no es json", encoding="utf-8")
    with pytest.raises(ValueError):
        I.install("opencode")
    with pytest.raises(ValueError):
        I.install("nada")


def test_cli_version_help_and_doctor():
    run = lambda *a: subprocess.run([sys.executable, "-m", "apa7_mcp", *a], capture_output=True, text=True, encoding="utf-8", timeout=60)  # noqa: E731
    v = run("--version")
    assert v.returncode == 0 and v.stdout.strip().count(".") == 2
    assert "install" in run("--help").stdout
    d = run("doctor", "--offline")
    assert d.returncode == 0 and "Todo en orden" in d.stdout
    bad = run("install", "nada")
    assert bad.returncode == 2
