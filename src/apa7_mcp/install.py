"""Conecta el servidor a un agente:  apa7-tools install claude | claude-desktop | opencode | cursor  [--print]"""
from __future__ import annotations

import json
import os
import shutil
import subprocess
import sys
from pathlib import Path

NAME = "apa7-tools"
AGENTS = ("claude", "claude-desktop", "opencode", "cursor")
ENV = {"PYTHONUTF8": "1", "PYTHONIOENCODING": "utf-8"}


def command() -> tuple[str, list[str]]:
    """Cómo arrancar el servidor: el Python que ejecuta esta instalación (funciona con venv, pipx y uvx)."""
    return sys.executable, ["-m", "apa7_mcp"]


def target(agent: str, cwd: Path | None = None) -> tuple[Path, dict]:
    """(archivo de configuración, entrada a escribir) de cada agente."""
    py, args = command()
    home = Path.home()
    if agent in ("claude", "cursor"):
        entry = {"type": "stdio", "command": py, "args": args, "env": dict(ENV)}
        path = (cwd or Path.cwd()) / ".mcp.json" if agent == "claude" else home / ".cursor" / "mcp.json"
        return path, entry
    if agent == "claude-desktop":
        base = (Path(os.environ.get("APPDATA", home)) / "Claude" if os.name == "nt"
                else home / "Library" / "Application Support" / "Claude" if sys.platform == "darwin" else home / ".config" / "Claude")
        return base / "claude_desktop_config.json", {"command": py, "args": args, "env": dict(ENV)}
    if agent == "opencode":
        return home / ".config" / "opencode" / "opencode.json", {"type": "local", "command": [py, *args], "environment": dict(ENV), "enabled": True}
    raise ValueError(f"Agente no soportado: {agent}. Use: {', '.join(AGENTS)}")


def _write(path: Path, agent: str, entry: dict) -> None:
    cfg: dict = {}
    if path.exists():
        raw = path.read_text(encoding="utf-8")
        try:
            cfg = json.loads(raw) if raw.strip() else {}
        except json.JSONDecodeError as e:
            raise ValueError(f"{path} no es JSON válido ({e}). Añada la entrada manualmente (use --print).") from e
        try:
            shutil.copyfile(path, str(path) + ".bak")
        except OSError:
            pass  # la copia de seguridad es opcional
    else:
        path.parent.mkdir(parents=True, exist_ok=True)
    if agent == "opencode":
        cfg.setdefault("$schema", "https://opencode.ai/config.json")
        cfg.setdefault("mcp", {})[NAME] = entry
    else:
        cfg.setdefault("mcpServers", {})[NAME] = entry
    path.write_text(json.dumps(cfg, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")


def snippet(agent: str) -> str:
    path, entry = target(agent)
    key = "mcp" if agent == "opencode" else "mcpServers"
    return f"# {path}\n" + json.dumps({key: {NAME: entry}}, indent=2, ensure_ascii=False)


def install(agent: str, print_only: bool = False) -> str:
    if agent not in AGENTS:
        raise ValueError(f"Uso: apa7-tools install {' | '.join(AGENTS)} [--print]")
    if print_only:
        return snippet(agent)
    py, args = command()
    if agent == "claude":
        exe = shutil.which("claude")
        if exe:
            r = subprocess.run([exe, "mcp", "add", NAME, "--scope", "user", "--", py, *args], capture_output=True, text=True)
            if r.returncode == 0:
                return f'Claude Code: "{NAME}" registrado (scope user). Reinicie la sesión.'
    path, entry = target(agent)
    try:
        _write(path, agent, entry)
    except PermissionError:
        return (f"No se pudo escribir {path} (permiso denegado). Pegue esto a mano:\n" + snippet(agent))
    where = {"claude": " (aplica a sesiones abiertas en esta carpeta; apruebe el servidor al abrir)",
             "opencode": ". Reinicie opencode", "claude-desktop": ". Reinicie Claude Desktop", "cursor": ". Reinicie Cursor"}[agent]
    return f"{agent}: añadido a {path}{where}."
