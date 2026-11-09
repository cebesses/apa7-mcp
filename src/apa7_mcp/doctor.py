"""Diagnóstico del entorno: python -m apa7_mcp doctor  (o apa7_help('doctor'))."""
from __future__ import annotations

import importlib.metadata as md
import os
import platform
import shutil
import sys
import tempfile
import urllib.request


def _ver(pkg: str) -> str | None:
    try:
        return md.version(pkg)
    except md.PackageNotFoundError:
        return None


def _writable(d: str) -> bool:
    try:
        with tempfile.NamedTemporaryFile(dir=d, delete=True):
            return True
    except OSError:
        return False


def _word_available() -> bool:
    if platform.system() != "Windows":
        return False
    for base in (os.environ.get("ProgramFiles", r"C:\Program Files"), os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")):
        for ver in ("Office16", "Office15", "Office14"):
            if os.path.exists(os.path.join(base, "Microsoft Office", "root", ver, "WINWORD.EXE")) or \
               os.path.exists(os.path.join(base, "Microsoft Office", ver, "WINWORD.EXE")):
                return True
    return False


def _net(url: str = "https://api.crossref.org/works?rows=0", timeout: float = 4.0) -> bool:
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "apa7-tools-doctor"})
        with urllib.request.urlopen(req, timeout=timeout) as r:
            return r.status == 200
    except Exception:  # noqa: BLE001
        return False


def run(check_network: bool = True) -> dict:
    checks: list[dict] = []

    def add(name: str, ok: bool, detail: str = "", hint: str = "", required: bool = True) -> None:
        checks.append({"check": name, "ok": ok, "detail": detail, "hint": hint, "required": required})

    py = sys.version_info
    add("Python ≥ 3.10", py >= (3, 10), platform.python_version(), "Instale Python 3.10 o superior.")
    for pkg, need, hint in (("mcp", True, "pip install \"mcp>=1.2,<2\""), ("python-docx", True, "pip install python-docx"),
                            ("lxml", True, "pip install lxml"), ("matplotlib", False, "pip install matplotlib (solo para gráficos)")):
        v = _ver(pkg)
        add(f"Paquete {pkg}", v is not None, v or "no instalado", hint, required=need)
    try:
        import mcp.server.fastmcp  # noqa: F401
        add("API FastMCP (mcp 1.x)", True, _ver("mcp") or "")
    except Exception as e:  # noqa: BLE001
        add("API FastMCP (mcp 1.x)", False, str(e)[:120], "Use mcp>=1.2,<2.")
    add("Escritura en carpeta temporal", _writable(tempfile.gettempdir()), tempfile.gettempdir(),
        "El Python debe poder escribir en TEMP (si corre restringido, use el Python global).")
    add("Escritura en la carpeta actual", _writable(os.getcwd()), os.getcwd(), "Solo importa si guarda salidas aquí.", required=False)
    add("Word instalado (verificación visual)", _word_available(), "", "Opcional: solo para pruebas visuales.", required=False)
    add("LibreOffice instalado", shutil.which("soffice") is not None, "", "Opcional.", required=False)
    if check_network:
        add("Red (Crossref)", _net(), "", "Solo necesaria para apa7_verify_references.", required=False)
    problems = [c for c in checks if c["required"] and not c["ok"]]
    return {"ok": not problems, "python": sys.executable, "version": md.version("apa7-tools") if _ver("apa7-tools") else "dev",
            "checks": checks, "problemas": [f"{c['check']}: {c['hint']}" for c in problems]}


def render(report: dict) -> str:
    lines = [f"apa7-tools {report['version']}  |  {report['python']}"]
    for c in report["checks"]:
        mark = "OK " if c["ok"] else ("FALLA" if c["required"] else "aviso")
        lines.append(f"  [{mark}] {c['check']} {('— ' + c['detail']) if c['detail'] else ''}" + ("" if c["ok"] or not c["hint"] else f"   → {c['hint']}"))
    lines.append("Todo en orden." if report["ok"] else "Hay problemas que resolver (ver líneas FALLA).")
    return "\n".join(lines)
