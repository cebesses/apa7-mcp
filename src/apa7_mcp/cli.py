"""Punto de entrada único:  apa7-tools            → servidor MCP (stdio)
                           apa7-tools install   → conecta un agente
                           apa7-tools doctor    → diagnóstico del entorno
                           apa7-tools --version
"""
from __future__ import annotations

import sys

HELP = """apa7-tools — herramientas APA 7 para agentes de IA (MCP)

  apa7-tools                         inicia el servidor MCP por stdio (lo hace el agente)
  apa7-tools install <agente> [--print]   claude | claude-desktop | opencode | cursor
  apa7-tools doctor [--offline]      diagnostica Python, dependencias, permisos, Word y red
  apa7-tools --version
"""


def main(argv: list[str] | None = None) -> None:
    argv = sys.argv[1:] if argv is None else argv
    cmd = argv[0] if argv else ""
    if cmd in ("-V", "--version", "version"):
        from . import __version__

        print(__version__)
    elif cmd in ("-h", "--help", "help"):
        print(HELP)
    elif cmd == "install":
        from .install import install

        agent = next((a for a in argv[1:] if not a.startswith("-")), "")
        try:
            print(install(agent, print_only="--print" in argv))
        except ValueError as e:
            print(e, file=sys.stderr)
            sys.exit(2)
    elif cmd == "doctor":
        from . import doctor

        report = doctor.run(check_network="--offline" not in argv)
        print(doctor.render(report))
        sys.exit(0 if report["ok"] else 1)
    else:
        from .server import main as serve

        serve()
