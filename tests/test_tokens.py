"""Presupuesto de tokens: lo que el agente lee en cada sesión y en cada respuesta debe seguir siendo pequeño."""
import asyncio
import json

import corpus
from apa7_mcp import rules, server


def _call(name, args):
    return asyncio.run(server.mcp.call_tool(name, args))[0].text


def test_tool_schemas_stay_small():
    tools = asyncio.run(server.mcp.list_tools())
    size = sum(len(json.dumps(t.model_dump(exclude_none=True), ensure_ascii=False)) for t in tools)
    assert len(tools) <= 10 and size <= 7500, (len(tools), size)
    assert len(server.INSTRUCTIONS) <= 1000


def test_responses_are_compact(tmp_path):
    f = str(corpus.big_doc(tmp_path / "b.docx", 300))
    assert len(_call("apa7_check", {"file_path": f})) <= 3000
    assert len(_call("apa7_apply_apa", {"file_path": f})) <= 3000
    assert len(_call("apa7_read", {"file_path": f})) <= 3000
    assert len(_call("apa7_help", {})) <= 4000
    assert "\n  " not in _call("apa7_help", {"topic": "blocks"})  # JSON sin sangría


def test_new_rules(tmp_path):
    many = ", ".join(f"Autor{chr(65 + i % 26)}{i}, A." for i in range(25))
    assert "ref-020" in {i["ruleId"] for i in rules.validate_reference(many + " (2020). Título. Revista, 1(2), 3–4.", "es")}
    r = server.apa7_apply_apa(file_path=str(corpus.plain_bad(tmp_path / "rh.docx")),
                              running_head="X" * 60, dry_run=True)
    assert r["ok"] is False and r["error"]["code"] == "INVALID_ARGUMENT"
