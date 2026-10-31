"""El catálogo de reglas (apa7_list_rules) no puede quedarse atrás del código."""
import re
from pathlib import Path

from apa7_mcp import rules

SRC = Path(__file__).parent.parent / "src" / "apa7_mcp"


def emitted_ids() -> set[str]:
    ids = set()
    for f in SRC.glob("*.py"):
        ids |= set(re.findall(r'issue\(\s*"([a-z]+-\d{3})"', f.read_text(encoding="utf-8")))
    return ids


def test_every_emitted_rule_is_in_the_catalog():
    catalog = {r["id"] for r in rules.list_rules()}
    missing = sorted(emitted_ids() - catalog)
    assert not missing, f"reglas que el código emite pero faltan en data/rules.json: {missing}"


def test_catalog_entries_are_complete():
    for r in rules.list_rules():
        assert r["id"] and r["category"] and r["title"] and r["description"], r
    assert len({r["id"] for r in rules.list_rules()}) == len(rules.list_rules()), "ids duplicados"
