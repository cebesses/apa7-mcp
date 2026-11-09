"""Servidor MCP (stdio): herramientas APA 7 para cualquier agente (Claude, opencode, Cursor, Cline, Codex…).
Pensado para gastar pocos tokens: pocas herramientas, descripciones cortas y respuestas resumidas (detail="full" da todo)."""
from __future__ import annotations

from collections import Counter
from typing import Any

from mcp.server.fastmcp import FastMCP

from . import doctor, pipeline, rules
from .analyze import group_issues
from .errors import safe_tool

INSTRUCTIONS = """\
APA 7 para Word (.docx). Flujo: 1) apa7_apply_apa(file_path) corrige y valida todo en una llamada; no llame a apa7_check
después salvo que haya editado a mano. 2) Resuelva 'needs_agent_action' con apa7_format_reference + apa7_replace_references,
apa7_replace_text o apa7_insert_blocks; no invente fuentes, pregunte. 'suggestions' son opcionales. Entregue outputPath.
Redacción: apa7_read(mode='text') y corrija con apa7_replace_text. apa7_verify_references envía datos a servicios externos:
pida permiso y use confirm_external=true. Trabajo de estudiante: sin running_head, sin título 'Introducción', sin TOC salvo
que la institución la pida. No agregue portada ni resumen si no los piden. Si ok=false lea error.hint; si falla raro, apa7_help('doctor')."""

mcp = FastMCP("apa7-tools", instructions=INSTRUCTIONS)


def _compact_json() -> None:
    """FastMCP serializa con indent=2; sin sangría la misma respuesta usa ~30 % menos tokens."""
    import json

    from mcp.server.fastmcp.utilities import func_metadata as F
    from mcp.types import TextContent

    original = F._convert_to_content

    def convert(result):
        if isinstance(result, (dict, list)):
            return [TextContent(type="text", text=json.dumps(result, ensure_ascii=False, separators=(",", ":"), default=str))]
        return original(result)

    F._convert_to_content = convert


_compact_json()

BLOCKS = (
    "Bloques (type): title{text}; title_page{title,authors(texto|lista),affiliation,course,instructor,due_date,author_note(texto|lista; trabajo profesional)}; heading{level 1-5,text}; "
    "paragraph{text,indent?}; block_quote{text}; list{style:bullet|numbered,items[]}; table{title,header[],rows[][],note?}; "
    "figure{title,image_path|chart{type:bar|line|pie|scatter|hist,labels[],series[{name,values[]}],x_label?,y_label?},width_cm?,alt?,note?}; "
    "toc{levels?,title?}; list_of_tables; list_of_figures; references{entries[]}; page_break. En text: **negrita**, *cursiva*."
)
PROFILE = ("profile: 'apa7'|'a4'|ruta .json|{font,size,page_size,margin_cm,line_spacing,first_line_cm,numbered_headings}. "
           "language: es|en (auto por defecto).")


def _prune(o):
    """Quita None, '' y colecciones vacías: menos tokens sin perder información."""
    if isinstance(o, dict):
        return {k: v for k, v in ((k, _prune(v)) for k, v in o.items()) if v not in (None, "", [], {})}
    if isinstance(o, list):
        return [_prune(v) for v in o]
    return o


def _summary_of(r: dict) -> dict:
    allv = r.get("citationIssues", []) + r.get("referenceIssues", []) + r.get("formatIssues", [])
    sev = Counter(i["severity"] for i in allv)
    return {"passed": not sev["error"] and r["score"] >= 80, "score": r["score"], "errors": sev["error"], "warnings": sev["warning"],
            "summary": r["summary"], "citations": len(r.get("citations", [])), "references": len(r.get("references", [])),
            "issues": group_issues(allv)}


@mcp.tool(description=(
    "TODO EN UNO: deja un .docx en APA 7 (fuente, doble espacio, márgenes, sangrías, títulos, referencias, tablas, número de página) "
    "y lo valida. Devuelve changes, score antes/después, passed y needs_agent_action. No toca el original (crea *.apa7.docx). "
    "dry_run simula. accept_tracked_changes acepta control de cambios. running_head solo en trabajos profesionales. " + PROFILE))
@safe_tool
def apa7_apply_apa(file_path: str, output_path: str | None = None, running_head: str | None = None,
                   toc: bool = False, page_numbers: bool = True, dry_run: bool = False, overwrite: bool = False,
                   accept_tracked_changes: bool = False, profile: str | dict | None = None,
                   language: str | None = None) -> dict:
    return pipeline.apply_apa(file_path, output_path, running_head=running_head, toc=toc, page_numbers=page_numbers,
                              dry_run=dry_run, overwrite=overwrite, accept_tracked_changes=accept_tracked_changes, profile=profile, language=language)


@mcp.tool()
@safe_tool
def apa7_check(file_path: str | None = None, text: str | None = None, reference: str | None = None,
               detail: str = "summary", profile: str | dict | None = None, language: str | None = None) -> dict:
    """Valida sin modificar: un .docx (file_path), texto plano (text) o una referencia (reference).
    Resumen agrupado por regla; detail='full' devuelve cada aviso, citas y referencias."""
    lang = language if language in ("es", "en") else None
    if reference:
        from . import lang as L

        issues = rules.validate_reference(reference, lang or L.detect([reference]))
        return _prune({"valid": not any(i["severity"] == "error" for i in issues),
                       "issues": issues if detail == "full" else group_issues(issues)})
    if text:
        from . import analyze

        r = analyze.analyze_text(text, lang=lang)
    elif file_path:
        r = pipeline.analyze_file(file_path, profile=profile, language=language)
    else:
        from .errors import ApaError

        raise ApaError("INVALID_ARGUMENT", "Indique file_path, text o reference.", "Ejemplo: apa7_check(file_path='trabajo.docx').")
    if detail == "full":
        s = _summary_of(r)
        s.pop("issues")
        return {**r, **s}
    s = _summary_of(r)
    for k in ("paragraphCount", "tables", "images", "measured"):
        if k in r:
            s[k] = r[k]
    return _prune(s)


@mcp.tool()
@safe_tool
def apa7_read(file_path: str, mode: str = "outline", start: int = 0, limit: int = 60) -> dict:
    """Lee un .docx. mode='outline': títulos, rótulos y conteos (índices para insertar). mode='text': párrafos con índice y sección
    para revisar ortografía/gramática/estilo (pagine con start/limit y next_start); corrija con apa7_replace_text sin tocar citas."""
    if mode == "text":
        r = pipeline.extract_text(file_path, start, limit)
        r.pop("how_to_use", None)
        return _prune(r)
    s = pipeline.read_structure(file_path)
    paras = s.pop("paragraphs")
    s["paragraphs"] = len(paras)
    s["outline"] = [{"i": p["index"], "lvl": p["heading_level"], "text": p["text"][:70]} for p in paras
                    if p["heading_level"] or rules_caption(p["text"])][start:start + max(limit, 1) * 3]
    return _prune(s)


def rules_caption(t: str) -> bool:
    from .wordfmt import CAPTION, REF_HEADING

    return bool(CAPTION.match(t) or REF_HEADING.match(t))


@mcp.tool(description="Crea un .docx APA 7 NUEVO desde bloques (los mismos de apa7_insert_blocks). " + PROFILE)
@safe_tool
def apa7_build_document(output_path: str, blocks: list[dict[str, Any]], page_numbers: bool = True,
                        running_head: str | None = None, overwrite: bool = False, profile: str | dict | None = None,
                        language: str | None = None) -> dict:
    return pipeline.build_document(output_path, blocks, page_numbers=page_numbers, running_head=running_head, overwrite=overwrite,
                                   profile=profile, language=language)


@mcp.tool(description="Inserta contenido en un .docx existente. where: end|start|before_references|after_paragraph (paragraph_index "
                      "de apa7_read)|after_text (text). Numera Tabla/Figura a continuación. " + BLOCKS)
@safe_tool
def apa7_insert_blocks(file_path: str, blocks: list[dict[str, Any]], where: str = "end",
                       paragraph_index: int | None = None, text: str | None = None,
                       output_path: str | None = None, overwrite: bool = False,
                       accept_tracked_changes: bool = False, profile: str | dict | None = None,
                       language: str | None = None) -> dict:
    return pipeline.insert_blocks(file_path, output_path, blocks, where, paragraph_index, text, overwrite, accept_tracked_changes, profile=profile, language=language)


@mcp.tool()
@safe_tool
def apa7_replace_text(file_path: str, replacements: list[dict[str, str]], output_path: str | None = None,
                      overwrite: bool = False, accept_tracked_changes: bool = False) -> dict:
    """Reemplaza fragmentos exactos: replacements=[{find, replace}] (correcciones de redacción o de citas)."""
    return pipeline.replace_text(file_path, output_path, replacements, overwrite, accept_tracked_changes)


@mcp.tool()
@safe_tool
def apa7_replace_references(file_path: str, entries: list[str], output_path: str | None = None,
                            overwrite: bool = False, accept_tracked_changes: bool = False,
                            profile: str | dict | None = None, language: str | None = None) -> dict:
    """Reemplaza TODA la lista de referencias por entries (se ordenan; *cursiva* para títulos/revistas)."""
    return pipeline.replace_references(file_path, output_path, entries, overwrite, accept_tracked_changes, profile=profile, language=language)


@mcp.tool()
@safe_tool
def apa7_format_reference(ref: dict[str, Any], language: str = "es") -> dict:
    """Arma una referencia APA 7 (con *cursiva*). ref: {type, title, authors[apellidos], initials['K. S.',…], corporate_author, year,
    publisher, journal, volume, number, pages, url (o DOI), edition, site_name, book_title, book_editors[], subtitle, full_date, bracket}.
    type: libro|capitulo_libro|articulo_revista|articulo_periodico|tesis|informe|pagina_web|pelicula|video|podcast|ley|red_social."""
    keys = {"corporate_author": "corporateAuthor", "site_name": "siteName", "book_title": "bookTitle",
            "book_editors": "bookEditors", "full_date": "fullDate"}
    return rules.format_reference({keys.get(k, k): v for k, v in (ref or {}).items()}, lang=language)


@mcp.tool()
@safe_tool
def apa7_verify_references(file_path: str | None = None, entries: list[str] | None = None,
                           confirm_external: bool = False) -> dict:
    """Comprueba que las fuentes existen (DOI, Crossref, Open Library, enlaces); detecta referencias inventadas.
    Envía datos a servicios externos: requiere permiso del usuario y confirm_external=true."""
    return pipeline.verify_sources(file_path, entries, confirm_external)


@mcp.tool()
@safe_tool
def apa7_help(topic: str = "rules", category: str | None = None, rule_id: str | None = None, network: bool = True) -> dict:
    """Ayuda bajo demanda. topic: rules (id y título; category o rule_id para el detalle) | blocks (tipos de bloque) |
    doctor (diagnóstico del entorno si algo falla)."""
    if topic == "blocks":
        return {"blocks": BLOCKS}
    if topic == "doctor":
        return doctor.run(network)
    found = rules.list_rules(category)
    if rule_id:
        return {"rules": [r for r in found if r["id"] == rule_id]}
    if category:
        return {"count": len(found), "rules": found}
    return {"count": len(found), "rules": [f"{r['id']} {r['title']}" for r in found]}


def _slim(schema):
    """Esquemas JSON más cortos (lo que el agente lee en cada sesión): sin 'title' y 'X | None' como X."""
    if isinstance(schema, dict):
        any_of = schema.get("anyOf")
        if isinstance(any_of, list) and len(any_of) == 2 and {"type": "null"} in any_of:
            rest = {k: v for k, v in schema.items() if k != "anyOf"}
            schema = {**next(a for a in any_of if a != {"type": "null"}), **rest}
            if schema.get("default") is None:
                schema.pop("default", None)
        return {k: _slim(v) for k, v in schema.items() if not (k == "title" and isinstance(v, str))}
    if isinstance(schema, list):
        return [_slim(v) for v in schema]
    return schema


for _t in mcp._tool_manager.list_tools():
    _t.parameters = _slim(_t.parameters)


def main() -> None:
    mcp.run()
