"""Servidor MCP (stdio): herramientas APA 7 para cualquier agente (Claude, opencode, Cursor, Cline, Codex…)."""
from __future__ import annotations

from typing import Any

from mcp.server.fastmcp import FastMCP

from . import doctor, pipeline, rules
from .errors import safe_tool

INSTRUCTIONS = """\
Herramientas APA 7 (7.ª ed.) para documentos Word (.docx). Para 'deja este Word en APA 7':
1) apa7_apply_apa(file_path) hace TODO el proceso (formato, títulos, sangrías, referencias, tablas, número de página) y valida.
2) Revise 'needs_agent_action': lo que no se corrige solo (citas sin referencia, referencias mal formadas, tablas/figuras sin rótulo).
   Resuélvalo con apa7_format_reference + apa7_replace_references, apa7_replace_text o apa7_insert_blocks. No invente fuentes: pregunte al usuario.
   'suggestions' son recomendaciones opcionales (números en letras, siglas): menciónelas, no las exija.
3) Repita apa7_validate_document hasta passed=true. Entregue la ruta de outputPath.
Revisión de redacción: apa7_extract_text entrega los párrafos; usted corrige ortografía/gramática con apa7_replace_text.
Fuentes: apa7_verify_references comprueba que las referencias existen (DOI, Crossref, Open Library) pero envía datos a servicios
externos: pida permiso al usuario y use confirm_external=true. 'doi_inexistente' o 'no_coincide' son señales fuertes de referencia inventada.
Portada, resumen y encabezado abreviado son opcionales en APA 7: no los agregue si el usuario no los pide.
Trabajo de estudiante (portada con estudiante, curso, docente): NO use running_head (solo número de página), sin 'Introducción' como título
(el texto inicia con el título del trabajo), sin tabla de contenido salvo que la institución la pida, siglas con paréntesis (corchetes solo dentro de un paréntesis),
sin 'Elaboración propia' en notas. Los niveles 4 y 5 son títulos en línea (el texto sigue en el mismo renglón tras el punto).
Para contenido nuevo (tablas, figuras, gráficos, listas, TOC, portada) use apa7_insert_blocks o apa7_build_document.
Nunca sobrescribe el original. Todas las herramientas devuelven ok=true/false; si ok=false lea error.code y error.hint (explican cómo
resolverlo: archivo inexistente, .doc antiguo, cifrado, control de cambios pendiente…). Si algo falla de forma rara, use apa7_doctor.
"""

mcp = FastMCP("apa7-tools", instructions=INSTRUCTIONS)

BLOCKS = (
    "Bloques (campo 'type'): title{text}; title_page{title,authors,affiliation,course,instructor,due_date} (authors acepta texto o lista de nombres; los demás, texto);"
    "heading{level 1-5,text}; paragraph{text,indent?}; block_quote{text}; list{style:'bullet'|'numbered',items[]}; "
    "table{title,header[],rows[][],note?}; figure{title, image_path | chart{type:'bar'|'line'|'pie'|'scatter'|'hist',labels[],"
    "series[{name,values[]}],x_label?,y_label?}, width_cm?, alt?, note?}; toc{levels?,title?}; list_of_tables; list_of_figures; "
    "references{entries[]}; page_break. En 'text' se admite **negrita** y *cursiva*."
)


@mcp.tool()
@safe_tool
def apa7_apply_apa(file_path: str, output_path: str | None = None, running_head: str | None = None,
                   toc: bool = False, page_numbers: bool = True, dry_run: bool = False, overwrite: bool = False,
                   accept_tracked_changes: bool = False, profile: str | dict | None = None,
                   language: str | None = None) -> dict:
    """TODO EN UNO: toma un .docx y lo deja en APA 7 (Times New Roman 12, doble espacio, márgenes 2.54 cm, Carta,
    sangrías, títulos, referencias con sangría francesa y orden alfabético, tablas APA, número de página) y lo valida.
    Devuelve cambios, puntaje antes/después, passed y 'needs_agent_action' con lo que debe resolver el agente.
    dry_run=true solo simula (no escribe). Si ya existe la salida se crea nombre-2.docx (overwrite=true la reemplaza).
    Si el documento tiene control de cambios pendiente responde TRACKED_CHANGES (accept_tracked_changes=true los acepta).
    profile: 'apa7' (Carta), 'a4', la ruta a un .json, o un objeto con font, size, page_size (auto|letter|a4), margin_cm, line_spacing,
    first_line_cm, numbered_headings, language. language='es'|'en' fija el idioma (por defecto se detecta del texto). Sin profile se usa apa7.profile.json junto al documento, si existe."""
    return pipeline.apply_apa(file_path, output_path, running_head=running_head, toc=toc, page_numbers=page_numbers,
                              dry_run=dry_run, overwrite=overwrite, accept_tracked_changes=accept_tracked_changes, profile=profile, language=language)


@mcp.tool()
@safe_tool
def apa7_validate_document(file_path: str, profile: str | dict | None = None, language: str | None = None) -> dict:
    """Valida un .docx contra APA 7: passed, puntaje 0-100, errores y advertencias restantes. profile: ver apa7_apply_apa."""
    return pipeline.validate_document(file_path, profile=profile, language=language)


@mcp.tool()
@safe_tool
def apa7_analyze_document(file_path: str, profile: str | dict | None = None, language: str | None = None) -> dict:
    """Informe completo: formato, citas, referencias, coherencia cita↔referencia, tablas/figuras, puntaje."""
    return pipeline.analyze_file(file_path, profile=profile, language=language)


@mcp.tool()
@safe_tool
def apa7_analyze_text(text: str, language: str | None = None) -> dict:
    """Analiza texto plano (citas y referencias) contra APA 7. language: es | en (por defecto se detecta)."""
    from . import analyze

    return analyze.analyze_text(text, lang=language if language in ("es", "en") else None)


@mcp.tool()
@safe_tool
def apa7_read_structure(file_path: str) -> dict:
    """Estructura del .docx: párrafos con índice y nivel de título, tablas, imágenes, TOC, números de página."""
    return pipeline.read_structure(file_path)


@mcp.tool(description="Crea un documento APA 7 NUEVO desde bloques. " + BLOCKS)
@safe_tool
def apa7_build_document(output_path: str, blocks: list[dict[str, Any]], page_numbers: bool = True,
                        running_head: str | None = None, overwrite: bool = False, profile: str | dict | None = None,
                        language: str | None = None) -> dict:
    return pipeline.build_document(output_path, blocks, page_numbers=page_numbers, running_head=running_head, overwrite=overwrite,
                                   profile=profile, language=language)


@mcp.tool(description="Inserta contenido (tablas, figuras, gráficos, listas, TOC, párrafos, referencias…) en un .docx existente. "
                      "where: end | start | before_references | after_paragraph (con paragraph_index) | after_text (con text). "
                      "Continúa la numeración de Tabla/Figura. " + BLOCKS)
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
    """Reemplaza texto: replacements=[{find, replace}]. Úselo para corregir citas o frases puntuales."""
    return pipeline.replace_text(file_path, output_path, replacements, overwrite, accept_tracked_changes)


@mcp.tool()
@safe_tool
def apa7_replace_references(file_path: str, entries: list[str], output_path: str | None = None,
                            overwrite: bool = False, accept_tracked_changes: bool = False,
                            profile: str | dict | None = None, language: str | None = None) -> dict:
    """Reemplaza TODA la lista de referencias por 'entries' (ordenadas, sangría francesa; use *cursiva* para títulos)."""
    return pipeline.replace_references(file_path, output_path, entries, overwrite, accept_tracked_changes, profile=profile, language=language)


@mcp.tool()
@safe_tool
def apa7_format_reference(type: str, title: str, authors: list[str] | None = None, initials: list[str] | None = None,
                          corporate_author: str | None = None, year: str | None = None, publisher: str | None = None,
                          journal: str | None = None, volume: str | None = None, number: str | None = None,
                          pages: str | None = None, url: str | None = None, edition: str | None = None,
                          site_name: str | None = None, book_title: str | None = None,
                          book_editors: list[str] | None = None, subtitle: str | None = None,
                          full_date: str | None = None, bracket: str | None = None, language: str = "es") -> dict:
    """Construye una referencia APA 7 (con *cursiva*). type: libro, capitulo_libro, articulo_revista, articulo_periodico,
    tesis, informe, pagina_web, pelicula, video, podcast, ley, red_social. authors = apellidos; initials = ['K. S.', …]."""
    return rules.format_reference({
        "type": type, "title": title, "authors": authors, "initials": initials, "corporateAuthor": corporate_author,
        "year": year, "publisher": publisher, "journal": journal, "volume": volume, "number": number, "pages": pages,
        "url": url, "edition": edition, "siteName": site_name, "bookTitle": book_title, "bookEditors": book_editors,
        "subtitle": subtitle, "fullDate": full_date, "bracket": bracket}, lang=language)


@mcp.tool()
@safe_tool
def apa7_validate_reference(reference: str, language: str | None = None) -> dict:
    """Valida una entrada de la lista de referencias. language: es | en (por defecto se detecta)."""
    from . import lang as L

    issues = rules.validate_reference(reference, language if language in ("es", "en") else L.detect([reference]))
    return {"valid": not any(i["severity"] == "error" for i in issues), "issues": issues}


@mcp.tool()
@safe_tool
def apa7_add_page_numbers(file_path: str, running_head: str | None = None, output_path: str | None = None,
                          overwrite: bool = False) -> dict:
    """Número de página arriba a la derecha (APA 7) en todas las secciones; con running_head agrega el encabezado abreviado."""
    return pipeline.add_page_numbers(file_path, output_path, running_head, overwrite)


@mcp.tool()
@safe_tool
def apa7_format_tables(file_path: str, output_path: str | None = None, overwrite: bool = False) -> dict:
    """Pasa las tablas existentes al estilo APA (solo líneas horizontales)."""
    return pipeline.format_tables(file_path, output_path, overwrite)


@mcp.tool()
@safe_tool
def apa7_extract_text(file_path: str, start: int = 0, limit: int | None = 80) -> dict:
    """Devuelve los párrafos del .docx con índice y sección (cuerpo, título, referencias, apéndice) para que usted revise
    ortografía, gramática y estilo académico. Corrija con apa7_replace_text. Use start/limit para documentos largos (next_start)."""
    return pipeline.extract_text(file_path, start, limit)


@mcp.tool()
@safe_tool
def apa7_verify_references(file_path: str | None = None, entries: list[str] | None = None,
                           confirm_external: bool = False) -> dict:
    """Comprueba que las fuentes EXISTEN: valida DOI (doi.org + Crossref comparando título/año/autor), busca artículos en Crossref y
    libros en Open Library, y revisa enlaces. Detecta referencias inventadas. Envía las referencias a servicios externos: pida permiso
    al usuario y llame con confirm_external=true. Use file_path (lee la sección Referencias) o entries."""
    return pipeline.verify_sources(file_path, entries, confirm_external)


@mcp.tool()
@safe_tool
def apa7_list_rules(category: str | None = None) -> dict:
    """Reglas APA 7 extraídas de las guías. category: formato|citacion|referencias|titulos|tablas_figuras|estructura|texto."""
    found = rules.list_rules(category)
    return {"count": len(found), "rules": found}


@mcp.tool()
@safe_tool
def apa7_doctor(check_network: bool = True) -> dict:
    """Diagnóstico del entorno (Python, dependencias, permisos de escritura, Word/LibreOffice, red). Úselo si algo falla."""
    return doctor.run(check_network)


def main() -> None:
    mcp.run()
