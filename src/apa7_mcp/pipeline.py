"""Operaciones de alto nivel que usan las herramientas MCP (entrada validada, salida verificada, original intacto)."""
from __future__ import annotations

import functools
import inspect
import os
import re
import tempfile
import time

from docx import Document

from . import lang as L
from . import options as O
from . import wordfmt as W
from .analyze import analyze_document
from .blocks import (Block, State, add_inline, body_children, count_existing, move_new, render_blocks, resolve_markers)
from .errors import ApaError, clean_path, resolve_input
from .integrity import guard, snapshot
from .safeio import resolve_output, save_verified
from .scan import accept_tracked, clear_cache, has_toc, scan, tracked_count

REQUIRED = {
    "title": ("text",), "title_page": ("title",), "heading": ("level", "text"), "paragraph": ("text",),
    "block_quote": ("text",), "list": ("items",), "table": ("title", "header", "rows"), "figure": ("title",),
    "references": ("entries",), "toc": (), "list_of_tables": (), "list_of_figures": (), "page_break": (),
}



def profiled(path_param: str | None = "input_path"):
    """Añade el parámetro `profile` a una operación y activa las opciones durante la llamada (se restauran al salir)."""

    def deco(fn):
        sig = inspect.signature(fn)

        @functools.wraps(fn)
        def wrapper(*args, profile=None, language=None, **kw):
            path = sig.bind_partial(*args, **kw).arguments.get(path_param) if path_param else None
            opts = O.load(profile, str(clean_path(path)) if path else None)
            if language:
                if language not in ("auto", "es", "en"):
                    raise ApaError("INVALID_ARGUMENT", f"language inválido: «{language}».", "Use auto, es o en.")
                opts = O.replace(opts, language=language)
            with O.use_options(opts):
                return fn(*args, **kw)

        return wrapper

    return deco


def validate_blocks(blocks: list) -> None:
    """Revisa la forma de los bloques antes de tocar nada, con mensajes que dicen cuál bloque y qué campo."""
    if not isinstance(blocks, list) or not blocks:
        raise ApaError("INVALID_ARGUMENT", "blocks debe ser una lista con al menos un bloque.", "Ejemplo: [{\"type\":\"paragraph\",\"text\":\"…\"}].")
    for i, b in enumerate(blocks):
        if not isinstance(b, dict) or "type" not in b:
            raise ApaError("INVALID_ARGUMENT", f"El bloque {i} no es un objeto con campo 'type'.", f"Tipos válidos: {', '.join(REQUIRED)}.")
        if b["type"] not in REQUIRED:
            raise ApaError("INVALID_ARGUMENT", f"Bloque {i}: tipo desconocido «{b['type']}».", f"Tipos válidos: {', '.join(REQUIRED)}.")
        missing = [f for f in REQUIRED[b["type"]] if f not in b or b[f] in (None, "")]
        if b["type"] == "figure" and not (b.get("chart") or b.get("image_path")):
            missing.append("chart o image_path")
        if b["type"] == "heading" and b.get("level") not in (1, 2, 3, 4, 5):
            raise ApaError("INVALID_ARGUMENT", f"Bloque {i}: heading.level debe ser 1 a 5.", "")
        if missing:
            raise ApaError("INVALID_ARGUMENT", f"Bloque {i} ({b['type']}): falta {', '.join(missing)}.", "Revise los campos requeridos del bloque.")
        if b["type"] == "figure" and b.get("image_path") and not os.path.exists(os.path.expanduser(str(b["image_path"]))):
            raise ApaError("FILE_NOT_FOUND", f"Bloque {i}: no existe la imagen {b['image_path']}.", "Compruebe la ruta de la imagen (PNG o JPG).")


def _open(path: str):
    inp = resolve_input(path)
    try:
        doc = Document(inp)
    except Exception as e:  # noqa: BLE001
        raise ApaError("CORRUPT_DOCX", f"No se pudo abrir el documento ({type(e).__name__}: {e}).",
                       "Abra el archivo en Word y guárdelo de nuevo como .docx.") from e
    clear_cache()
    L.resolve(pi.text for pi in scan(doc))  # idioma del documento (perfil o detección automática)
    return inp, doc


def _tracked(doc, accept: bool, changes: list[str] | None = None) -> None:
    n = tracked_count(doc)
    if not n:
        return
    if not accept:
        raise ApaError("TRACKED_CHANGES", f"El documento tiene {n} cambio(s) controlado(s) pendiente(s) (control de cambios).",
                       "Acéptelos en Word (Revisar → Aceptar todos los cambios) o repita con accept_tracked_changes=true "
                       "para aceptarlos automáticamente.")
    accept_tracked(doc)
    if changes is not None:
        changes.append(f"{n} cambio(s) controlado(s) aceptado(s).")


# ------------------------------------------------------------------ todo en uno


@profiled("input_path")
def apply_apa(input_path: str, output_path: str | None = None, *, running_head: str | None = None,
              toc: bool = False, page_numbers: bool = True, overwrite: bool = False, dry_run: bool = False,
              accept_tracked_changes: bool = False) -> dict:
    """Analiza → corrige formato/estructura → valida. Devuelve el informe y lo que debe resolver el agente."""
    t0 = time.perf_counter()
    inp = resolve_input(input_path)
    out = None if dry_run else resolve_output(inp, output_path, overwrite=overwrite)
    before = analyze_document(inp)
    _, doc = _open(inp)
    changes: list[str] = []
    _tracked(doc, accept_tracked_changes, changes)
    snap = snapshot(doc)

    tgt = W.choose_target(doc)
    if W.fix_page(doc):
        changes.append("Márgenes de 2.54 cm; página Carta (se conserva A4 si el documento ya lo era).")
    n = W.fix_styles(doc)
    changes.append(f"Estilos ({n}): {tgt['font']} {tgt['size']:g} pt, doble espacio, sin espacio entre párrafos, títulos APA 1-5.")
    if W.fix_settings(doc):
        changes.append("Separación automática con guiones desactivada.")
    promoted = W.promote_headings(doc)
    if promoted:
        changes.append(f"{promoted} título(s) sin estilo convertidos a Heading 1-5 (habilita la tabla de contenido).")
    bq = W.fix_block_quotes(doc)
    if bq:
        changes.append(f"{bq} cita(s) textual(es) de 40+ palabras convertida(s) a cita en bloque.")
    c = W.fix_paragraphs(doc)
    changes.append(f"Párrafos: {c['body']} con sangría de 1.27 cm, {c['headings']} títulos, {c['references']} referencias con sangría francesa"
                   + (f", {c['abstract']} de resumen/palabras clave" if c["abstract"] else "")
                   + (f", {c['appendix']} apéndice(s)" if c["appendix"] else "") + ".")
    nf = W.normalize_fonts(doc)
    if nf["fonts"] or nf["sizes"] or nf["colors"]:
        changes.append(f"Fuente/tamaño unificados ({nf['fonts']} fuentes, {nf['sizes']} tamaños)"
                       + (f"; {nf['colors']} color(es) de texto a negro." if nf["colors"] else "."))
    t = W.fix_text(doc)
    if t["fixed"]:
        changes.append(f"Texto: {t['fixed']} corrección(es) (DOI, \"Recuperado de\", \"&\"→\"y\", p./pp., et al., iniciales, rangos, "
                       "ciudad de editorial, orden de citas, espacios y puntuación).")
    ri = W.fix_reference_italics(doc)
    if ri:
        changes.append(f"{ri} referencia(s) con título/revista en cursiva.")
    st_n = W.fix_stat_italics(doc)
    if st_n:
        changes.append(f"{st_n} símbolo(s) estadístico(s) (p, M, SD, t, F, r, n…) en cursiva.")
    if W.sort_references(doc):
        changes.append("Lista de referencias en orden alfabético (letra por letra).")
    nt = W.format_tables(doc)
    if nt:
        changes.append(f"{nt} tabla(s) con estilo APA (solo líneas horizontales).")
    alt = W.fix_alt_text(doc)
    if alt:
        changes.append(f"Texto alternativo agregado a {alt} imagen(es).")
    if page_numbers and (running_head or not W.has_page_numbers(doc)):
        W.add_page_numbers(doc, running_head)
        changes.append("Número de página arriba a la derecha" + (" con encabezado abreviado." if running_head else "."))
    if toc:
        if has_toc(doc):
            changes.append("Ya existe una tabla de contenido: no se agregó otra.")
        else:
            before_children = body_children(doc)
            render_blocks(doc, [{"type": "toc"}], count_existing(doc))
            move_new(doc, before_children, "start", None, None)
            resolve_markers(doc)
            W.request_field_update(doc)
            changes.append("Tabla de contenido insertada (Word pedirá actualizar campos al abrir).")
    guard(snap, doc)  # si algo se perdiera, no se escribe nada

    if dry_run:
        with tempfile.TemporaryDirectory() as d:
            tmp = save_verified(doc, os.path.join(d, "dry.docx"))
            after = analyze_document(tmp)
    else:
        save_verified(doc, out)
        after = analyze_document(out)
    all_issues = after["citationIssues"] + after["referenceIssues"] + after["formatIssues"]
    pending = [i for i in all_issues if i["severity"] in ("error", "warning")]
    notes = [i for i in all_issues if i["severity"] == "info"]
    errors = [i for i in pending if i["severity"] == "error"]
    return {
        "outputPath": out,
        "dry_run": dry_run,
        "changes": changes,
        "before": {"score": before["score"], "issues": before["totalIssues"]},
        "after": {"score": after["score"], "issues": after["totalIssues"], "summary": after["summary"]},
        "passed": not errors and after["score"] >= 80,
        "needs_agent_action": pending,
        "suggestions": notes,
        "elapsed_s": round(time.perf_counter() - t0, 2),
        "how_to_resolve": (
            "Para cada pendiente: citas sin referencia → pedir al usuario la fuente o corregir la cita con apa7_replace_text; "
            "referencias con errores → reescribir con apa7_format_reference y reemplazar la lista con apa7_replace_references; "
            "tablas/figuras sin rótulo → apa7_insert_blocks; citas textuales sin página → pedir la página al usuario. "
            "'suggestions' son recomendaciones opcionales (números en letras, siglas). Luego ejecutar apa7_validate_document."
        ),
    }


# ------------------------------------------------------------------ construir e insertar


@profiled(None)
def build_document(output_path: str, blocks: list[Block], *, page_numbers: bool = True,
                   running_head: str | None = None, overwrite: bool = False) -> dict:
    validate_blocks(blocks)
    out = resolve_output(None, output_path, overwrite=overwrite)
    clear_cache()
    L.resolve(str(v) for b in blocks for k, v in b.items() if k in ("text", "title") and isinstance(v, str))
    doc = Document()
    W.fix_page(doc)
    W.fix_styles(doc)
    st = State()
    render_blocks(doc, blocks, st)
    # El Document() base puede traer un párrafo vacío inicial: se elimina si está vacío al inicio.
    first = doc.paragraphs[0] if doc.paragraphs else None
    if first is not None and not first.text.strip() and len(doc.paragraphs) > 1 and not first._p.xpath(".//w:br"):
        first._p.getparent().remove(first._p)
    if resolve_markers(doc):
        W.request_field_update(doc)
    if page_numbers:
        W.add_page_numbers(doc, running_head)
    save_verified(doc, out)
    return {"outputPath": out, "blocks": len(blocks), "tables": st.tables, "figures": st.figures}


@profiled("input_path")
def insert_blocks(input_path: str, output_path: str | None, blocks: list[Block], where: str = "end",
                  paragraph_index: int | None = None, text: str | None = None, overwrite: bool = False,
                  accept_tracked_changes: bool = False) -> dict:
    validate_blocks(blocks)
    if where not in ("end", "start", "before_references", "after_paragraph", "after_text"):
        raise ApaError("INVALID_ARGUMENT", f"where inválido: «{where}».", "Use end, start, before_references, after_paragraph o after_text.")
    inp, doc = _open(input_path)
    out = resolve_output(inp, output_path, overwrite=overwrite)
    _tracked(doc, accept_tracked_changes)
    snap = snapshot(doc)
    W.choose_target(doc)  # el contenido insertado usa la fuente del documento
    st = count_existing(doc)
    before = body_children(doc)
    render_blocks(doc, blocks, st)
    move_new(doc, before, where, paragraph_index, text)
    if resolve_markers(doc):
        W.request_field_update(doc)
    guard(snap, doc)
    save_verified(doc, out)
    return {"outputPath": out, "inserted": len(blocks), "tables_total": st.tables, "figures_total": st.figures}


def read_structure(path: str) -> dict:
    _, doc = _open(path)
    paras = [{"index": pi.i, "text": (pi.text[:80] + "…") if len(pi.text) > 80 else pi.text, "heading_level": pi.level}
             for pi in scan(doc)]
    return {
        "paragraphs": [p for p in paras if p["text"] or p["heading_level"]],
        "tables": sum(1 for _ in doc.element.body.iter(W.W + "tbl")),
        "images": len(doc.inline_shapes),
        "has_page_numbers": W.has_page_numbers(doc),
        "has_toc": has_toc(doc),
        "tracked_changes": tracked_count(doc),
    }


# ------------------------------------------------------------------ edición de contenido


def replace_text(input_path: str, output_path: str | None, replacements: list[dict], overwrite: bool = False,
                 accept_tracked_changes: bool = False) -> dict:
    if not isinstance(replacements, list) or not replacements or not all(isinstance(r, dict) and r.get("find") and "replace" in r for r in replacements):
        raise ApaError("INVALID_ARGUMENT", "replacements debe ser una lista de {find, replace} con 'find' no vacío.",
                       "Ejemplo: [{\"find\": \"(Pérez, 2019)\", \"replace\": \"(Pérez, 2020)\"}].")
    inp, doc = _open(input_path)
    out = resolve_output(inp, output_path, overwrite=overwrite)
    _tracked(doc, accept_tracked_changes)
    snap = snapshot(doc)
    applied = 0
    hit: set[str] = set()
    for p in W.iter_all_paragraphs(doc):
        for r in replacements:
            pat = re.compile(re.escape(r["find"]))
            if W._replace_in_paragraph(p, pat, lambda _m, new=r["replace"]: new):
                applied += 1
                hit.add(r["find"])
    guard(snap, doc, allow_word_loss=0.25)
    save_verified(doc, out)
    return {"outputPath": out, "applied": applied, "not_found": [r["find"] for r in replacements if r["find"] not in hit]}


@profiled("input_path")
def replace_references(input_path: str, output_path: str | None, entries: list[str], overwrite: bool = False,
                       accept_tracked_changes: bool = False) -> dict:
    """Reemplaza la lista de referencias: ordenada, con sangría francesa y *cursiva*."""
    if not isinstance(entries, list) or not entries or not all(isinstance(e, str) and e.strip() for e in entries):
        raise ApaError("INVALID_ARGUMENT", "entries debe ser una lista de textos de referencia no vacíos.", "")
    inp, doc = _open(input_path)
    out = resolve_output(inp, output_path, overwrite=overwrite)
    _tracked(doc, accept_tracked_changes)
    snap = snapshot(doc)
    W.choose_target(doc)
    infos = scan(doc)
    rr = W.reference_range(infos)
    ordered = sorted(entries, key=W._ref_key)
    removed = 0
    if rr is None:
        h = doc.add_paragraph()
        h.paragraph_format.page_break_before = True
        try:
            h.style = doc.styles["Heading 1"]
        except KeyError:
            pass
        add_inline(h, L.current().ref_heading)
        W.format_heading(h, 1)
        anchor = h._p
    else:
        for pi in infos[rr[0] + 1: rr[1]]:
            pi.el.getparent().remove(pi.el)
            removed += 1
        anchor = infos[rr[0]].el
    for e in ordered:
        q = doc.add_paragraph()
        add_inline(q, e)
        W.format_reference(q)
        anchor.addnext(q._p)
        anchor = q._p
    guard(snap, doc, allow_word_loss=1.0)  # la lista anterior se sustituye a propósito
    save_verified(doc, out)
    return {"outputPath": out, "removed": removed, "written": len(ordered)}


@profiled("input_path")
def add_page_numbers(input_path: str, output_path: str | None, running_head: str | None, overwrite: bool = False) -> dict:
    inp, doc = _open(input_path)
    out = resolve_output(inp, output_path, overwrite=overwrite)
    snap = snapshot(doc)
    n = W.add_page_numbers(doc, running_head)
    guard(snap, doc)
    save_verified(doc, out)
    return {"outputPath": out, "sections": n}


def format_tables(input_path: str, output_path: str | None, overwrite: bool = False) -> dict:
    inp, doc = _open(input_path)
    out = resolve_output(inp, output_path, overwrite=overwrite)
    snap = snapshot(doc)
    n = W.format_tables(doc)
    guard(snap, doc)
    save_verified(doc, out)
    return {"outputPath": out, "tablesFormatted": n}


@profiled("path")
def analyze_file(path: str) -> dict:
    return analyze_document(resolve_input(path))


@profiled("path")
def validate_document(path: str) -> dict:
    r = analyze_document(resolve_input(path))
    allv = r["citationIssues"] + r["referenceIssues"] + r["formatIssues"]
    errors = [i for i in allv if i["severity"] == "error"]
    warnings = [i for i in allv if i["severity"] == "warning"]
    return {"passed": not errors and r["score"] >= 80, "score": r["score"], "errors": len(errors),
            "warnings": len(warnings), "summary": r["summary"], "remaining": errors + warnings}


# ------------------------------------------------------------------ revisión de texto y verificación de fuentes


def extract_text(path: str, start: int = 0, limit: int | None = None) -> dict:
    """Párrafos con índice y sección, para que el agente revise ortografía, gramática y estilo académico."""
    if start < 0 or (limit is not None and limit < 1):
        raise ApaError("INVALID_ARGUMENT", "start debe ser ≥ 0 y limit ≥ 1.", "")
    _, doc = _open(path)
    infos = scan(doc)
    rr = W.reference_range(infos)
    items = []
    for pi in infos:
        if not pi.text:
            continue
        if rr and pi.i == rr[0]:
            sec = "título referencias"
        elif rr and rr[0] < pi.i < rr[1]:
            sec = "referencias"
        elif rr and pi.i >= rr[1]:
            sec = "apéndice"
        else:
            sec = "título" if pi.level else "cuerpo"
        items.append({"index": pi.i, "section": sec, "text": pi.text})
    end = len(items) if limit is None else start + limit
    page = items[start:end]
    return {
        "total": len(items), "returned": len(page), "next_start": end if end < len(items) else None,
        "paragraphs": page,
        "how_to_use": ("Revise ortografía, gramática, concordancia, registro académico y claridad. Corrija con apa7_replace_text "
                       "(find/replace con fragmentos exactos). No cambie citas ni datos. La puntuación y los espacios mecánicos ya los corrige apa7_apply_apa."),
    }


def verify_sources(file_path: str | None, entries: list[str] | None, confirm_external: bool) -> dict:
    """Comprueba que las referencias existen (Crossref, doi.org, Open Library). Requiere confirmación expresa."""
    if not confirm_external:
        return {"verificado": False,
                "mensaje": ("Esta herramienta envía el texto de las referencias y sus DOI a servicios externos (api.crossref.org, doi.org, "
                            "openlibrary.org). Pregunte al usuario si lo autoriza y vuelva a llamar con confirm_external=true.")}
    from .verify import verify_references

    if not entries:
        if not file_path:
            raise ApaError("INVALID_ARGUMENT", "Indique file_path o entries.", "")
        _, doc = _open(file_path)
        infos = scan(doc)
        rr = W.reference_range(infos)
        entries = [pi.text for pi in infos[rr[0] + 1: rr[1]] if pi.text] if rr else []
        if not entries:
            return {"verificado": False, "mensaje": "El documento no tiene sección de referencias."}
    return {"verificado": True, **verify_references(entries)}
