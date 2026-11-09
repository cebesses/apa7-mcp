"""Bloques de contenido APA 7 → elementos de Word (títulos, listas, tablas, figuras, gráficos, TOC…)."""
from __future__ import annotations

import io
import re
from pathlib import Path
from typing import Any

from docx.document import Document as DocumentT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Cm, Inches, Pt
from docx.text.paragraph import Paragraph

from .charts import render_chart
from . import lang as L
from .scan import scan
from .wordfmt import (CENTER, TARGET, LEFT, REF_HEADING, format_heading,
                      format_reference, format_table, heading_level, set_rfonts, text_of)

Block = dict[str, Any]


# ------------------------------------------------------------------ texto con *cursiva* y **negrita**


def add_inline(p: Paragraph, text: str, bold: bool = False, italic: bool = False) -> None:
    for part in re.split(r"(\*\*[^*]+\*\*|\*[^*]+\*)", text):
        if not part:
            continue
        b, i, t = bold, italic, part
        if part.startswith("**") and part.endswith("**") and len(part) > 4:
            b, t = True, part[2:-2]
        elif part.startswith("*") and part.endswith("*") and len(part) > 2:
            i, t = True, part[1:-1]
        r = p.add_run(t)
        r.bold = True if b else None
        r.italic = True if i else None
        r.font.name = TARGET["font"]
        set_rfonts(r._r.get_or_add_rPr())
        r.font.size = Pt(TARGET["size"])


def _para(doc: DocumentT, text: str = "", *, align=LEFT, first: float = 0, left: float = 0, single=False,
          bold=False, italic=False, keep=False, style: str | None = None) -> Paragraph:
    p = doc.add_paragraph(style=style) if style else doc.add_paragraph()
    pf = p.paragraph_format
    p.alignment = align
    pf.first_line_indent = Inches(first) if first else Inches(0)
    pf.left_indent = Inches(left)
    pf.line_spacing = 1.0 if single else 2.0
    pf.space_before = pf.space_after = Pt(0)
    pf.keep_with_next = True if keep else None
    add_inline(p, text, bold=bold, italic=italic)
    return p


# ------------------------------------------------------------------ listas (numeración real de Word)


def new_list(doc: DocumentT, kind: str) -> int:
    el = doc.part.numbering_part.element
    abs_ids = [int(a.get(qn("w:abstractNumId"))) for a in el.findall(qn("w:abstractNum"))]
    num_ids = [int(n.get(qn("w:numId"))) for n in el.findall(qn("w:num"))]
    aid, nid = max(abs_ids, default=0) + 1, max(num_ids, default=0) + 1
    fmt, text = ("bullet", "•") if kind == "bullet" else ("decimal", "%1.")
    abstract = parse_xml(
        f'<w:abstractNum {nsdecls("w")} w:abstractNumId="{aid}"><w:multiLevelType w:val="hybridMultilevel"/>'
        f'<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="{fmt}"/><w:lvlText w:val="{text}"/><w:lvlJc w:val="left"/>'
        f'<w:pPr><w:ind w:left="1080" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>'
    )
    first = el.find(qn("w:num"))
    if first is not None:
        first.addprevious(abstract)
    else:
        el.append(abstract)
    el.append(parse_xml(f'<w:num {nsdecls("w")} w:numId="{nid}"><w:abstractNumId w:val="{aid}"/></w:num>'))
    return nid



def _caption(doc: DocumentT, label: str, n: int, title: str) -> Paragraph:
    """Rótulo APA: '**Tabla N**' (campo SEQ, negrita), salto de línea y título en *cursiva* — un solo párrafo."""
    p = doc.add_paragraph()
    pf = p.paragraph_format
    p.alignment = LEFT
    pf.line_spacing, pf.space_before, pf.space_after, pf.keep_with_next = 2.0, Pt(0), Pt(0), True
    pf.first_line_indent = pf.left_indent = Inches(0)

    def run(xml: str | None = None, text: str | None = None):
        r = p.add_run(text) if text is not None else p.add_run()
        r.bold = True
        r.font.name = TARGET["font"]
        set_rfonts(r._r.get_or_add_rPr())
        r.font.size = Pt(TARGET["size"])
        if xml:
            r._r.append(parse_xml(xml))
        return r

    run(text=f"{label} ")
    run(f'<w:fldChar {nsdecls("w")} w:fldCharType="begin"/>')
    run(f'<w:instrText {nsdecls("w")} xml:space="preserve"> SEQ {label} \\* ARABIC </w:instrText>')
    run(f'<w:fldChar {nsdecls("w")} w:fldCharType="separate"/>')
    run(text=str(n))
    run(f'<w:fldChar {nsdecls("w")} w:fldCharType="end"/>')
    p.add_run().add_break()
    add_inline(p, title, italic=True)
    return p


def _caption_entries(doc: DocumentT, label: str) -> list[str]:
    """'Tabla N. Título' para cada rótulo (un párrafo con salto de línea, o rótulo + título en párrafos separados)."""
    paras = doc.paragraphs
    out: list[str] = []
    for i, p in enumerate(paras):
        t = text_of(p)
        m = re.match(rf"^{label}\s+(\d+)\.?\s*(.*)$", t, re.I)
        if not m:
            continue
        title = m.group(2) or (text_of(paras[i + 1]) if i + 1 < len(paras) else "")
        out.append(f"{label} {m.group(1)}. {title}".strip())
    return out

# ------------------------------------------------------------------ renderizado


class State:
    def __init__(self, tables: int = 0, figures: int = 0):
        self.tables, self.figures = tables, figures


def count_existing(doc: DocumentT) -> State:
    def mx(pat: str) -> int:
        return max([int(m.group(1)) for p in doc.paragraphs if (m := re.match(pat, text_of(p), re.I))] or [0])

    return State(mx(r"^(?:Tabla|Table)\s+(\d+)"), mx(r"^(?:Figura|Figure)\s+(\d+)"))


def _note(doc: DocumentT, note: str) -> None:
    p = doc.add_paragraph()
    p.paragraph_format.line_spacing = 1.0
    p.paragraph_format.space_before = p.paragraph_format.space_after = Pt(0)
    add_inline(p, f"*{L.current().note}.* " + note)


def _table(doc: DocumentT, header: list, rows: list[list]) -> None:
    cols = max([len(header), *[len(r) for r in rows], 1])
    t = doc.add_table(rows=1 + len(rows), cols=cols)
    tblW = t._tbl.tblPr.find(qn("w:tblW"))
    if tblW is None:
        tblW = OxmlElement("w:tblW")
        t._tbl.tblPr.append(tblW)
    tblW.set(qn("w:type"), "pct")
    tblW.set(qn("w:w"), "5000")
    for ri, row in enumerate([header, *rows]):
        for ci in range(cols):
            val = str(row[ci]) if ci < len(row) else ""
            p = t.cell(ri, ci).paragraphs[0]
            p.paragraph_format.line_spacing = 1.0
            p.paragraph_format.space_before = p.paragraph_format.space_after = Pt(2)
            p.alignment = CENTER if (ri == 0 or ci > 0) else LEFT
            add_inline(p, val, bold=(ri == 0))
    format_table(t)


def render_blocks(doc: DocumentT, blocks: list[Block], st: State) -> None:
    """Agrega los bloques AL FINAL del documento (el llamador los reubica si hace falta)."""
    for b in blocks:
        t = b.get("type")
        if t == "title":
            _para(doc, b["text"], align=CENTER, bold=True, keep=True)
        elif t == "title_page":
            for _ in range(3):
                _para(doc, "", align=CENTER)
            _para(doc, b["title"], align=CENTER, bold=True)
            _para(doc, "", align=CENTER)
            for k in ("authors", "affiliation", "course", "instructor", "due_date"):
                v = b.get(k)
                if isinstance(v, (list, tuple)):  # authors puede venir como lista: «A, B y C» / «A, B, and C»
                    v = [str(x).strip() for x in v if str(x).strip()]
                    conj = L.current().narr_and
                    v = (v[0] if len(v) == 1 else f"{v[0]} {conj} {v[1]}" if len(v) == 2
                         else f"{', '.join(v[:-1])}{',' if conj == 'and' else ''} {conj} {v[-1]}") if v else ""
                if v:
                    _para(doc, str(v), align=CENTER)
            note = b.get("author_note")  # portada profesional: ORCID, cambios de afiliación, agradecimientos, contacto
            if note:
                _para(doc, "", align=CENTER)
                _para(doc, "Author Note" if L.current().code == "en" else "Nota del autor", align=CENTER, bold=True)
                for part in ([note] if isinstance(note, str) else note):
                    if str(part).strip():
                        _para(doc, str(part).strip(), first=0.5)
            doc.add_page_break()
        elif t == "heading":
            lvl = int(b["level"])
            text = b["text"] + ("." if lvl >= 4 and not b["text"].rstrip().endswith((".", ":")) else "")
            try:
                p = doc.add_paragraph(style=f"Heading {lvl}")
            except KeyError:
                p = doc.add_paragraph()
            add_inline(p, text)
            format_heading(p, lvl)
        elif t == "paragraph":
            _para(doc, b["text"], first=0 if b.get("indent") is False else 0.5)
        elif t == "block_quote":
            _para(doc, b["text"], left=0.5)
        elif t == "list":
            nid = new_list(doc, "decimal" if b.get("style") == "numbered" else "bullet")
            for item in b["items"]:
                p = _para(doc, item)
                ppr = p._p.get_or_add_pPr()
                npr = ppr.get_or_add_numPr()
                npr.get_or_add_ilvl().val = 0
                npr.get_or_add_numId().val = nid
                p.paragraph_format.first_line_indent = None
                p.paragraph_format.left_indent = None
        elif t == "table":
            st.tables += 1
            _caption(doc, L.current().table, st.tables, b["title"])
            _table(doc, b["header"], b["rows"])
            if b.get("note"):
                _note(doc, b["note"])
            _para(doc, "", single=True)
        elif t == "figure":
            st.figures += 1
            if b.get("chart"):
                data = io.BytesIO(render_chart(b["chart"]))
            elif b.get("image_path"):
                data = str(Path(b["image_path"]))
            else:
                raise ValueError("figure requiere image_path o chart.")
            _caption(doc, L.current().figure, st.figures, b["title"])
            shape = doc.add_picture(data, width=Cm(float(b.get("width_cm", 14))))
            shape._inline.docPr.set("descr", b.get("alt") or b["title"])
            p = doc.paragraphs[-1]
            p.alignment = CENTER
            p.paragraph_format.line_spacing = 1.0
            p.paragraph_format.keep_with_next = bool(b.get("note"))
            if b.get("note"):
                _note(doc, b["note"])
            _para(doc, "", single=True)
        elif t in ("toc", "list_of_tables", "list_of_figures"):
            _para(doc, f"@@{t}:{b.get('levels', 3)}:{b.get('title', '')}@@")
        elif t == "references":
            p = doc.add_paragraph()
            p.paragraph_format.page_break_before = True
            try:
                p.style = doc.styles["Heading 1"]
            except KeyError:
                pass
            add_inline(p, b.get("title") or L.current().ref_heading)
            format_heading(p, 1)
            for e in sorted(b["entries"], key=lambda s: re.sub(r"\W", "", s.lower())):
                q = doc.add_paragraph()
                add_inline(q, e)
                format_reference(q)
        elif t == "page_break":
            doc.add_page_break()
        else:
            raise ValueError(f"Tipo de bloque desconocido: {t}")


# ------------------------------------------------------------------ TOC e índices


def _entries(doc: DocumentT, kind: str, levels: int) -> tuple[list[tuple[int, str]], str]:
    if kind == "toc":
        ents = [(heading_level(p), text_of(p)) for p in doc.paragraphs
                if 1 <= heading_level(p) <= levels and text_of(p) and not text_of(p).startswith("@@")]
        return ents, f'TOC \\o "1-{levels}" \\h \\z \\u'
    label = L.current().table if kind == "list_of_tables" else L.current().figure
    return [(1, t) for t in _caption_entries(doc, label)], f'TOC \\h \\z \\c "{label}"'


def resolve_markers(doc: DocumentT) -> bool:
    found = False
    for marker in [p for p in doc.paragraphs if text_of(p).startswith("@@")]:
        m = re.match(r"@@(toc|list_of_tables|list_of_figures):(\d):(.*)@@", text_of(marker))
        if not m:
            continue
        found = True
        kind, levels, title = m.group(1), int(m.group(2)), m.group(3)
        lg = L.current()
        default = {"toc": lg.toc_title, "list_of_tables": lg.lot_title, "list_of_figures": lg.lof_title}[kind]
        entries, instr = _entries(doc, kind, levels)
        head = doc.add_paragraph()
        add_inline(head, title or default, bold=True)
        head.alignment = CENTER
        head.paragraph_format.line_spacing = 2.0
        new = [head._p]
        paras: list[Paragraph] = []
        for lvl, txt in (entries or [(1, "(Actualice los campos en Word: clic derecho → Actualizar campo)")]):
            q = doc.add_paragraph()
            q.paragraph_format.left_indent = Inches(0.25 * (lvl - 1))
            q.paragraph_format.line_spacing = 2.0
            q.paragraph_format.first_line_indent = Inches(0)
            add_inline(q, txt)
            paras.append(q)
        # el campo abarca todas las entradas: begin en la primera, end en la última
        first, last = paras[0], paras[-1]
        begin = [parse_xml(f'<w:r {nsdecls("w")}><w:fldChar w:fldCharType="begin" w:dirty="true"/></w:r>'),
                 parse_xml(f'<w:r {nsdecls("w")}><w:instrText xml:space="preserve"> {instr} </w:instrText></w:r>'),
                 parse_xml(f'<w:r {nsdecls("w")}><w:fldChar w:fldCharType="separate"/></w:r>')]
        anchor = first._p.pPr if first._p.pPr is not None else None
        for el in reversed(begin):
            if anchor is not None:
                anchor.addnext(el)
            else:
                first._p.insert(0, el)
        last._p.append(parse_xml(f'<w:r {nsdecls("w")}><w:fldChar w:fldCharType="end"/></w:r>'))
        new += [q._p for q in paras]
        for el in new:
            marker._p.addprevious(el)
        marker._p.getparent().remove(marker._p)
    return found


# ------------------------------------------------------------------ ubicación de lo insertado


def body_children(doc: DocumentT) -> list:
    return [e for e in doc.element.body if e.tag != qn("w:sectPr")]


def move_new(doc: DocumentT, before: list, where: str, paragraph_index: int | None, text: str | None) -> None:
    body = doc.element.body
    seen = {id(e) for e in before}
    new = [e for e in body_children(doc) if id(e) not in seen]
    if where == "end" or not new:
        return
    paras = [pi.p for pi in scan(doc)]  # mismos índices que apa7_read (sin TOC; incluye controles de contenido)
    if where == "start":
        for e in reversed(new):
            body.insert(0, e)
        return
    if where == "before_references":
        old = [p for p in paras if id(p._p) in seen]
        ref = [p for p in old if REF_HEADING.match(text_of(p))]
        if ref:
            for e in new:
                ref[-1]._p.addprevious(e)
        return
    old_paras = [p for p in paras if id(p._p) in seen]
    if where == "after_paragraph":
        if paragraph_index is None or not (0 <= paragraph_index < len(old_paras)):
            raise ValueError(f"paragraph_index fuera de rango (0-{len(old_paras) - 1}).")
        anchor = old_paras[paragraph_index]._p
    elif where == "after_text":
        needle = (text or "").lower()
        hit = next((p for p in old_paras if needle and needle in text_of(p).lower()), None)
        if hit is None:
            raise ValueError(f"No se encontró un párrafo con el texto: {text}")
        anchor = hit._p
    else:
        raise ValueError(f"Posición desconocida: {where}")
    for e in new:
        anchor.addnext(e)
        anchor = e
