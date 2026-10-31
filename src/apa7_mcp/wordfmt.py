"""Formato APA 7 sobre documentos Word (python-docx + OOXML directo donde es más rápido o la librería no llega)."""
from __future__ import annotations

import re
from collections import Counter
from typing import Iterable

from docx.document import Document as DocumentT
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Cm, Inches, Pt, RGBColor
from docx.text.paragraph import Paragraph

from copy import deepcopy

from lxml import etree

from docx.text.run import Run

from . import lang as L
from . import options as O
from . import rules, textrules
from .options import TARGET
from .scan import (P, PInfo, T, W, clear_cache, eff_align, eff_size, raw_el, scan, styles_of, text_el)

# fuentes aceptadas por APA 7 → tamaño (pt). Aptos 11 se admite si la institución lo permite.
ACCEPTED_FONTS = {"times new roman": 12, "georgia": 11, "computer modern": 10, "calibri": 11, "arial": 11,
                  "lucida sans unicode": 10, "aptos": 11}
FONT = "Times New Roman"  # para documentos nuevos

REF_HEADING = re.compile(r"^(?:referencias?(?:\s+bibliogr[áa]ficas?)?|bibliograf[ií]a|references?|reference list|bibliography|works cited)$", re.I)
CAPTION = re.compile(r"^(?:Tabla|Figura|Table|Figure)\s+(?:\d+|[A-Z]\d+)\b", re.I)
APPENDIX = re.compile(r"^(anexo|ap[eé]ndice|appendix|appendices)\b", re.I)
APPENDIX_LABEL = re.compile(r"^(ap[eé]ndice|anexo|appendix)\s+[A-Z0-9]{1,3}\b.{0,60}$", re.I)
ABSTRACT = re.compile(r"^(resumen|abstract)$", re.I)
KEYWORDS = re.compile(r"^(palabras clave|keywords|key words)\s*:", re.I)
SECTION_NAMES = re.compile(r"^(introducci[óo]n|m[ée]todo|metodolog[íi]a|resultados|discusi[óo]n|conclusi[óo]n(?:es)?|"
                           r"marco te[óo]rico|resumen|abstract|referencias|objetivos|justificaci[óo]n|planteamiento del problema|"
                           r"introduction|methods?|methodology|results|discussion|conclusions?|literature review|theoretical framework|references)$", re.I)

LEFT = WD_ALIGN_PARAGRAPH.LEFT


def _ind() -> int:
    """Sangría de primera línea / francesa (1.27 cm por defecto)."""
    return Cm(O.active().first_line_cm)


def _ls() -> float:
    return O.active().line_spacing
CENTER = WD_ALIGN_PARAGRAPH.CENTER


# ------------------------------------------------------------------ utilidades


def set_rfonts(rpr) -> bool:
    """Fuerza la fuente objetivo quitando atributos de tema que la anularían. True si cambió algo."""
    rf = rpr.find(qn("w:rFonts"))
    changed = False
    if rf is None:
        rf = OxmlElement("w:rFonts")
        rpr.insert(0, rf)
        changed = True
    for att in list(rf.attrib):
        if "Theme" in att or att.endswith("}hint"):
            del rf.attrib[att]
            changed = True
    for att in ("ascii", "hAnsi", "eastAsia", "cs"):
        if rf.get(qn(f"w:{att}")) != TARGET["font"]:
            rf.set(qn(f"w:{att}"), TARGET["font"])
            changed = True
    return changed


def heading_level(p: Paragraph) -> int:
    """0 = no es título; 1-5 = nivel; 9 = título del trabajo (estilo Title)."""
    el = p._p
    ppr = el.find(W + "pPr")
    sid = None if ppr is None else (ppr.find(W + "pStyle").get(W + "val") if ppr.find(W + "pStyle") is not None else None)
    return styles_of(p.part).level(sid)


def is_list(p: Paragraph) -> bool:
    ppr = p._p.find(W + "pPr")
    if ppr is not None and ppr.find(W + "numPr") is not None:
        return True
    sid = None if ppr is None or ppr.find(W + "pStyle") is None else ppr.find(W + "pStyle").get(W + "val")
    return styles_of(p.part).listy(sid)


def text_of(p: Paragraph) -> str:
    return text_el(p._p)


def lang_is_spanish(texts: Iterable[str]) -> bool:  # compatibilidad; preferir lang.resolve/current
    words = re.findall(r"[a-záéíóúñ]+", " ".join(texts).lower())
    if len(words) < 30:
        return True
    es = sum(1 for w in words if w in {"el", "la", "los", "las", "de", "del", "que", "y", "en", "por", "con", "para", "se", "una", "un"})
    en = sum(1 for w in words if w in {"the", "of", "and", "to", "in", "is", "that", "for", "with", "are"})
    return es >= en


def reference_range(infos: list[PInfo]) -> tuple[int, int] | None:
    """(índice del encabezado, índice final exclusivo) de Referencias. Usa la ÚLTIMA coincidencia
    (la primera puede ser una entrada de la tabla de contenido)."""
    idx = None
    for k, pi in enumerate(infos):
        if REF_HEADING.match(pi.text):
            idx = k
    if idx is None:
        return None
    j = idx + 1
    while j < len(infos) and not APPENDIX.match(infos[j].text):
        j += 1
    return idx, j


def table_paragraphs(doc: DocumentT):
    for tbl in doc.element.body.iter(W + "tbl"):
        for el in tbl.iter(P):
            yield Paragraph(el, doc._body)


def iter_all_paragraphs(doc: DocumentT):
    yield from (pi.p for pi in scan(doc))
    yield from table_paragraphs(doc)


# ------------------------------------------------------------------ página, estilos, fuentes


def choose_target(doc: DocumentT) -> dict:
    """Conserva la fuente/tamaño del documento si ya cumple APA 7; si no, Times New Roman 12 (o la fuente del perfil)."""
    a = O.active()
    if a.font:
        TARGET.update(font=a.font, size=float(a.size or ACCEPTED_FONTS.get(a.font.lower(), 12)))
        return TARGET
    infos = scan(doc)
    st = styles_of(doc)
    body = [pi for pi in infos if len(pi.text) > 40 and not pi.level and not pi.listy]
    if body:
        font = Counter((pi.font or st.eff(pi.sid)["font"]) for pi in body).most_common(1)[0][0]
        size = Counter(eff_size(pi, st) for pi in body).most_common(1)[0][0]
        if font.lower() in ACCEPTED_FONTS and abs(size - ACCEPTED_FONTS[font.lower()]) < 0.01:
            TARGET.update(font=font, size=float(size))
            return TARGET
    TARGET.update(font="Times New Roman", size=12.0)
    return TARGET


def fix_page(doc: DocumentT) -> bool:
    """Márgenes de 2.54 cm; tamaño Carta, salvo que el documento ya sea A4 (válido según la institución)."""
    changed = False
    for s in doc.sections:
        before = (s.page_width, s.page_height, s.left_margin, s.right_margin, s.top_margin, s.bottom_margin)
        ps = O.active().page_size
        is_a4 = abs(s.page_width - Cm(21.0)) < Cm(0.3) and abs(s.page_height - Cm(29.7)) < Cm(0.3)
        if ps == "a4":
            s.page_width, s.page_height = Cm(21.0), Cm(29.7)
        elif ps == "letter" or not is_a4:
            s.page_width, s.page_height = Inches(8.5), Inches(11)
        s.left_margin = s.right_margin = s.top_margin = s.bottom_margin = Cm(O.active().margin_cm)
        after = (s.page_width, s.page_height, s.left_margin, s.right_margin, s.top_margin, s.bottom_margin)
        changed |= any(abs(int(a) - int(b)) > 100 for a, b in zip(before, after))
    return changed


def _style_heading(style, level: int) -> None:
    f, pf = style.font, style.paragraph_format
    f.bold = True
    f.italic = level in (3, 5)
    pf.alignment = CENTER if level == 1 else LEFT
    pf.keep_with_next = True
    pf.left_indent = Inches(0)
    pf.first_line_indent = _ind() if level >= 4 else Inches(0)


def fix_styles(doc: DocumentT) -> int:
    """Fuente objetivo, doble espacio, sin color, títulos APA 1-5. Devuelve estilos tocados."""
    size = Pt(TARGET["size"])
    dd = doc.styles.element.find(qn("w:docDefaults"))
    if dd is not None:
        rpr = dd.find(qn("w:rPrDefault") + "/" + qn("w:rPr"))
        if rpr is not None:
            set_rfonts(rpr)
            for tag in ("w:sz", "w:szCs"):
                e = rpr.find(qn(tag))
                if e is None:
                    e = OxmlElement(tag)
                    lang = rpr.find(qn("w:lang"))
                    lang.addprevious(e) if lang is not None else rpr.append(e)
                e.set(qn("w:val"), str(int(TARGET["size"] * 2)))
    touched = 0
    skip = ("header", "footer", "footnote", "endnote", "toc", "table of", "caption", "index", "comment", "balloon")
    for st in doc.styles:
        if st.type != WD_STYLE_TYPE.PARAGRAPH:
            continue
        name = (st.name or "").lower()
        if any(name.startswith(k) for k in skip):
            continue
        set_rfonts(st.element.get_or_add_rPr())
        st.font.size = size
        if st.font.color is not None and st.font.color.type is not None:
            st.font.color.rgb = RGBColor(0, 0, 0)
        pf = st.paragraph_format
        pf.line_spacing = _ls()
        pf.space_before = pf.space_after = Pt(0)
        m = re.match(r"^(?:heading|t[ií]tulo)\s*(\d)$", st.name or "", re.I)
        if m and 1 <= int(m.group(1)) <= 5:
            _style_heading(st, int(m.group(1)))
        elif name == "title":
            st.font.bold = True
            pf.alignment = CENTER
        elif name == "normal":
            pf.alignment = LEFT
        touched += 1
    return touched


def fix_settings(doc: DocumentT) -> bool:
    """Sin guionado automático (APA 7: no dividir palabras con guiones)."""
    el = doc.settings.element.find(qn("w:autoHyphenation"))
    if el is not None and el.get(qn("w:val")) not in ("0", "false", "off"):
        el.getparent().remove(el)
        return True
    return False


def normalize_fonts(doc: DocumentT) -> dict[str, int]:
    """Una pasada sobre el cuerpo: fuente y tamaño de todos los runs; color de texto negro (salvo enlaces)."""
    body = doc.element.body
    half = str(int(TARGET["size"] * 2))
    fonts = sizes = colors = 0
    for rf in body.iter(W + "rFonts"):
        before = dict(rf.attrib)
        for att in list(rf.attrib):
            if "Theme" in att or att.endswith("}hint"):
                del rf.attrib[att]
        for att in ("ascii", "hAnsi", "eastAsia", "cs"):
            rf.set(qn(f"w:{att}"), TARGET["font"])
        fonts += dict(before) != dict(rf.attrib)
    for tag in ("sz", "szCs"):
        for e in body.iter(W + tag):
            if e.get(W + "val") != half:
                e.set(W + "val", half)
                sizes += 1
    for c in list(body.iter(W + "color")):
        if c.get(W + "val") in (None, "auto", "000000"):
            continue
        rpr = c.getparent()
        run = rpr.getparent() if rpr is not None else None
        in_link = run is not None and run.getparent() is not None and run.getparent().tag == W + "hyperlink"
        has_style = rpr is not None and rpr.find(W + "rStyle") is not None and "hyperlink" in rpr.find(W + "rStyle").get(W + "val", "").lower()
        if in_link or has_style:
            continue
        rpr.remove(c)
        colors += 1
    return {"fonts": fonts, "sizes": sizes, "colors": colors}


# ------------------------------------------------------------------ runs y párrafos


def fix_runs(p: Paragraph, bold: bool | None = None, italic: bool | None = None) -> None:
    """Negrita/cursiva de todos los runs. (La fuente y el tamaño se normalizan en una pasada global.)"""
    if bold is None and italic is None:
        return
    for r in p.runs:
        if bold is not None:
            r.bold = True if bold else None
        if italic is not None:
            r.italic = True if italic else None


def _spacing(pf) -> None:
    pf.line_spacing = _ls()
    pf.space_before = pf.space_after = Pt(0)


def format_heading(p: Paragraph, level: int) -> None:
    pf = p.paragraph_format
    p.alignment = CENTER if level == 1 else LEFT
    pf.left_indent = Inches(0)
    pf.first_line_indent = _ind() if level >= 4 else Inches(0)
    pf.keep_with_next = True
    _spacing(pf)
    fix_runs(p, bold=True, italic=level in (3, 5))


def format_body(p: Paragraph, indent: bool = True) -> None:
    pf = p.paragraph_format
    p.alignment = LEFT
    pf.first_line_indent = _ind() if indent else Inches(0)
    pf.left_indent = Inches(0)
    _spacing(pf)


def format_reference(p: Paragraph) -> None:
    pf = p.paragraph_format
    p.alignment = LEFT
    pf.left_indent = _ind()
    pf.first_line_indent = -_ind()
    _spacing(pf)


def _rebuild_runs(p: Paragraph, parts: list[tuple[str, bool]]) -> None:
    for r in list(p.runs):
        r._r.getparent().remove(r._r)
    for text, italic in parts:
        r = p.add_run(text)
        r.italic = True if italic else None
        r.font.name = TARGET["font"]
        set_rfonts(r._r.get_or_add_rPr())
        r.font.size = Pt(TARGET["size"])


def fix_paragraphs(doc: DocumentT) -> dict[str, int]:
    st = styles_of(doc)
    infos = scan(doc)
    rr = reference_range(infos)
    counts = {"headings": 0, "body": 0, "references": 0, "other": 0, "abstract": 0, "appendix": 0}
    in_abstract = False
    for pi in infos:
        t, p, i = pi.text, pi.p, pi.i
        if not t:
            continue
        pf = p.paragraph_format
        if rr and i == rr[0]:  # Referencias: título nivel 1, en página nueva
            if t != L.current().ref_heading:
                _rebuild_runs(p, [(L.current().ref_heading, False)])
            format_heading(p, 1)
            pf.page_break_before = True if i > 0 else None
            counts["headings"] += 1
            in_abstract = False
        elif ABSTRACT.match(t):
            format_heading(p, 1)
            pf.page_break_before = True if i > 0 else None
            in_abstract = True
            counts["abstract"] += 1
        elif APPENDIX_LABEL.match(t) and (rr is None or i > rr[0]):
            p.alignment = CENTER
            pf.first_line_indent = pf.left_indent = Inches(0)
            pf.page_break_before = True
            pf.keep_with_next = True
            _spacing(pf)
            fix_runs(p, bold=True)
            counts["appendix"] += 1
        elif pi.level in (1, 2, 3, 4, 5):
            format_heading(p, pi.level)
            counts["headings"] += 1
            in_abstract = False
        elif pi.level == 9:
            p.alignment = CENTER
            _spacing(pf)
            fix_runs(p, bold=True)
            counts["headings"] += 1
        elif rr and rr[0] < i < rr[1]:
            format_reference(p)
            counts["references"] += 1
        elif KEYWORDS.match(t):
            label = KEYWORDS.match(t).group(0)
            p.alignment = LEFT
            pf.first_line_indent, pf.left_indent = _ind(), Inches(0)
            _spacing(pf)
            _rebuild_runs(p, [(label, True), (t[len(label):], False)])
            counts["abstract"] += 1
            in_abstract = False
        elif pi.listy:
            p.alignment = LEFT
            pf.line_spacing = _ls()
            counts["other"] += 1
        elif CAPTION.match(t) or re.match(r"^(?:nota|note)\.", t, re.I):
            p.alignment = LEFT
            pf.first_line_indent = pf.left_indent = Inches(0)
            counts["other"] += 1
        elif eff_align(pi, st) == "center" and len(t) <= 120:
            pf.line_spacing = _ls()
            counts["other"] += 1
        elif (pi.left or 0) >= 567 and (pi.first or 0) >= 0:  # cita en bloque: conserva sangría izquierda
            p.alignment = LEFT
            pf.line_spacing = _ls()
            counts["other"] += 1
        elif in_abstract:  # el resumen va en un solo párrafo, sin sangría
            format_body(p, indent=False)
            in_abstract = False
            counts["abstract"] += 1
        elif len(t) > 40 or pi.first is not None:
            format_body(p)
            counts["body"] += 1
        else:
            p.alignment = LEFT
            pf.line_spacing = _ls()
            counts["other"] += 1
    # celdas de tabla: interlineado sencillo (APA lo permite en tablas)
    for p in table_paragraphs(doc):
        pf = p.paragraph_format
        pf.line_spacing = 1.0
        pf.space_before = pf.space_after = Pt(0)
        pf.first_line_indent = Inches(0)
    return counts


def merge_inline_headings(doc: DocumentT) -> int:
    """Niveles 4 y 5 de APA 7 son títulos en línea: «Título.» en negrita (nivel 5 también cursiva) y el texto sigue en el mismo renglón.
    Une cada título de nivel 4/5 con el párrafo de texto que le sigue y lo deja como párrafo normal (ya no es un Heading)."""
    infos = scan(doc)
    by_el = {id(pi.el): pi for pi in infos}
    n = 0
    for pi in infos:
        if pi.level not in (4, 5) or not pi.text or pi.in_table:
            continue
        nxt_el = pi.el.getnext()
        nxt = by_el.get(id(nxt_el))
        if (nxt is None or not nxt.text or nxt.level or nxt.listy or nxt.in_table or CAPTION.match(nxt.text) or KEYWORDS.match(nxt.text)
                or ABSTRACT.match(nxt.text) or REF_HEADING.match(nxt.text)):
            continue
        ts = [t for t in pi.el.iter(T) if t.text]
        if ts and not ts[-1].text.rstrip().endswith((".", "?", "!", ":")):
            ts[-1].text = ts[-1].text.rstrip() + "."
        p = pi.p
        fix_runs(p, bold=True, italic=pi.level == 5)
        old_ppr, new_ppr = pi.el.find(W + "pPr"), nxt_el.find(W + "pPr")
        if old_ppr is not None:
            pi.el.remove(old_ppr)
        if new_ppr is not None:
            pi.el.insert(0, deepcopy(new_ppr))
        sep = OxmlElement("w:r")
        t = OxmlElement("w:t")
        t.text = " "
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        sep.append(t)
        pi.el.append(sep)
        for child in list(nxt_el):
            if child.tag != W + "pPr":
                pi.el.append(child)
        nxt_el.getparent().remove(nxt_el)
        format_body(p)
        n += 1
    if n:
        clear_cache()
    return n


def promote_headings(doc: DocumentT) -> int:
    """Si el documento no usa estilos de título, convierte los párrafos que parecen título en Heading 1-5."""
    infos = scan(doc)
    if any(pi.level for pi in infos):
        return 0
    try:
        heads = [doc.styles[f"Heading {n}"] for n in range(1, 6)]
    except KeyError:
        return 0
    n = 0
    for pi in infos:
        t = pi.text
        if not t or len(t) > 90 or t.endswith((".", ":", ";", ",")) or pi.listy or CAPTION.match(t) or KEYWORDS.match(t):
            continue
        runs = [r for r in pi.el.iter(W + "r") if "".join(x.text or "" for x in r.iter(T)).strip()]
        bold = bool(runs) and all((r.find(W + "rPr") is not None and r.find(W + "rPr").find(W + "b") is not None
                                   and r.find(W + "rPr").find(W + "b").get(W + "val") not in ("0", "false")) for r in runs)
        known = bool(SECTION_NAMES.match(t))
        if not (bold or known):
            continue
        num = re.match(r"^(\d+(?:\.\d+)*)[.)]?\s", t)
        if num:
            level = min(5, num.group(1).count(".") + 1)
        elif known or REF_HEADING.match(t) or pi.align == "center":
            level = 1
        else:
            level = 2
        pi.p.style = heads[level - 1]
        n += 1
    if n:
        clear_cache()
    return n


# ------------------------------------------------------------------ correcciones de texto


def _uniform_runs(p: Paragraph) -> bool:
    sigs = {(etree.tostring(r._r.rPr) if r._r.rPr is not None else b"") for r in p.runs}  # XmlString no es hashable
    return len(sigs) <= 1


def _replace_in_paragraph(p: Paragraph, pattern: re.Pattern, repl) -> bool:
    """Sustituye por runs; si el texto cruza runs, unifica solo cuando el formato es uniforme (no pierde cursivas)."""
    changed = False
    for r in p.runs:
        new = pattern.sub(repl, r.text)
        if new != r.text:
            r.text = new
            changed = True
    if not changed and p.runs and _uniform_runs(p):
        full = "".join(r.text for r in p.runs)
        new = pattern.sub(repl, full)
        if new != full:
            p.runs[0].text = new
            for r in p.runs[1:]:
                r.text = ""
            changed = True
    return changed


_RECUP = re.compile(r"(?i)\s*\b(?:recuperado|retrieved|disponible)\s+(?:el [^,]{3,30},?\s*)?(?:de|from|en)\s+(?=https?://)")
FIXES_ALL = [
    (_RECUP, " "),
    (rules.OLD_DOI, r"https://doi.org/\1"),
    (re.compile(r"\bp\.\s*(\d+)\s*[-–]\s*(\d+)"), r"pp. \1–\2"),
    (re.compile(r"\bpp\.\s*(\d+)\s*[-–]\s*(\d+)"), r"pp. \1–\2"),  # rango ya con «pp.»: guion → raya
    (re.compile(r"\bpp\.\s*(\d+)(?!\s*[-–\d])"), r"p. \1"),
    (re.compile(r"\bet\.?\s?al\b\.?(?!\w)"), "et al."),
    # puntuación y espacios (seguros: no tocan URL, correos ni rutas)
    (textrules.DOUBLE_SPACE, " "),
    (textrules.SPACE_BEFORE, ""),
    (textrules.SPACE_BEFORE_DOT, "."),
    (textrules.NO_SPACE_COMMA, lambda m: m.group(1) + " " if textrules._is_token_safe(m.string, m.start()) else m.group(0)),
]
FIXES_ES = [(re.compile(r"(?<=[\w.)])\s*,?\s*&\s*(?=[A-ZÁÉÍÓÚÑ])"), " y ")]
_NM = r"[A-Z][\w'’-]+"
FIXES_EN = [  # inglés: (Smith & Jones, 2019) entre paréntesis; Smith and Jones (2019) en narrativa
    (re.compile(rf"\b({_NM})\s+and\s+({_NM})(,\s*(?:\d{{4}}[a-z]?|n\.d\.))"), r"\1 & \2\3"),
    (re.compile(rf"\b({_NM})\s+&\s+({_NM})(\s*\((?:\d{{4}}[a-z]?|n\.d\.)\))"), r"\1 and \2\3"),
]
FIXES_REF = [
    (re.compile(r"(?<=\.\s)(?:Bogotá|México|Madrid|Buenos Aires|Lima|Santiago|Barcelona|Valencia|New York|London)\s*[,:]\s*(?=[A-ZÁÉÍÓÚ])"), ""),
    (rules.URL_TAIL, r"\1"),
    (rules.INITIALS_GLUED, lambda m: m.group(0) + " "),
]
FIXES_REF_ES = [(rules.ORDINAL_ED, r"(\1.ª ed.)")]
FIXES_REF_EN = [
    (re.compile(r"(?<=[A-Z]\.)\s+&\s"), ", & "),
    (re.compile(r"(?<=[A-Z]\.),?\s+and\s+(?=[A-Z][\w'’-]+,\s*[A-Z]\.)"), ", & "),
    (rules.ORDINAL_ED_EN, rules._en_ed_fix),
]


def _fix_ranges(m: re.Match) -> str:
    return rules.RANGE_HYPHEN.sub(r"\1–\2", m.group(0))


_PLAIN_PART = re.compile(r"^.*?(?=https?://|$)", re.S)


def _sort_multi_citation(m: re.Match) -> str:
    inner = m.group(1)
    parts = [x.strip() for x in inner.split(";") if x.strip()]
    if len(parts) > 1 and all(rules.AUTHOR_YEAR.search(x) for x in parts):
        return "(" + "; ".join(sorted(parts, key=lambda x: rules.norm(rules.AUTHOR_YEAR.search(x).group(1)))) + ")"
    return m.group(0)


def fix_text(doc: DocumentT) -> dict[str, int]:
    infos = scan(doc)
    es = L.current().code == "es"
    rr = reference_range(infos)
    fixes = FIXES_ALL + (FIXES_ES if es else FIXES_EN)
    n = skipped = 0
    for pi in infos:
        if not pi.text:
            continue
        in_ref = bool(rr and rr[0] < pi.i < rr[1])
        cands = list(fixes)
        if in_ref:
            cands += FIXES_REF + (FIXES_REF_ES if es else FIXES_REF_EN) + [(_PLAIN_PART, _fix_ranges)]
        elif ";" in pi.text and "(" in pi.text:
            cands.append((rules.PAREN, _sort_multi_citation))
        cur = pi.raw
        for pat, rep in cands:  # en orden: cada corrección ve el texto ya corregido por las anteriores
            if pat is _PLAIN_PART:
                if not rules.RANGE_HYPHEN.search(cur.split("http", 1)[0]):
                    continue
            elif not pat.search(cur):
                continue
            if _replace_in_paragraph(pi.p, pat, rep):
                n += 1
                cur = raw_el(pi.el)
            else:
                skipped += 1
    return {"fixed": n, "skipped": skipped}


def fix_alt_text(doc: DocumentT) -> int:
    """Texto alternativo de imágenes sin descripción: usa el título del rótulo 'Figura N' más cercano."""
    infos = scan(doc)
    n = 0
    wp = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
    for k, pi in enumerate(infos):
        for d in pi.el.iter(wp + "docPr"):
            if d.get("descr"):
                continue
            title = None
            for prev in infos[max(0, k - 3):k][::-1]:
                m = re.match(r"^(?:Figura|Figure)\s+\d+\.?\s*(.*)$", prev.text, re.I)
                if m:
                    title = m.group(1) or (infos[prev.i + 1].text if prev.i + 1 < len(infos) else "")
                    break
            if title:
                d.set("descr", title)
                n += 1
    return n


# ------------------------------------------------------------------ tablas


_TBL_AFTER = ("w:shd", "w:tblLayout", "w:tblCellMar", "w:tblLook", "w:tblCaption", "w:tblDescription")
_TC_AFTER = ("w:shd", "w:noWrap", "w:tcMar", "w:textDirection", "w:tcFitText", "w:vAlign", "w:hideMark")


def _insert_before_any(parent, new, tags: tuple[str, ...]) -> None:
    wanted = {qn(t) for t in tags}
    for child in parent:
        if child.tag in wanted:
            child.addprevious(new)
            return
    parent.append(new)


_BORDERS = (f"<w:tblBorders {nsdecls('w')}>"
            '<w:top w:val="single" w:sz="8" w:space="0" w:color="000000"/><w:left w:val="nil"/>'
            '<w:bottom w:val="single" w:sz="8" w:space="0" w:color="000000"/><w:right w:val="nil"/>'
            '<w:insideH w:val="nil"/><w:insideV w:val="nil"/></w:tblBorders>')
_HEAD_BORDER = (f"<w:tcBorders {nsdecls('w')}>"
                '<w:bottom w:val="single" w:sz="8" w:space="0" w:color="000000"/></w:tcBorders>')


def format_table(table) -> None:
    """Estilo APA: solo líneas horizontales (superior, bajo el encabezado, inferior)."""
    tbl = table._tbl
    tblPr = tbl.tblPr
    for old in tblPr.findall(qn("w:tblBorders")):
        tblPr.remove(old)
    _insert_before_any(tblPr, parse_xml(_BORDERS), _TBL_AFTER)
    for ri, tr in enumerate(tbl.findall(qn("w:tr"))):
        for tc in tr.findall(qn("w:tc")):
            tcPr = tc.get_or_add_tcPr()
            for old in tcPr.findall(qn("w:tcBorders")):
                tcPr.remove(old)
            if ri == 0:
                _insert_before_any(tcPr, parse_xml(_HEAD_BORDER), _TC_AFTER)


def format_tables(doc: DocumentT) -> int:
    tables = doc.tables
    for t in tables:
        format_table(t)
    return len(tables)


# ------------------------------------------------------------------ campos (número de página, TOC)


def _fld(p: Paragraph, instr: str, placeholder: str = "1") -> None:
    def r(inner: str) -> None:
        run = p.add_run()
        run._r.append(parse_xml(inner))

    r(f'<w:fldChar {nsdecls("w")} w:fldCharType="begin"/>')
    r(f'<w:instrText {nsdecls("w")} xml:space="preserve"> {instr} </w:instrText>')
    r(f'<w:fldChar {nsdecls("w")} w:fldCharType="separate"/>')
    p.add_run(placeholder)
    r(f'<w:fldChar {nsdecls("w")} w:fldCharType="end"/>')


def _hdr_has_page(h) -> bool:
    return "PAGE" in h._element.xml


def has_page_numbers(doc: DocumentT) -> bool:
    """True si TODAS las páginas (todas las secciones, incluida la primera si es distinta) llevan número."""
    prev_ok = prev_first_ok = False
    for s in doc.sections:
        ok = prev_ok if s.header.is_linked_to_previous else _hdr_has_page(s.header)
        if not ok:
            return False
        if s.different_first_page_header_footer:
            fh = s.first_page_header
            first_ok = prev_first_ok if fh.is_linked_to_previous else _hdr_has_page(fh)
            if not first_ok:
                return False
            prev_first_ok = first_ok
        prev_ok = ok
    return True


def _write_header(hdr, section, running_head: str | None) -> None:
    hdr.is_linked_to_previous = False
    for p in list(hdr.paragraphs)[1:]:
        p._p.getparent().remove(p._p)
    p = hdr.paragraphs[0]
    for r in list(p._p.findall(W + "r")):
        p._p.remove(r)
    pf = p.paragraph_format
    pf.line_spacing, pf.space_before, pf.space_after = 1.0, Pt(0), Pt(0)
    if running_head:
        from docx.enum.text import WD_TAB_ALIGNMENT
        from docx.shared import Twips

        width = section.page_width - section.left_margin - section.right_margin
        pf.tab_stops.add_tab_stop(Twips(4680), alignment=WD_TAB_ALIGNMENT.CLEAR)  # quita la tabulación central del estilo Header
        pf.tab_stops.add_tab_stop(width, alignment=WD_TAB_ALIGNMENT.RIGHT)
        p.alignment = LEFT
        p.add_run(running_head.upper())
        p.add_run("\t")
    else:
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    _fld(p, "PAGE")
    for r in p.runs:
        r.font.name = TARGET["font"]
        set_rfonts(r._r.get_or_add_rPr())
        r.font.size = Pt(TARGET["size"])


def add_page_numbers(doc: DocumentT, running_head: str | None = None) -> int:
    """Número de página arriba a la derecha (APA 7) en todas las secciones; con encabezado abreviado a la izquierda si se indica.
    Si una sección tiene «primera página distinta» se desactiva para que la portada también lleve número. Devuelve secciones tocadas."""
    n = 0
    for k, s in enumerate(doc.sections):
        own = k == 0 or not s.header.is_linked_to_previous
        if own:
            _write_header(s.header, s, running_head)
            n += 1
        if s.different_first_page_header_footer:
            s.different_first_page_header_footer = False
            if not own:  # hereda el encabezado de la sección anterior
                continue
    return n


def request_field_update(doc: DocumentT) -> None:
    settings = doc.settings.element
    if settings.find(qn("w:updateFields")) is None:
        e = OxmlElement("w:updateFields")
        e.set(qn("w:val"), "true")
        settings.append(e)


# ------------------------------------------------------------------ referencias


def _ref_key(text: str) -> str:
    """Orden letra por letra (APA 7): sin acentos, sin espacios; 'Brown, J.' antes de 'Brown, J. A.'; mismo autor por año."""
    return re.sub(r"\s+", "", rules.norm(text))


def sort_references(doc: DocumentT) -> bool:
    infos = scan(doc)
    rr = reference_range(infos)
    if not rr:
        return False
    items = [pi for pi in infos[rr[0] + 1: rr[1]] if pi.text]
    if len(items) < 2:
        return False
    ordered = sorted(items, key=lambda pi: _ref_key(pi.text))
    if [pi.i for pi in ordered] == [pi.i for pi in items]:
        return False
    anchor = infos[rr[0]].el
    for pi in ordered:
        anchor.addnext(pi.el)
        anchor = pi.el
    return True


# ------------------------------------------------------------------ símbolos estadísticos en cursiva


def _run_italic(r) -> bool:
    rpr = r.find(W + "rPr")
    i = rpr.find(W + "i") if rpr is not None else None
    return i is not None and i.get(W + "val") not in ("0", "false")


def _runs_layout(el) -> tuple[list[tuple], str]:
    """Runs directos del párrafo con sus offsets en el texto continuo (solo contenido de <w:t>)."""
    runs, raw = [], ""
    for r in el.findall(W + "r"):
        t = "".join(x.text or "" for x in r.iter(T))
        runs.append((r, len(raw), len(raw) + len(t)))
        raw += t
    return runs, raw


def stat_unitalic(el) -> int:
    """Cantidad de símbolos estadísticos (p, M, SD, t, F…) que no están en cursiva."""
    runs, raw = _runs_layout(el)
    if not raw:
        return 0
    n = 0
    for a, _b in textrules.stat_spans(raw):
        run = next((r for r, s, e in runs if s <= a < e), None)
        if run is not None and not _run_italic(run):
            n += 1
    return n


def italicize_spans(pi: PInfo, spans: list[tuple[int, int]]) -> int:
    """Pone en cursiva los fragmentos [a, b) del texto continuo del párrafo, dividiendo los runs solo donde hace falta
    (conserva el resto del formato). Devuelve cuántos fragmentos se aplicaron."""
    runs, _raw = _runs_layout(pi.el)
    done = 0
    for r_el, s, e in runs:
        local = [(max(a, s) - s, min(b, e) - s) for a, b in spans if a < e and b > s and not _run_italic(r_el)]
        if not local or not all(c.tag in (RPR_TAG, T) for c in r_el):
            continue
        text = "".join(x.text or "" for x in r_el.iter(T))
        segs, pos = [], 0
        for a, b in local:
            if a > pos:
                segs.append((text[pos:a], False))
            segs.append((text[a:b], True))
            pos = b
        if pos < len(text):
            segs.append((text[pos:], False))
        prev = r_el
        for k, (t, it) in enumerate(segs):
            el = r_el if k == 0 else deepcopy(r_el)
            run = Run(el, pi.p)
            run.text = t
            if it:
                run.italic = True
            if k > 0:
                prev.addnext(el)
            prev = el
        done += len(local)
    return done


def fix_stat_italics(doc: DocumentT) -> int:
    """Pone en cursiva los símbolos estadísticos (divide el run solo lo necesario; conserva el resto del formato)."""
    infos = scan(doc)
    rr = reference_range(infos)
    done = 0
    for pi in infos:
        if not pi.text or (rr and rr[0] <= pi.i < rr[1]):
            continue
        spans = textrules.stat_spans(_runs_layout(pi.el)[1])
        if spans:
            done += italicize_spans(pi, spans)
    return done


def _has_italic(el) -> bool:
    return any(_run_italic(r) for r in el.iter(W + "r"))


def reference_italic_missing(infos: list[PInfo], rr) -> list[PInfo]:
    """Referencias con forma reconocible que no tienen ninguna cursiva (APA 7 exige título o revista+volumen en cursiva)."""
    if not rr:
        return []
    return [pi for pi in infos[rr[0] + 1: rr[1]] if pi.text and rules.reference_italic_spans(_runs_layout(pi.el)[1]) and not _has_italic(pi.el)]


def fix_reference_italics(doc: DocumentT) -> int:
    """Cursiva en título/revista+volumen de las referencias que no tienen ninguna (respeta las que ya tienen cursivas)."""
    infos = scan(doc)
    n = 0
    for pi in reference_italic_missing(infos, reference_range(infos)):
        n += italicize_spans(pi, rules.reference_italic_spans(_runs_layout(pi.el)[1])) > 0
    return n


_BLOCK_QUOTE = re.compile(r'^\s*[“"«](?P<q>.+?)[”"»]\s*(?P<c>\([^()]{3,140}\))?\s*[.,]?\s*$', re.S)


def fix_block_quotes(doc: DocumentT) -> int:
    """Párrafo que es una sola cita textual de 40 o más palabras → cita en bloque: sin comillas, sangría de 1.27 cm a la
    izquierda, y la fuente después del punto final. Si la cita se mezcla con texto propio solo se avisa (no se toca)."""
    infos = scan(doc)
    rr = reference_range(infos)
    n = 0

    def repl(m: re.Match) -> str:
        q = m.group("q").strip()
        if not q.endswith((".", "?", "!")):
            q += "."
        return q + (" " + m.group("c") if m.group("c") else "")

    for pi in infos:
        if not pi.text or pi.level or pi.listy or (rr and rr[0] <= pi.i < rr[1]):
            continue
        m = _BLOCK_QUOTE.match(pi.raw)
        if not m or len(m.group("q").split()) < 40:
            continue
        if _replace_in_paragraph(pi.p, _BLOCK_QUOTE, repl):
            pf = pi.p.paragraph_format
            pi.p.alignment = LEFT
            pf.left_indent, pf.first_line_indent = _ind(), Cm(0)
            n += 1
    return n


RPR_TAG = W + "rPr"


# ------------------------------------------------------------------ portada


def _title_page_range(infos: list[PInfo]) -> tuple[int, int] | None:
    """Posiciones (ini, fin) de la portada: lo que hay antes del primer salto de página, si ocurre en los primeros 20 párrafos."""
    for k, pi in enumerate(infos[:20]):
        brk = any(b.get(W + "type") == "page" for b in pi.el.iter(W + "br"))
        nxt_pb = k + 1 < len(infos) and infos[k + 1].page_break_before
        if brk or nxt_pb:
            body = [x for x in infos[:k + 1] if x.text]
            return (0, k + 1) if len(body) >= 2 else None
    return None


def title_page_extra_blanks(infos: list[PInfo]) -> list[PInfo]:
    """Párrafos vacíos sobrantes dentro de la portada: entre el título y el autor APA 7 deja una sola línea en blanco."""
    rng = _title_page_range(infos)
    if not rng:
        return []
    page = infos[rng[0]:rng[1]]
    first = next((k for k, x in enumerate(page) if x.text), None)
    if first is None:
        return []
    extra, run = [], []
    for x in page[first:]:
        if x.text or any(b.get(W + "type") == "page" for b in x.el.iter(W + "br")):
            extra += run[1:]
            run = []
        else:
            run.append(x)
    return extra


def fix_title_page_spacing(doc: DocumentT) -> int:
    extra = title_page_extra_blanks(scan(doc))
    for x in extra:
        x.el.getparent().remove(x.el)
    if extra:
        clear_cache()
    return len(extra)
