"""Lectura rápida del documento: índice de estilos en caché y datos de cada párrafo (XML directo, sin objetos python-docx).

python-docx resuelve `paragraph.style` recorriendo todos los estilos en cada llamada (O(n) por párrafo); aquí se
construye una sola vez una tabla id→propiedades efectivas y se lee cada párrafo en una pasada.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field

from docx.document import Document as DocumentT
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
T, TAB, BR, CR = W + "t", W + "tab", W + "br", W + "cr"
P, PPR, RPR, STYLE = W + "p", W + "pPr", W + "rPr", W + "style"
HEADING_NAME = re.compile(r"^(?:heading|t[ií]tulo)\s*(\d)$", re.I)
THEME_FONT = "(fuente del tema)"


def _val(el, tag: str, att: str = "val"):
    if el is None:
        return None
    c = el.find(W + tag)
    return None if c is None else c.get(W + att)


def _spacing_ratio(line: str | None, rule: str | None) -> float | None:
    if line is None:
        return None
    n = int(line)
    return n / 240 if rule in (None, "auto") else round((n / 20) / 12, 2)  # exact/atLeast: puntos → razón aprox.


def _first_indent(ppr) -> int | None:
    ind = ppr.find(W + "ind") if ppr is not None else None
    if ind is None:
        return None
    if ind.get(W + "firstLine") is not None:
        return int(ind.get(W + "firstLine"))
    if ind.get(W + "hanging") is not None:
        return -int(ind.get(W + "hanging"))
    return None


class Styles:
    """Propiedades efectivas por styleId (cadena basedOn + valores por defecto del documento)."""

    def __init__(self, doc: DocumentT):
        root = doc.styles.element
        self.raw: dict[str, dict] = {}
        self.default_id: str | None = None
        for s in root.findall(STYLE):
            if s.get(W + "type") != "paragraph":
                continue
            sid = s.get(W + "styleId")
            ppr, rpr = s.find(PPR), s.find(RPR)
            rf = rpr.find(W + "rFonts") if rpr is not None else None
            sp = ppr.find(W + "spacing") if ppr is not None else None
            self.raw[sid] = {
                "name": (_val(s, "name") or "").strip(),
                "based": _val(s, "basedOn"),
                "font": (rf.get(W + "ascii") or (THEME_FONT if rf.get(W + "asciiTheme") else None)) if rf is not None else None,
                "size": (int(_val(rpr, "sz")) / 2) if rpr is not None and _val(rpr, "sz") else None,
                "line": _spacing_ratio(sp.get(W + "line"), sp.get(W + "lineRule")) if sp is not None else None,
                "align": _val(ppr, "jc") if ppr is not None else None,
                "first": _first_indent(ppr),
                "list": ppr is not None and ppr.find(W + "numPr") is not None,
            }
            if s.get(W + "default") == "1":
                self.default_id = sid
        dd = root.find(W + "docDefaults")
        r = dd.find(W + "rPrDefault/" + RPR) if dd is not None else None
        rf = r.find(W + "rFonts") if r is not None else None
        self.d_font = (rf.get(W + "ascii") or (THEME_FONT if rf.get(W + "asciiTheme") else None)) if rf is not None else None
        self.d_size = (int(_val(r, "sz")) / 2) if r is not None and _val(r, "sz") else 10.0
        pp = dd.find(W + "pPrDefault/" + PPR) if dd is not None else None
        sp = pp.find(W + "spacing") if pp is not None else None
        self.d_line = (_spacing_ratio(sp.get(W + "line"), sp.get(W + "lineRule")) if sp is not None else None) or 1.0
        self._memo: dict[str | None, dict] = {}

    def eff(self, sid: str | None) -> dict:
        sid = sid or self.default_id
        if sid in self._memo:
            return self._memo[sid]
        out = {"font": None, "size": None, "line": None, "align": None, "first": None, "list": False}
        seen: set[str] = set()
        cur = sid
        while cur and cur in self.raw and cur not in seen:
            seen.add(cur)
            r = self.raw[cur]
            for k in ("font", "size", "line", "align", "first"):
                if out[k] is None:
                    out[k] = r[k]
            out["list"] = out["list"] or r["list"]
            cur = r["based"]
        out["font"] = out["font"] or self.d_font or THEME_FONT
        out["size"] = out["size"] or self.d_size
        out["line"] = out["line"] if out["line"] is not None else self.d_line
        out["align"] = out["align"] or "left"
        self._memo[sid] = out
        return out

    def name(self, sid: str | None) -> str:
        return self.raw.get(sid or self.default_id or "", {}).get("name", "")

    def level(self, sid: str | None) -> int:
        n = self.name(sid)
        m = HEADING_NAME.match(n)
        if m and 1 <= int(m.group(1)) <= 5:
            return int(m.group(1))
        return 9 if n.lower() == "title" else 0

    def listy(self, sid: str | None) -> bool:
        return self.eff(sid)["list"] or self.name(sid).lower().startswith("list")


_CACHE: dict[int, Styles] = {}


def clear_cache() -> None:
    _CACHE.clear()


def styles_of(doc_or_part) -> Styles:
    doc = doc_or_part
    el = doc.styles.element
    key = id(el)
    if key not in _CACHE:
        _CACHE[key] = Styles(doc)
    return _CACHE[key]


def text_el(el) -> str:
    """Texto de un <w:p> (normaliza espacios; tabulaciones y saltos → espacio)."""
    parts: list[str] = []
    for n in el.iter(T, TAB, BR, CR):
        parts.append(n.text or "" if n.tag == T else " ")
    return re.sub(r"\s+", " ", "".join(parts)).strip()


def raw_el(el) -> str:
    """Texto sin normalizar (conserva espacios dobles y tabulaciones como espacio)."""
    return "".join((n.text or "") if n.tag == T else " " for n in el.iter(T, TAB))


@dataclass(slots=True)
class PInfo:
    i: int
    p: Paragraph
    el: object
    text: str
    raw: str
    sid: str | None
    level: int
    listy: bool
    align: str | None  # jc directo
    line: float | None  # interlineado directo
    first: int | None  # sangría de primera línea (twips; negativa = francesa)
    left: int | None
    font: str | None  # fuente del primer run con fuente explícita
    size: float | None
    page_break_before: bool
    in_table: bool = False


SDT, SDT_CONTENT = W + "sdt", W + "sdtContent"
FLD, INSTR = W + "fldChar", W + "instrText"
_TOC_INSTR = re.compile(r"^\s*TOC(?:\s|$)")


def _is_toc_sdt(sdt) -> bool:
    g = sdt.find(W + "sdtPr/" + W + "docPartObj/" + W + "docPartGallery")
    return g is not None and "table of contents" in (g.get(W + "val") or "").lower()


def _walk(parent):
    """Párrafos de primer nivel, incluidos los dentro de controles de contenido (w:sdt), sin las tablas de contenido de Word."""
    for ch in parent:
        if ch.tag == P:
            yield ch
        elif ch.tag == SDT and not _is_toc_sdt(ch):
            content = ch.find(SDT_CONTENT)
            if content is not None:
                yield from _walk(content)


def body_paragraphs(doc: DocumentT) -> list:
    """Párrafos del cuerpo en orden, sin los que pertenecen a un campo TOC (tabla de contenido / índices), que no se reformatean."""
    out, depth = [], 0
    for el in _walk(doc.element.body):
        begins = ends = 0
        instr = ""
        for f in el.iter(FLD, INSTR):
            if f.tag == FLD:
                t = f.get(W + "fldCharType")
                begins += t == "begin"
                ends += t == "end"
            else:
                instr += f.text or ""
        if depth > 0:
            depth += begins - ends
            if depth < 0:
                depth = 0
            continue
        if begins and _TOC_INSTR.match(instr):
            net = begins - ends
            depth = net if net > 0 else 0
            continue
        out.append(el)
    return out


def tracked_count(doc: DocumentT) -> int:
    """Cantidad de cambios controlados (inserciones, eliminaciones, movimientos) pendientes en el cuerpo."""
    body = doc.element.body
    return sum(1 for _ in body.iter(W + "ins", W + "del", W + "moveFrom", W + "moveTo"))


def accept_tracked(doc: DocumentT) -> int:
    """Acepta todos los cambios controlados (elimina lo borrado, deja lo insertado)."""
    body = doc.element.body
    n = 0
    for el in list(body.iter(W + "del", W + "moveFrom")):
        el.getparent().remove(el)
        n += 1
    for el in list(body.iter(W + "ins", W + "moveTo")):
        parent = el.getparent()
        idx = parent.index(el)
        for k, ch in enumerate(list(el)):
            parent.insert(idx + k, ch)
        parent.remove(el)
        n += 1
    for tag in ("rPrChange", "pPrChange", "sectPrChange", "tblPrChange", "tcPrChange", "tblGridChange"):
        for el in list(body.iter(W + tag)):
            el.getparent().remove(el)
    return n


def scan(doc: DocumentT) -> list[PInfo]:
    st = styles_of(doc)
    out: list[PInfo] = []
    for el in body_paragraphs(doc):
        i = len(out)
        ppr = el.find(PPR)
        sid = _val(ppr, "pStyle")
        sp = ppr.find(W + "spacing") if ppr is not None else None
        ind = ppr.find(W + "ind") if ppr is not None else None
        first = left = None
        if ind is not None:
            if ind.get(W + "firstLine") is not None:
                first = int(ind.get(W + "firstLine"))
            elif ind.get(W + "hanging") is not None:
                first = -int(ind.get(W + "hanging"))
            if ind.get(W + "left") is not None:
                left = int(ind.get(W + "left"))
        rf = next((r for r in el.iter(W + "rFonts") if r.get(W + "ascii")), None)
        sz = next((s for s in el.iter(W + "sz") if s.get(W + "val")), None)
        direct_num = ppr is not None and ppr.find(W + "numPr") is not None
        out.append(PInfo(
            i=i, p=Paragraph(el, doc._body), el=el, text=text_el(el), raw=raw_el(el), sid=sid,
            level=st.level(sid), listy=direct_num or st.listy(sid),
            align=_val(ppr, "jc") if ppr is not None else None,
            line=_spacing_ratio(sp.get(W + "line"), sp.get(W + "lineRule")) if sp is not None else None,
            first=first, left=left,
            font=rf.get(W + "ascii") if rf is not None else None,
            size=int(sz.get(W + "val")) / 2 if sz is not None else None,
            page_break_before=ppr is not None and ppr.find(W + "pageBreakBefore") is not None,
        ))
    return out


def eff_font(pi: PInfo, st: Styles) -> str:
    return pi.font or st.eff(pi.sid)["font"]


def eff_size(pi: PInfo, st: Styles) -> float:
    return pi.size or st.eff(pi.sid)["size"]


def eff_line(pi: PInfo, st: Styles) -> float:
    return pi.line if pi.line is not None else st.eff(pi.sid)["line"]


def eff_align(pi: PInfo, st: Styles) -> str:
    a = pi.align or st.eff(pi.sid)["align"]
    return {"both": "justify", "start": "left", "end": "right"}.get(a, a)


def eff_first(pi: PInfo, st: Styles) -> int:
    return pi.first if pi.first is not None else (st.eff(pi.sid)["first"] or 0)


def has_toc(doc: DocumentT) -> bool:
    """True si el documento ya tiene una tabla de contenido (campo TOC con interruptor o, o control de contenido de Word)."""
    body = doc.element.body
    if any(_is_toc_sdt(s) for s in body.iter(SDT)):
        return True
    return any(re.match(r"^\s*TOC\s+\\o", i.text or "") for i in body.iter(INSTR))
