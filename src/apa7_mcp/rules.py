"""Reglas APA 7: catálogo, citas, referencias y formato. Extraídas de las dos guías APA 7 (español)."""
from __future__ import annotations

import json
import re
import unicodedata
from importlib import resources
from typing import Any

from . import lang as L

Issue = dict[str, Any]

# ---------------------------------------------------------------- catálogo


def list_rules(category: str | None = None) -> list[dict]:
    data = json.loads(resources.files("apa7_mcp").joinpath("data/rules.json").read_text(encoding="utf-8"))
    return [r for r in data if not category or r["category"] == category]


# ---------------------------------------------------------------- utilidades


def issue(rule_id: str, category: str, severity: str, message: str, **extra: Any) -> Issue:
    return {"ruleId": rule_id, "category": category, "severity": severity, "message": message, **extra}


def norm(s: str) -> str:
    return "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn").lower()


# ---------------------------------------------------------------- citas

PAREN = re.compile(r"\(([^()]{1,400}?)\)")
NON_CITATION = re.compile(
    r"^(?:\d|[A-Z]{2,}|[a-z]{1,3}|\d{4}(?:[-/]\d{4})?|e\.g\.|i\.e\.|etc\.?|cf\.?|p\.|pp\.|n\.|núm\.?|vol\.?|cap\.?|fig\.?|tabla\s?\d|figura\s?\d|anexo\s?[A-Z]|ver\s|vs\.?|table\s?\d|figure\s?\d|appendix\s?[A-Z]|see\s|ch\.)$",
    re.I,
)
AUTHOR_YEAR = re.compile(
    r"([A-ZÁÉÍÓÚÜÑ][a-záéíóúüñ'’-]+(?:\s+(?:y|e|and|&)\s+[A-ZÁÉÍÓÚÜÑ][a-záéíóúüñ'’-]+)*(?:\s+et\s+al\.?)?)\s*,\s*"
    r"(\d{4}[a-z]?(?:/\d{4})?|s\.?f\.?|n\.?d\.?|en\s+prensa|in\s+press)",
    re.I,
)
PAGE_LOCATOR = re.compile(r"(p\.|pp\.|párr\.|para\.|página|páginas)\s*[.,]?\s*([\d–-]+(?:\s*[-–]\s*[\d]+)?)", re.I)
_NAME = r"[A-ZÁÉÍÓÚÜÑ][a-záéíóúüñ'’-]+"
NARRATIVE = re.compile(
    rf"\b({_NAME}(?:\s{_NAME})*(?:(?:,\s+|\s+(?:y|e|&|and)\s+){_NAME}(?:\s{_NAME})*)*(?:\s+et al\.)?)\s*\((\d{{4}}[a-z]?|s\.f\.|n\.d\.)\)")


def _split_authors(s: str) -> list[str]:
    s = re.sub(r"\s+et al\.?", "", s)
    return [a.strip() for a in re.split(r",\s+|\s+(?:y|e|&|and)\s+", s, flags=re.I) if a.strip()]


def _norm_year(y: str) -> str:
    k = y.replace(".", "").lower()
    return "s.f." if k == "sf" else "n.d." if k == "nd" else y


def _locator(s: str) -> str | None:
    m = PAGE_LOCATOR.search(s)
    return m.group(0).strip() if m else None


SECONDARY = re.compile(r"(?i)\b(?:citado en|as cited in|cit\.? en)\s+")


def _parse_citation(raw: str) -> dict | None:
    sec = SECONDARY.search(raw)
    if sec:  # (Autor1, 2010, citado en Autor2, 2020): solo Autor2 aparece en la lista de referencias
        det = _parse_citation(raw[sec.end():])
        if det:
            det.update(raw=raw, secondary=True)
        return det
    parts = [p.strip() for p in raw.split(";") if p.strip()]
    first = parts[0] if parts else raw
    m = AUTHOR_YEAR.search(first)
    if not m:
        if len(parts) <= 1:
            t = re.match(r'^["“]([^"”]{1,120})["”]\s*,\s*(\d{4}|s\.?f\.?|n\.?d\.?)', raw, re.I)
            if t:
                return {"raw": raw, "authors": [], "year": _norm_year(t.group(2)), "narrative": False}
        return None
    return {
        "raw": raw,
        "authors": _split_authors(m.group(1)),
        "year": _norm_year(m.group(2)),
        "locator": _locator(first),
        "narrative": False,
    }


def _validate_citation(c: dict, issues: list[Issue], lang: str | None = None) -> None:
    raw = c["raw"]
    lg = L.get(lang)
    base = {"citation": raw, "original": f"({raw})"}
    if lg.code == "es" and "&" in raw:
        issues.append(issue("cit-002", "citacion", "warning",
                            'En español se usa "y" (no "&") para unir al último autor.',
                            suggestion=raw.replace("&", "y"), autofix=True, **base))
    elif lg.code == "en" and re.search(r"\band\b", raw) and not c.get("secondary"):
        issues.append(issue("cit-002", "citacion", "warning",
                            'En inglés, dentro del paréntesis se usa "&" (no "and"): (Smith & Jones, 2019).',
                            suggestion=re.sub(r"\s+and\s+", " & ", raw), autofix=True, **base))
    if re.search(r"\bet\.?\s?al\b", raw, re.I) and not re.search(r"\bet al\.", raw, re.I):
        issues.append(issue("cit-003", "citacion", "error", '"et al." debe escribirse con punto: et al.',
                            suggestion=re.sub(r"\bet\.?\s?al\.?", "et al.", raw, flags=re.I), **base))
    if len(c["authors"]) >= 3:
        issues.append(issue("cit-004", "citacion", "error",
                            'Con tres o más autores se cita el primero seguido de "et al." desde la primera cita.',
                            suggestion=f"({c['authors'][0]} et al., {c['year']})", **base))
    m = PAGE_LOCATOR.search(raw)
    if m:
        is_range = bool(re.search(r"[-–]", m.group(2) or ""))
        used = m.group(1).lower()
        if is_range and used == "p.":
            issues.append(issue("cit-005", "citacion", "warning", 'Para un rango de páginas use "pp.".',
                                suggestion=re.sub(r"\bp\.", "pp.", raw, flags=re.I), **base))
        if not is_range and used == "pp.":
            issues.append(issue("cit-005", "citacion", "warning", 'Para una sola página use "p.".',
                                suggestion=re.sub(r"\bpp\.", "p.", raw, flags=re.I), **base))
    if not re.search(r"\d{4}|s\.f\.|n\.d\.|en prensa|in press", raw, re.I) and c["authors"]:
        issues.append(issue("cit-006", "citacion", "warning",
                            f'Falta el año de publicación. Si no hay fecha use "{lg.no_date}".', **base))


def detect_citations(text: str, lang: str | None = None) -> tuple[list[dict], list[Issue]]:
    citations: list[dict] = []
    issues: list[Issue] = []
    for m in PAREN.finditer(text):
        raw = m.group(1).strip()
        if not raw or NON_CITATION.match(raw) or re.fullmatch(r"\d{4}", raw):
            continue
        if re.match(r"^(?:p\.|pp\.)\s*\d", raw) and len(raw) < 12:
            continue
        parts = [x.strip() for x in raw.split(";") if x.strip()]
        found = [d for d in (_parse_citation(x) for x in parts) if d] if len(parts) > 1 else []
        det = None if found else _parse_citation(raw)
        if found:  # citas múltiples: cada obra se verifica por separado
            citations.extend(found)
        elif det:
            citations.append(det)
        elif re.search(r"\b(19|20)\d{2}\b", raw):  # parece cita (tiene año) pero mal formada
            issues.append(issue("cit-001", "citacion", "warning",
                                "Cita entre paréntesis no reconocible por el patrón Autor-Fecha.",
                                citation=raw, original=f"({raw})",
                                suggestion="Use (Apellido, Año) o (Apellido, Año, p. #); varias obras separadas por ';'."))
    for c in citations:
        _validate_citation(c, issues, lang)
    for m in PAREN.finditer(text):
        raw = m.group(1).strip()
        parts = [x.strip() for x in raw.split(";") if x.strip()]
        if len(parts) > 1 and all(AUTHOR_YEAR.search(x) for x in parts):
            keys = [norm(AUTHOR_YEAR.search(x).group(1)) for x in parts]
            if keys != sorted(keys):
                ordered = "; ".join(sorted(parts, key=lambda x: norm(AUTHOR_YEAR.search(x).group(1))))
                issues.append(issue("cit-008", "citacion", "warning",
                                    "Las citas múltiples dentro de un paréntesis van en orden alfabético y separadas por punto y coma.",
                                    citation=raw, original=f"({raw})", suggestion=f"({ordered})", autofix=True))
    return citations, issues


QUOTE = re.compile(r"[“\"]([^”\"]{8,}?)[”\"]\s*(\(([^()]{1,120})\))?")
LOCATOR = re.compile(r"\b(?:p|pp|párr|párrs|para|paras|cap|ch|fig|tabla|table)\.|\bpágs?\.", re.I)
STOP = re.compile(r"[.!?]\s*$")


def detect_quotes(text: str) -> list[Issue]:
    """Citas textuales: <40 palabras entre comillas con página; ≥40 palabras en bloque."""
    out: list[Issue] = []
    for m in QUOTE.finditer(text):
        quote, paren = m.group(1), m.group(3)
        words = len(quote.split())
        if words >= 40:
            out.append(issue("cit-010", "citacion", "warning",
                             "Las citas de 40 o más palabras van en bloque: aparte, con sangría de 1.27 cm, sin comillas.",
                             citation=quote[:60] + "…", original=m.group(0)[:80]))
        elif words >= 4 and paren is not None and AUTHOR_YEAR.search(paren) and not LOCATOR.search(paren):
            out.append(issue("cit-009", "citacion", "warning",
                             "Toda cita textual lleva número de página o párrafo: (Autor, año, p. #).",
                             citation=paren, original=m.group(0)[:80],
                             suggestion=f"({paren}, p. #)"))
    return out


NARRATIVE_AMP = re.compile(r"\b([A-ZÁÉÍÓÚÑ][a-záéíóúñ]+)\s+&\s+([A-ZÁÉÍÓÚÑ][a-záéíóúñ]+)\s*\((\d{4}[a-z]?|s\.f\.|n\.d\.)\)")


def detect_narrative_amp(text: str, lang: str | None = None) -> list[Issue]:
    lg = L.get(lang)
    ex = "Pérez y Soto (2019)" if lg.code == "es" else "Smith and Jones (2019)"
    return [issue("cit-011", "citacion", "warning",
                  f'En citas narrativas se usa "{lg.narr_and}" (no "&"): {ex}.',
                  citation=m.group(0), original=m.group(0), suggestion=f"{m.group(1)} {lg.narr_and} {m.group(2)} ({m.group(3)})", autofix=True)
            for m in NARRATIVE_AMP.finditer(text)]


def detect_narrative(text: str) -> list[dict]:
    return [
        {"raw": f"{m.group(1)} ({m.group(2)})", "authors": _split_authors(m.group(1)),
         "year": _norm_year(m.group(2)), "narrative": True}
        for m in NARRATIVE.finditer(text)
    ]


# ---------------------------------------------------------------- referencias

CITIES = r"(Bogotá|México|Madrid|Buenos Aires|Lima|Santiago|New York|London|Barcelona|Valencia)"
URL_TAIL = re.compile(r"(https?://\S+?)\.\s*$")
OLD_DOI = re.compile(r"(?i)(?:\bdoi\s*:\s*|https?://dx\.doi\.org/|http://doi\.org/|\bDOI\s+)(10\.\d{4,9}/\S+)")
INITIALS_GLUED = re.compile(r"\b[A-ZÁÉÍÓÚÑ]\.(?=[A-ZÁÉÍÓÚÑ]\.)")
ORDINAL_ED = re.compile(r"\((\d+)\s*(?:ra|da|era|a|º|°|st|nd|rd|th)\.?\s*ed\.?\)", re.I)
ORDINAL_ED_EN = re.compile(r"\((\d+)\s*(?:st|nd|rd|th)?\.?\s*ed(?:ition)?\.?\)", re.I)


def en_ordinal(n: int) -> str:
    suf = "th" if 10 <= n % 100 <= 20 else {1: "st", 2: "nd", 3: "rd"}.get(n % 10, "th")
    return f"{n}{suf}"


def _en_ed_fix(m: "re.Match") -> str:
    return f"({en_ordinal(int(m.group(1)))} ed.)"
RANGE_HYPHEN = re.compile(r"(?<=[,(\s])(\d{1,5})-(\d{1,5})(?=[.,)\s]|$)")
LOWER_WORDS = {"de", "del", "la", "las", "el", "los", "y", "e", "o", "u", "en", "a", "con", "por", "para", "sobre", "un", "una",
               "al", "su", "sus", "lo", "que", "se", "entre", "sin", "ante", "the", "of", "and", "in", "on", "for", "to", "a", "an"}


def _title_of(ref: str) -> str | None:
    """Título del trabajo: primera oración después de '(año).'"""
    m = re.search(r"\((?:\d{4}[a-z]?|s\.f\.|n\.d\.|en prensa|in press)[^)]*\)\.\s*(.+?)(?:\.\s|\.$|\s*\[|$)", ref)
    return m.group(1) if m else None


def _title_case_ratio(title: str) -> tuple[int, int]:
    words = [w for w in re.findall(r"[A-Za-zÁÉÍÓÚÜÑáéíóúüñ][\wÁÉÍÓÚÜÑáéíóúüñ'-]*", title)]
    # se ignora la primera palabra y las que siguen a ':' (también llevan mayúscula en oración)
    cands = [w for i, w in enumerate(words) if i > 0 and w.lower() not in LOWER_WORDS and len(w) > 3]
    caps = [w for w in cands if w[0].isupper()]
    return len(caps), len(cands)


def validate_reference(ref: str, lang: "str | bool | None" = None) -> list[Issue]:
    code = "es" if lang is True else "en" if lang is False else (lang or L.current().code)
    es = code == "es"
    ref = ref.strip()
    out: list[Issue] = []
    if not ref:
        return out
    base = {"reference": ref}
    if re.search(r"(?i)\b(?:recuperado|retrieved|disponible)\s+(?:el [^,]{3,30},?\s*)?(?:de|from|en)\s+(?=https?://)", ref):
        sug = re.sub(r"(?i)\s*\b(?:recuperado|retrieved|disponible)\s+(?:el [^,]{3,30},?\s*)?(?:de|from|en)\s+(?=https?://)", " ", ref).strip()
        out.append(issue("ref-001", "referencias", "error",
                         'APA 7 eliminó "Recuperado de"/"Retrieved from": indique solo la URL o el DOI.', suggestion=sug, autofix=True, **base))
    if OLD_DOI.search(ref):
        out.append(issue("ref-010", "referencias", "error",
                         "El DOI debe ir como enlace https://doi.org/10.xxxx (sin 'doi:' ni dx.doi.org).",
                         suggestion=OLD_DOI.sub(r"https://doi.org/\1", ref), autofix=True, **base))
    if URL_TAIL.search(ref):
        out.append(issue("ref-011", "referencias", "warning", "No se coloca punto final después de una URL o DOI.",
                         suggestion=URL_TAIL.sub(r"\1", ref), autofix=True, **base))
    if re.search(CITIES + r"\s*[,:]\s*[A-Z]", ref) and re.search(r"Editorial|Ediciones|Press|Universidad|Libros", ref, re.I):
        out.append(issue("ref-002", "referencias", "warning", "En APA 7 ya no se incluye la ciudad/país de la editorial.",
                         autofix=True, **base))
    if es and re.search(r",?\s*&\s*", ref) and "doi.org" not in ref:
        out.append(issue("ref-003", "referencias", "warning",
                         'En referencias en español se usa "y" (no "&") y no va precedida de coma.',
                         suggestion=re.sub(r",?\s*&\s*", " y ", ref, count=1), autofix=True, **base))
    if not es:
        if re.search(r"[A-Z]\.\s+&\s", ref):
            out.append(issue("ref-003", "referencias", "warning", 'En inglés se escribe coma antes de "&": Smith, J., & Jones, M.',
                             suggestion=re.sub(r"(?<=[A-Z]\.)\s+&\s", ", & ", ref, count=1), autofix=True, **base))
        elif re.search(r"[A-Z]\.,?\s+and\s+[A-Z][\w'’-]+,\s*[A-Z]\.", ref):
            out.append(issue("ref-003", "referencias", "warning", 'En la lista de referencias en inglés se usa "&", no "and".',
                             suggestion=re.sub(r"(?<=[A-Z]\.),?\s+and\s+(?=[A-Z][\w'’-]+,\s*[A-Z]\.)", ", & ", ref, count=1), autofix=True, **base))
    if re.search(r"et al\.", ref, re.I):
        out.append(issue("ref-005", "referencias", "error",
                         '"et al." se usa en citas, NO en la lista de referencias: liste hasta 20 autores.', **base))
    if re.search(r"\b(?:Dr\.?|Dra\.?|Prof\.?|PhD|PsyD|EdD)\b", ref):
        out.append(issue("ref-006", "referencias", "warning", "Omita títulos (Dr., Prof.) y grados (PhD, EdD).", **base))
    if not re.search(r"\((?:\d{4}[a-z]?|s\.f\.|n\.d\.|en prensa|in press|\d{4},\s*[^)]+)\)", ref):
        out.append(issue("ref-004", "referencias", "error",
                         "Falta el año de publicación entre paréntesis después del autor: (Año).", **base))
    author_part = ref.split("(", 1)[0]
    if re.search(r",\s+[A-ZÁÉÍÓÚÑ][a-záéíóúñ]{2,}(?:\s+[A-ZÁÉÍÓÚÑ][a-záéíóúñ]{2,})*\s*[,.]?\s*$", author_part.strip() + " ") \
            and "(" in ref and not re.search(r"\b(?:Ed|Eds|Coord|Dir)\.", author_part):
        out.append(issue("ref-012", "referencias", "warning",
                         "Los nombres de pila van como iniciales (Pérez, J. M.), no completos.", **base))
    if INITIALS_GLUED.search(ref):
        out.append(issue("ref-013", "referencias", "warning", "Separe las iniciales con espacio: J. M. (no J.M.).",
                         suggestion=INITIALS_GLUED.sub(lambda m: m.group(0) + " ", ref), autofix=True, **base))
    plain = ref.split("http", 1)[0]
    if RANGE_HYPHEN.search(plain):
        out.append(issue("ref-014", "referencias", "warning", "Los rangos de páginas usan raya corta (–): 20–30, no 20-30.",
                         suggestion=RANGE_HYPHEN.sub(r"\1–\2", plain) + ref[len(plain):], autofix=True, **base))
    if es and ORDINAL_ED.search(ref):
        out.append(issue("ref-016", "referencias", "warning", "La edición se escribe '(2.ª ed.)'.",
                         suggestion=ORDINAL_ED.sub(r"(\1.ª ed.)", ref), autofix=True, **base))
    if not es:
        m = ORDINAL_ED_EN.search(ref)
        if m and m.group(0) != _en_ed_fix(m):
            out.append(issue("ref-016", "referencias", "warning", "La edición se escribe '(2nd ed.)'.",
                             suggestion=ORDINAL_ED_EN.sub(_en_ed_fix, ref), autofix=True, **base))
    t = _title_of(ref)
    if t:
        caps, total = _title_case_ratio(t)
        if total >= 3 and caps / total >= 0.6:
            msg = ("En español los títulos de las referencias llevan mayúscula solo en la primera palabra y en nombres propios."
                   if es else "In English reference titles use sentence case: capitalize only the first word, after a colon, and proper nouns.")
            out.append(issue("ref-015", "referencias", "warning", msg, suggestion=f"Revise: «{t}»", **base))
    return out


def looks_like_reference(line: str) -> bool:
    l = line.strip()
    if len(l) < 12 or re.match(r"^(Referencias|Bibliografía|Anexo|Apéndice|Tabla|Figura)\b", l, re.I):
        return False
    return bool(re.match(
        r"^[A-ZÁÉÍÓÚÜÑ][a-záéíóúüñ]+(?:\s+[A-ZÁÉÍÓÚÜÑ][a-záéíóúüñ]+)*,?\s+[A-ZÁÉÍÓÚÜÑ]\.?,?\s*(\([^)]*\d{4}[^)]*\)|\d{4})", l))


def _author_to_ref(author: str, initials: str | None = None) -> str:
    a = author.strip()
    if initials:
        return f"{a}, {initials.strip()}"
    parts = a.split()
    if len(parts) <= 1:
        return a
    given = parts[-1]
    init = " ".join(f"{c}." for c in given if c.isupper())
    return f"{' '.join(parts[:-1])}, {init}"


def _authors_text(i: dict, lg: "L.Lang | None" = None) -> str:
    lg = lg or L.current()
    names: list[str] = []
    if i.get("corporateAuthor"):
        names.append(i["corporateAuthor"])
    inits = i.get("initials") or []
    for k, a in enumerate(i.get("authors") or []):
        names.append(_author_to_ref(a, inits[k] if k < len(inits) else None))
    if not names:
        return ""
    role = f" ({i['role']})" if i.get("role") else ""
    if len(names) == 1:
        return names[0] + role
    last = names.pop()
    if lg.code == "en":  # inglés: coma + "&" también con dos autores (Smith, J., & Jones, M.)
        return f"{', '.join(names)}, & {last}{role}"
    return f"{', '.join(names)} y {last}{role}"  # español: "y" sin coma previa (Manual APA 7 adaptado)


def format_reference(i: dict, lang: str | None = None) -> dict:
    """Referencia APA 7 desde datos estructurados. Usa *cursiva* para los elementos en cursiva."""
    lg = L.get(lang)
    t = i.get("type", "libro")
    title = i["title"]
    sub = f": {i['subtitle']}" if i.get("subtitle") else ""
    author = _authors_text(i, lg)
    if author and not author.endswith("."):
        author += "."
    year = f"({i.get('fullDate') or i.get('year') or lg.no_date})"
    url = f" {i['url']}" if i.get("url") else ""
    pub = f" {i['publisher']}." if i.get("publisher") else ""
    notes: list[str] = []
    head = f"{author} {year}." if author else None

    if t in ("libro", "libro_editado"):
        ed = f" ({i['edition']})" if i.get("edition") else ""
        text = f"{head} *{title}{sub}*{ed}.{pub}{url}" if head else f"*{title}{sub}*{ed}. {year}.{pub}{url}"
    elif t == "capitulo_libro":
        eds = ", ".join(_author_to_ref(e) for e in (i.get("bookEditors") or []))
        many = "s" if len(i.get("bookEditors") or []) > 1 else ""
        pages = f" (pp. {i['pages']})" if i.get("pages") else ""
        text = f"{head or title} {title}{sub}. {lg.chapter_in} {eds} (Ed{many}.), *{i.get('bookTitle', '')}*{pages}.{pub}{url}"
        if not head:
            text = f"{title}{sub}. {year}. {lg.chapter_in} {eds} (Ed{many}.), *{i.get('bookTitle', '')}*{pages}.{pub}{url}"
    elif t == "articulo_revista":
        vol = f", *{i['volume']}*" if i.get("volume") else ""
        num = f"({i['number']})" if i.get("number") else ""
        pages = f", {i['pages']}" if i.get("pages") else ""
        text = f"{head} {title}{sub}. *{i.get('journal', '')}*{vol}{num}{pages}.{url}" if head else \
               f"{title}{sub}. {year}. *{i.get('journal', '')}*{vol}{num}{pages}.{url}"
    elif t == "articulo_periodico":
        text = f"{head} {title}. *{i.get('journal', '')}*.{url}" if head else f"{title}. {year}. *{i.get('journal', '')}*.{url}"
    elif t == "tesis":
        br = i.get("bracket") or lg.thesis
        text = f"{head} *{title}* [{br}, {i.get('publisher', '')}].{url}"
    elif t == "informe":
        n = f" ({i['reportNumber']})" if i.get("reportNumber") else ""
        text = f"{head or ''} *{title}*{n}.{pub}{url}".strip()
    elif t == "pagina_web":
        site = f" {i['siteName']}." if i.get("siteName") else ""
        text = f"{head} *{title}*.{site}{url}" if head else f"*{title}*. {year}.{site}{url}"
    elif t in ("pelicula", "video", "podcast"):
        br = i.get("bracket") or {"pelicula": "Película", "video": "Video", "podcast": "Podcast"}[t]
        text = f"{head or ''} *{title}* [{br}].{pub}{url}".strip()
    elif t == "ley":
        text = f"{title}. {year}.{pub}{url}"
    elif t == "red_social":
        handle = f" [{i['handle']}]" if i.get("handle") else ""
        br = f" [{i['bracket']}]" if i.get("bracket") else ""
        site = f" {i['siteName']}." if i.get("siteName") else ""
        text = f"{author.rstrip('.')}{handle}. {year}. {title}{br}.{site}{url}"
    else:
        text = f"{head or ''} {title}.{url}".strip()
    if not i.get("url"):
        notes.append("Sin URL/DOI. En APA 7 se incluye el DOI (https://doi.org/…) cuando existe; no use 'Recuperado de'.")
    return {"text": re.sub(r"\s+", " ", text).strip(), "type": t, "notes": notes}


# ---------------------------------------------------------------- formato

ACCEPTED_FONTS = {"times new roman": 12, "georgia": 11, "computer modern": 10, "calibri": 11, "arial": 11,
                  "lucida sans unicode": 10, "aptos": 11}


def validate_formatting(f: dict) -> list[Issue]:
    from . import options as O

    o = O.active()
    out: list[Issue] = []
    w, h = f["page_w_cm"], f["page_h_cm"]
    is_letter = abs(w - 21.59) <= 0.5 and abs(h - 27.94) <= 0.5
    is_a4 = abs(w - 21.0) <= 0.5 and abs(h - 29.7) <= 0.5
    ok_page = {"auto": is_letter or is_a4, "letter": is_letter, "a4": is_a4}[o.page_size]
    if w > 0 and not ok_page:
        want = {"auto": "Carta (21.59 x 27.94 cm) o A4 si la institución lo indica", "letter": "Carta (21.59 x 27.94 cm)", "a4": "A4 (21 x 29.7 cm)"}[o.page_size]
        out.append(issue("fmt-001", "formato", "warning", f"El tamaño de página debe ser {want}.",
                         current=f"{w:.2f} x {h:.2f} cm", expected=want, autofix=True))
    names = {"top": "superior", "bottom": "inferior", "left": "izquierdo", "right": "derecho"}
    for side, v in f["margins_cm"].items():
        if v > 0 and abs(v - o.margin_cm) > 0.15:
            out.append(issue("fmt-002", "formato", "warning", f"El margen {names[side]} debe ser de {o.margin_cm:g} cm.",
                             current=f"{v:.2f} cm", expected=f"{o.margin_cm:g} cm", autofix=True))
    font = (f["font"] or "").lower()
    if o.font:  # el perfil fija una fuente concreta
        want_size = o.size or ACCEPTED_FONTS.get(o.font.lower(), 12)
        if font and font != o.font.lower():
            out.append(issue("fmt-003", "formato", "warning", f'La fuente debe ser {o.font} {want_size:g} pt según el perfil.',
                             current=f["font"], expected=f"{o.font} {want_size:g}", autofix=True))
        elif font and f["size"] and abs(f["size"] - want_size) > 0.01:
            out.append(issue("fmt-004", "formato", "warning", f"El tamaño de fuente debe ser {want_size:g} pt.",
                             current=f"{f['size']} pt", expected=f"{want_size:g} pt", autofix=True))
    elif font and font not in ACCEPTED_FONTS:
        out.append(issue("fmt-003", "formato", "warning",
                         f'La fuente "{f["font"]}" no está en la lista APA 7 (Times New Roman 12, Calibri 11, Arial 11, Georgia 11, Aptos 11…).',
                         current=f["font"], expected="Times New Roman 12", autofix=True))
    elif font and f["size"] and abs(f["size"] - ACCEPTED_FONTS[font]) > 0.01:
        out.append(issue("fmt-004", "formato", "warning", f"El tamaño de fuente debe ser {ACCEPTED_FONTS[font]} pt.",
                         current=f"{f['size']} pt", expected=f"{ACCEPTED_FONTS[font]} pt", autofix=True))
    if f["spacing"] is not None and abs(f["spacing"] - o.line_spacing) > 0.05:
        out.append(issue("fmt-005", "formato", "warning", f"El interlineado debe ser {o.line_spacing:g} ({'doble' if o.line_spacing == 2 else 'según el perfil'}).",
                         current=str(f["spacing"]), expected=f"{o.line_spacing:g}", autofix=True))
    if f["alignment"] == "justify":
        out.append(issue("fmt-006", "formato", "warning", "El texto debe ir alineado a la izquierda (no justificado).",
                         current="justificado", expected="izquierda", autofix=True))
    if f["body_paragraphs"] > 5 and f["first_line_indented"] == 0:
        out.append(issue("fmt-007", "formato", "warning", f"La primera línea de cada párrafo lleva sangría de {o.first_line_cm:g} cm.",
                         current="0 cm", expected=f"{o.first_line_cm:g} cm", autofix=True))
    return out


# ---------------------------------------------------------------- cursivas en referencias (APA 7)
_YEAR_END = re.compile(r"\((?P<y>\d{4}[a-z]?|s\.f\.|n\.d\.|en prensa|in press)\)\.\s*")
_ARTICLE = re.compile(r"(?P<j>[A-ZÁÉÍÓÚÑ][^,.]{2,100}),\s+(?P<v>\d{1,4})(?P<i>\(\d[^)]*\))?(?=,|\.|\s|$)")
_CHAPTER = re.compile(r"\b(?:En|In)\s+[^()]{2,150}?\((?:Eds?|Coords?|Dirs?)\.\),\s+(?P<b>[^.(]{3,150}?)(?=\s*\(|\.\s|\.$|$)")


def reference_italic_spans(text: str) -> list[tuple[int, int]]:
    """Fragmentos que APA 7 escribe en cursiva en una referencia con forma «Autor. (Año). Título. …»:
    revista y volumen (artículos), título del libro del capítulo, o el título (libros, informes, tesis, páginas web).
    Devuelve [] cuando no puede decidirlo con seguridad (fechas completas, sin autor, formatos raros)."""
    m = _YEAR_END.search(text)
    if not m or m.start() < 3:  # sin autor delante o sin año reconocible
        return []
    after = m.end()
    rest = text[after:]
    t_end = rest.find(". ")
    if t_end > 0:
        tail = rest[t_end + 2:]
        art = _ARTICLE.match(tail)
        if art:
            base = after + t_end + 2
            return [(base + art.start("j"), base + art.end("j")), (base + art.start("v"), base + art.end("v"))]
    ch = _CHAPTER.search(rest)
    if ch:
        return [(after + ch.start("b"), after + ch.end("b"))]
    end = t_end if t_end > 0 else len(rest.rstrip("."))
    title = rest[:end]
    if "[" in title:
        title = title[: title.index("[")].rstrip()
    title = re.sub(r"\s*\([^)]*\bed\.[^)]*\)\s*$", "", title)  # la edición (2.ª ed.) no va en cursiva
    if len(title) < 3 or re.match(r"(?i)^(en|in)\s", title):
        return []
    return [(after, after + len(title))]


_REF_YEAR = re.compile(r"\((\d{4})([a-z]?)\)")


def ambiguous_year_groups(references: list[str]) -> list[list[str]]:
    """Grupos de referencias con el mismo primer autor y año sin letra (2019, 2019): APA exige 2019a, 2019b."""
    groups: dict[tuple[str, str], list[str]] = {}
    for r in references:
        y = _REF_YEAR.search(r)
        if not y or y.group(2):
            continue
        author = norm(re.split(r"[,(]", r.strip(), maxsplit=1)[0])
        groups.setdefault((author, y.group(1)), []).append(r)
    return [g for g in groups.values() if len(g) >= 2]
