"""Reglas de texto: símbolos estadísticos (cursiva) y revisión mecánica de puntuación/espacios."""
from __future__ import annotations

import re

from .rules import Issue, issue

# ---------------------------------------------------------------- estadísticos (APA 7: símbolos en cursiva)
# Símbolo seguido de comparador y número:  p < .05, M = 3.2, SD = 0.8, n = 20, r = .45
STAT_A = re.compile(r"(?<![\w.])(p|M|SD|SE|Mdn|N|n|r|d|z|U)(?=\s*(?:=|<|>|≤|≥)\s*[-−–]?\s*[.,]?\d)")
# Símbolo con grados de libertad:  t(23) = 2.1, F(1, 20) = 4.5, χ2(2)
STAT_B = re.compile(r"(?<![\w.])(t|F|χ2|χ²)(?=\s*\(\s*\d)")
STAT_PATTERNS = (STAT_A, STAT_B)
STAT_ZERO = re.compile(r"(?<![\w.])(p|r)\s*(?:=|<|>)\s*0[.,]\d")


def stat_spans(raw: str) -> list[tuple[int, int]]:
    return sorted(m.span(1) for pat in STAT_PATTERNS for m in pat.finditer(raw))


def detect_stat_zero(text: str) -> list[Issue]:
    return [issue("stat-002", "texto", "info",
                  "En p y r (no pueden pasar de 1) APA 7 omite el cero inicial: p = .05 (en español puede usarse coma según la institución).",
                  original=m.group(0)) for m in list(STAT_ZERO.finditer(text))[:5]]


# ---------------------------------------------------------------- puntuación y espacios
REPEATED = re.compile(r"(?<![\w.])([A-Za-zÁÉÍÓÚÜÑáéíóúüñ]{2,})\s+\1(?![\w])", re.I)
DOUBLE_SPACE = re.compile(r"(?<=\S)  +(?=\S)")
SPACE_BEFORE = re.compile(r"(?<=[\wáéíóúñ)”\"]) +(?=[,;])")
SPACE_BEFORE_DOT = re.compile(r"(?<=[a-záéíóúñ]) +\.(?!\.)(?=\s|$)")
NO_SPACE_COMMA = re.compile(r"(?<=[a-záéíóúñ])([,;])(?=[A-Za-záéíóúñ])")


def _is_token_safe(text: str, pos: int) -> bool:
    """False si la posición está dentro de una URL, correo o ruta (no se toca)."""
    start = text.rfind(" ", 0, pos) + 1
    end = text.find(" ", pos)
    token = text[start: end if end != -1 else len(text)]
    return not any(x in token for x in ("http", "@", "/", "www."))


def proofread(text: str) -> list[Issue]:
    """Revisión mecánica de un párrafo o texto: palabras repetidas, espacios y signos emparejados."""
    out: list[Issue] = []
    for line in text.splitlines():
        if not line.strip():
            continue
        for m in REPEATED.finditer(line):
            if m.group(1).lower() in {"ja", "ha", "bla", "no"} or m.group(1).isdigit():
                continue
            out.append(issue("txt-001", "texto", "warning", f"Palabra repetida: «{m.group(0)}».", original=m.group(0),
                             suggestion=m.group(1)))
        if DOUBLE_SPACE.search(line):
            out.append(issue("txt-002", "texto", "info", "Hay espacios dobles.", original=line[:60], autofix=True))
        if SPACE_BEFORE.search(line) or SPACE_BEFORE_DOT.search(line):
            out.append(issue("txt-003", "texto", "info", "Hay un espacio antes de un signo de puntuación.", original=line[:60], autofix=True))
        for m in NO_SPACE_COMMA.finditer(line):
            if _is_token_safe(line, m.start()):
                out.append(issue("txt-004", "texto", "info", "Falta un espacio después de la coma o el punto y coma.",
                                 original=line[max(0, m.start() - 15): m.end() + 15], autofix=True))
                break
        probs = []
        if line.count("(") > line.count(")"):
            probs.append("paréntesis sin cerrar")
        if line.count("“") != line.count("”"):
            probs.append("comillas “ ” sin emparejar")
        if line.count("¿") > line.count("?"):
            probs.append("¿ sin su ?")
        if line.count("¡") > line.count("!"):
            probs.append("¡ sin su !")
        if probs:
            out.append(issue("txt-005", "texto", "warning", "Signos sin emparejar: " + ", ".join(probs) + ".", original=line[:80]))
    return out
