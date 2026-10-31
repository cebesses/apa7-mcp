"""Idiomas (español / inglés): todo lo que APA 7 escribe distinto según el idioma del documento."""
from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Iterable

from . import options as O


@dataclass(frozen=True)
class Lang:
    code: str
    ref_heading: str  # título de la lista de referencias
    abstract: str
    keywords: str
    table: str
    figure: str
    note: str
    appendix: str
    toc_title: str
    lot_title: str
    lof_title: str
    paren_and: str  # conjunción entre autores dentro de un paréntesis: (Pérez y Soto, 2019) / (Smith & Jones, 2019)
    narr_and: str  # en citas narrativas: Pérez y Soto (2019) / Smith and Jones (2019)
    no_date: str
    in_press: str
    chapter_in: str  # «En» / «In» (capítulo de libro)
    thesis: str  # etiqueta por defecto de una tesis
    count_nouns: str  # sustantivos que se cuentan (números del 0 al 9 en letras)


ES = Lang(
    code="es", ref_heading="Referencias", abstract="Resumen", keywords="Palabras clave", table="Tabla", figure="Figura", note="Nota",
    appendix="Apéndice", toc_title="Tabla de contenido", lot_title="Índice de tablas", lof_title="Índice de figuras",
    paren_and="y", narr_and="y", no_date="s.f.", in_press="en prensa", chapter_in="En", thesis="Tesis",
    count_nouns=("participantes|estudiantes|grupos|personas|casos|sujetos|ítems|items|preguntas|factores|variables|objetivos|fases|"
                 "etapas|categorías|docentes|autores|estudios|sesiones|pruebas|niños|niñas|mujeres|hombres|países|escuelas|"
                 "instituciones|muestras|temas|criterios|dimensiones|pasos|hipótesis|encuestados|pacientes|empresas"),
)

EN = Lang(
    code="en", ref_heading="References", abstract="Abstract", keywords="Keywords", table="Table", figure="Figure", note="Note",
    appendix="Appendix", toc_title="Table of Contents", lot_title="List of Tables", lof_title="List of Figures",
    paren_and="&", narr_and="and", no_date="n.d.", in_press="in press", chapter_in="In", thesis="Thesis",
    count_nouns=("participants|students|groups|people|cases|subjects|items|questions|factors|variables|objectives|phases|stages|"
                 "categories|teachers|authors|studies|sessions|tests|children|women|men|countries|schools|institutions|samples|"
                 "themes|criteria|dimensions|steps|hypotheses|respondents|patients|companies"),
)

LANGS = {"es": ES, "en": EN}

_ES_WORDS = {"el", "la", "los", "las", "de", "del", "que", "y", "en", "por", "con", "para", "se", "una", "un", "es", "al", "su", "como", "más"}
_EN_WORDS = {"the", "of", "and", "to", "in", "is", "that", "for", "with", "are", "was", "were", "this", "which", "by", "from", "be", "on", "as"}


def detect(texts: Iterable[str]) -> str:
    """'es' o 'en' según las palabras funcionales más frecuentes (por defecto 'es' si no hay evidencia suficiente)."""
    words = re.findall(r"[a-záéíóúñü]+", " ".join(texts).lower())
    es = sum(w in _ES_WORDS for w in words)
    en = sum(w in _EN_WORDS for w in words)
    if en >= 3 and en > es * 1.15:
        return "en"
    return "es"


def resolve(texts: Iterable[str]) -> str:
    """Fija el idioma de la llamada en curso (opción del perfil o detección automática) y lo devuelve."""
    code = O.active().language
    if code == "auto":
        code = detect(texts)
    O.RESOLVED["lang"] = code
    return code


def current() -> Lang:
    code = O.active().language
    if code == "auto":
        code = O.RESOLVED.get("lang") or "es"
    return LANGS[code]


def get(code: str | None) -> Lang:
    return LANGS[code] if code in LANGS else current()


def any_label(attr: str) -> str:
    """Alternativa regex con el rótulo de ambos idiomas (p. ej. «Tabla|Table»)."""
    return "|".join(re.escape(getattr(l, attr)) for l in LANGS.values())
