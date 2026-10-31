"""Opciones y perfiles: lo que cambia entre instituciones (fuente, tamaño de página, márgenes, interlineado, idioma…).

Se activan con `use_options(...)`, que restaura el estado anterior al salir (sin fugas entre llamadas).
"""
from __future__ import annotations

import contextlib
import json
import os
from dataclasses import asdict, dataclass, fields, replace
from pathlib import Path

from .errors import ApaError

# Fuente/tamaño efectivos de la llamada en curso (los lee wordfmt). `use_options` los fija y restaura.
TARGET = {"font": "Times New Roman", "size": 12.0}
RESOLVED: dict = {"lang": None}  # idioma resuelto de la llamada en curso (lo fija lang.resolve)


@dataclass(frozen=True)
class Options:
    font: str | None = None  # None = conservar la del documento si ya cumple APA; si no, Times New Roman
    size: float | None = None  # None = el recomendado para la fuente
    page_size: str = "auto"  # auto (conserva A4 si ya lo es, si no Carta) | letter | a4
    margin_cm: float = 2.54
    line_spacing: float = 2.0
    first_line_cm: float = 1.27
    numbered_headings: bool = False  # True: la institución numera los títulos (no se avisa)
    language: str = "auto"  # auto | es | en
    title_page: str = "none"  # none | student | professional (informativo; la portada es opcional)

    @classmethod
    def from_dict(cls, d: dict) -> "Options":
        valid = {f.name for f in fields(cls)}
        unknown = sorted(set(d) - valid)
        if unknown:
            raise ApaError("INVALID_PROFILE", f"Opciones desconocidas en el perfil: {', '.join(unknown)}.",
                           f"Opciones válidas: {', '.join(sorted(valid))}.")
        o = cls(**d)
        if o.page_size not in ("auto", "letter", "a4"):
            raise ApaError("INVALID_PROFILE", f"page_size inválido: «{o.page_size}».", "Use auto, letter o a4.")
        if o.language not in ("auto", "es", "en"):
            raise ApaError("INVALID_PROFILE", f"language inválido: «{o.language}».", "Use auto, es o en.")
        if o.title_page not in ("none", "student", "professional"):
            raise ApaError("INVALID_PROFILE", f"title_page inválido: «{o.title_page}».", "Use none, student o professional.")
        for name in ("margin_cm", "line_spacing", "first_line_cm"):
            v = getattr(o, name)
            if not isinstance(v, (int, float)) or not (0 < v < 20):
                raise ApaError("INVALID_PROFILE", f"{name} inválido: {v!r}.", "Debe ser un número positivo.")
        if o.size is not None and not (6 <= o.size <= 24):
            raise ApaError("INVALID_PROFILE", f"size inválido: {o.size!r}.", "Entre 6 y 24 pt.")
        return o


PROFILES: dict[str, Options] = {
    "apa7": Options(page_size="letter"),
    "a4": Options(page_size="a4"),
    "auto": Options(),
}

_active: Options = Options()


def active() -> Options:
    return _active


def load(profile: str | dict | None, doc_path: str | None = None) -> Options:
    """Perfil explícito (nombre, ruta a JSON o dict) → `apa7.profile.json` junto al documento → APA7_PROFILE → por defecto."""
    src: str | dict | None = profile
    if src is None and doc_path:
        side = Path(doc_path).with_name("apa7.profile.json")
        if side.exists():
            src = str(side)
    if src is None:
        src = os.environ.get("APA7_PROFILE") or None
    if src is None:
        return Options()
    if isinstance(src, dict):
        return Options.from_dict(src)
    if src in PROFILES:
        return PROFILES[src]
    p = Path(str(src).strip().strip('"'))
    if p.exists():
        try:
            data = json.loads(p.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError) as e:
            raise ApaError("INVALID_PROFILE", f"No se pudo leer el perfil {p}: {e}", "Debe ser un JSON válido.") from e
        base = PROFILES.get(data.pop("extends", "auto"), Options())
        return Options.from_dict({**asdict(base), **data})
    raise ApaError("INVALID_PROFILE", f"Perfil desconocido: «{src}».", f"Perfiles incluidos: {', '.join(PROFILES)}; o la ruta a un .json.")


@contextlib.contextmanager
def use_options(opts: Options):
    """Activa las opciones durante la llamada y restaura el estado previo (incluida la fuente) al salir, aunque falle."""
    global _active
    prev, prev_target, prev_lang = _active, dict(TARGET), RESOLVED["lang"]
    _active = opts
    RESOLVED["lang"] = None
    size = opts.size
    TARGET.update(font=opts.font or "Times New Roman", size=float(size) if size else 12.0)
    try:
        yield opts
    finally:
        _active = prev
        TARGET.update(prev_target)
        RESOLVED["lang"] = prev_lang


def evolve(**kw) -> Options:
    return replace(_active, **kw)
