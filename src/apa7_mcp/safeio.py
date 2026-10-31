"""Escritura segura: temporal → verificación → reemplazo atómico. El original nunca se toca."""
from __future__ import annotations

import os
import re
import zipfile
from pathlib import Path

from docx import Document
from lxml import etree

from .errors import ApaError, clean_path


def resolve_output(inp: str | None, out: str | None, suffix: str = ".apa7.docx", overwrite: bool = False) -> str:
    """Destino de salida: nunca el mismo archivo de entrada; si existe, agrega -2, -3… salvo overwrite."""
    if out:
        target = clean_path(out)
        if target.suffix.lower() != ".docx":
            target = target.with_suffix(".docx")
    else:
        if inp is None:
            raise ApaError("INVALID_ARGUMENT", "Falta output_path.", "Indique dónde guardar el .docx.")
        p = Path(inp)
        target = p.with_name(p.stem + suffix)
    if inp is not None and target.resolve() == Path(inp).resolve():
        raise ApaError("INVALID_ARGUMENT", "output_path no puede ser el mismo archivo de entrada (el original nunca se sobrescribe).",
                       "Elija otro nombre o deje output_path vacío.")
    if not target.parent.exists():
        try:
            target.parent.mkdir(parents=True, exist_ok=True)
        except OSError as e:
            raise ApaError("PERMISSION_DENIED", f"No se pudo crear la carpeta de salida {target.parent}.", "Elija otra carpeta.") from e
    if target.exists() and not overwrite:
        m = re.match(r"^(.*)-(\d+)$", target.stem)
        base = m.group(1) if m else target.stem
        n = 2
        while target.with_name(f"{base}-{n}.docx").exists():
            n += 1
        target = target.with_name(f"{base}-{n}.docx")
    return str(target)


def verify_package(path: str) -> None:
    """El archivo generado debe ser un .docx íntegro: ZIP válido, XML bien formado y reabrible."""
    try:
        with zipfile.ZipFile(path) as z:
            if z.testzip() is not None:
                raise ApaError("OUTPUT_INVALID", "El archivo generado quedó dañado (ZIP).", "No se escribió nada; el original está intacto.")
            for name in z.namelist():
                if name.endswith((".xml", ".rels")):
                    etree.fromstring(z.read(name))
        Document(path)
    except ApaError:
        raise
    except Exception as e:  # noqa: BLE001
        raise ApaError("OUTPUT_INVALID", f"El archivo generado no pasó la verificación ({type(e).__name__}: {e}).",
                       "No se escribió nada; el original está intacto.") from e


def save_verified(doc, out: str) -> str:
    """Guarda en un temporal junto al destino, verifica y reemplaza de forma atómica."""
    target = Path(out)
    tmp = target.with_name(f".{target.stem}.{os.getpid()}.tmp.docx")
    try:
        try:
            doc.save(str(tmp))
        except PermissionError as e:
            raise ApaError("PERMISSION_DENIED", f"No se puede escribir en {target.parent}.",
                           "Elija otra carpeta en output_path.") from e
        verify_package(str(tmp))
        os.replace(tmp, target)
    finally:
        if tmp.exists():
            try:
                tmp.unlink()
            except OSError:
                pass
    return str(target)
