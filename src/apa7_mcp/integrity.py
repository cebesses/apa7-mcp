"""Guardia de integridad: antes de escribir, comprueba que no se perdió contenido del documento."""
from __future__ import annotations

from docx.document import Document as DocumentT

from .errors import ApaError
from .scan import W

WP = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"


def snapshot(doc: DocumentT) -> dict:
    body = doc.element.body
    words = sum(len((t.text or "").split()) for t in body.iter(W + "t"))
    return {
        "words": words,
        "tables": sum(1 for _ in body.iter(W + "tbl")),
        "images": sum(1 for _ in body.iter(WP + "inline", WP + "anchor")),
        "hyperlinks": sum(1 for _ in body.iter(W + "hyperlink")),
        "equations": sum(1 for _ in body.iter("{http://schemas.openxmlformats.org/officeDocument/2006/math}oMath")),
        "footnote_refs": sum(1 for _ in body.iter(W + "footnoteReference")),
        "bookmarks": sum(1 for _ in body.iter(W + "bookmarkStart")),
        "sections": len(doc.sections),
    }


def problems(before: dict, after: dict, *, allow_word_loss: float = 0.02, allow_fewer_words: int = 25,
             check: tuple[str, ...] = ("tables", "images", "hyperlinks", "equations", "footnote_refs", "sections")) -> list[str]:
    """Lista de pérdidas inesperadas. Las cantidades estructurales no pueden disminuir; las palabras toleran una pequeña pérdida
    (por ejemplo al quitar «Recuperado de»)."""
    out = []
    for k in check:
        if after[k] < before[k]:
            out.append(f"{k}: {before[k]} → {after[k]}")
    lost = before["words"] - after["words"]
    if lost > max(allow_fewer_words, allow_word_loss * before["words"]):
        out.append(f"palabras: {before['words']} → {after['words']}")
    return out


def guard(before: dict, doc: DocumentT, **kw) -> None:
    bad = problems(before, snapshot(doc), **kw)
    if bad:
        raise ApaError("INTEGRITY_ABORT",
                       "Se canceló la operación porque habría perdido contenido (" + "; ".join(bad) + ").",
                       "El archivo original no se modificó y no se escribió ninguna salida. Reporte este caso con el documento de ejemplo.")
