"""Errores estructurados y validación de entrada: ninguna herramienta lanza excepciones crudas al agente."""
from __future__ import annotations

import concurrent.futures as cf
import functools
import logging
import os
import sys
import zipfile
from pathlib import Path
from typing import Any, Callable

log = logging.getLogger("apa7")
if not log.handlers:  # nunca stdout: es el canal del protocolo MCP
    _h = logging.StreamHandler(sys.stderr)
    _h.setFormatter(logging.Formatter("apa7 %(levelname)s %(message)s"))
    log.addHandler(_h)
    log.setLevel(logging.INFO)

MAX_MB = int(os.environ.get("APA7_MAX_MB", "200"))
TIMEOUT_S = float(os.environ.get("APA7_TIMEOUT", "180"))


class ApaError(Exception):
    """Error esperado: lleva código, mensaje y una pista de cómo resolverlo."""

    def __init__(self, code: str, message: str, hint: str = ""):
        super().__init__(message)
        self.code, self.message, self.hint = code, message, hint

    def to_dict(self) -> dict:
        return {"ok": False, "error": {"code": self.code, "message": self.message, "hint": self.hint}}


LEGACY = {
    ".doc": "Es un Word antiguo (.doc). Ábralo en Word y use Archivo → Guardar como → Documento de Word (.docx).",
    ".docm": "Es un documento con macros (.docm). Guárdelo como .docx sin macros.",
    ".dotx": "Es una plantilla de Word (.dotx). Cree un documento a partir de ella y guárdelo como .docx.",
    ".odt": "Es un documento de LibreOffice (.odt). Guárdelo como .docx (Archivo → Guardar como → Word).",
    ".rtf": "Es un archivo RTF. Ábralo en Word y guárdelo como .docx.",
    ".pdf": "Es un PDF: no se puede editar como Word. Conviértalo a .docx primero (Word puede abrir PDFs).",
    ".txt": "Es texto plano: use apa7_analyze_text o cree un documento con apa7_build_document.",
}


def clean_path(p: str) -> Path:
    s = str(p).strip().strip('"').strip("'")
    return Path(os.path.expandvars(os.path.expanduser(s)))


def check_zip(path: Path) -> None:
    """Valida que sea un .docx íntegro y no una bomba de descompresión."""
    try:
        with zipfile.ZipFile(path) as z:
            names = set(z.namelist())
            if "word/document.xml" not in names:
                raise ApaError("NOT_DOCX", "El archivo es un ZIP pero no contiene un documento de Word (falta word/document.xml).",
                               "Compruebe que es un .docx real y no otro tipo de archivo renombrado.")
            infos = z.infolist()
            total = sum(i.file_size for i in infos)
            comp = sum(i.compress_size for i in infos) or 1
            if total > 2_000_000_000 or (total / comp > 500 and total > 50_000_000):
                raise ApaError("SUSPICIOUS_ARCHIVE", "El archivo se descomprime a un tamaño anormal; se rechaza por seguridad.",
                               "Si es un documento legítimo, divídalo o reduzca sus imágenes.")
            bad = z.testzip()
            if bad:
                raise ApaError("CORRUPT_DOCX", f"El archivo está dañado (parte «{bad}»).",
                               "Abra el documento en Word y guárdelo de nuevo, o use «Abrir y reparar».")
    except zipfile.BadZipFile as e:
        with open(path, "rb") as f:
            head = f.read(8)
        if head.startswith(b"\xd0\xcf\x11\xe0"):
            raise ApaError("ENCRYPTED_OR_LEGACY",
                           "El archivo parece cifrado con contraseña o es un Word antiguo (formato OLE).",
                           "Quite la contraseña (Archivo → Información → Proteger documento) o guárdelo como .docx.") from e
        raise ApaError("CORRUPT_DOCX", "El archivo no es un .docx válido o está dañado.",
                       "Abra el documento en Word y guárdelo de nuevo como .docx.") from e


def resolve_input(path: str) -> str:
    """Valida y normaliza la ruta de entrada; devuelve la ruta lista para abrir."""
    if not path or not str(path).strip():
        raise ApaError("INVALID_ARGUMENT", "Falta la ruta del archivo (file_path).", "Indique la ruta del .docx.")
    p = clean_path(path)
    if not p.exists():
        raise ApaError("FILE_NOT_FOUND", f"No existe el archivo: {p}",
                       "Compruebe la ruta (use la ruta completa, con comillas si tiene espacios) y que el archivo no se haya movido.")
    if p.is_dir():
        raise ApaError("INVALID_ARGUMENT", f"La ruta es una carpeta, no un archivo: {p}", "Indique la ruta del .docx.")
    ext = p.suffix.lower()
    if ext in LEGACY:
        raise ApaError("UNSUPPORTED_FORMAT", f"Formato no admitido ({ext}).", LEGACY[ext])
    if ext != ".docx":
        raise ApaError("UNSUPPORTED_FORMAT", f"Extensión no admitida ({ext or 'sin extensión'}): se esperaba .docx.", "Guarde el archivo como .docx.")
    if p.stat().st_size > MAX_MB * 1_000_000:
        raise ApaError("FILE_TOO_LARGE", f"El archivo supera {MAX_MB} MB.", "Reduzca el tamaño o suba el límite con APA7_MAX_MB.")
    try:
        with open(p, "rb"):
            pass
    except PermissionError as e:
        raise ApaError("PERMISSION_DENIED", f"No hay permiso para leer {p}.", "Cierre el archivo en Word o cambie sus permisos.") from e
    check_zip(p)
    return str(p)


def classify(exc: BaseException) -> ApaError:
    """Convierte una excepción cualquiera en un error estructurado."""
    if isinstance(exc, ApaError):
        return exc
    if isinstance(exc, zipfile.BadZipFile):
        return ApaError("CORRUPT_DOCX", "El archivo no es un .docx válido o está dañado.", "Guárdelo de nuevo desde Word.")
    if isinstance(exc, FileNotFoundError):
        return ApaError("FILE_NOT_FOUND", f"No se encontró el archivo: {exc.filename or exc}", "Compruebe la ruta.")
    if isinstance(exc, PermissionError):
        return ApaError("PERMISSION_DENIED", f"Sin permiso para leer o escribir: {getattr(exc, 'filename', '') or exc}",
                        "Cierre el archivo si está abierto en Word, o elija otra carpeta de salida (output_path).")
    if isinstance(exc, (ValueError, TypeError)):
        return ApaError("INVALID_ARGUMENT", str(exc), "Revise los parámetros de la herramienta.")
    if isinstance(exc, KeyError):
        return ApaError("INVALID_ARGUMENT", f"Falta o es inválido el campo {exc}.", "Revise la estructura de los parámetros.")
    return ApaError("INTERNAL_ERROR", f"Error inesperado ({type(exc).__name__}): {exc}",
                    "El archivo original no se modificó. Ejecute apa7_doctor y reporte este mensaje.")


def safe_tool(fn: Callable[..., Any]) -> Callable[..., Any]:
    """Decorador: valida/atrapa errores, aplica tiempo máximo y marca ok=true en los resultados dict."""

    @functools.wraps(fn)
    def wrapper(*args: Any, **kwargs: Any) -> Any:
        try:
            with cf.ThreadPoolExecutor(max_workers=1) as ex:
                fut = ex.submit(fn, *args, **kwargs)
                try:
                    res = fut.result(timeout=TIMEOUT_S)
                except cf.TimeoutError:
                    raise ApaError("TIMEOUT", f"La operación superó {TIMEOUT_S:g} s.",
                                   "Documento muy grande: divida el trabajo o suba APA7_TIMEOUT.") from None
            if isinstance(res, dict) and "ok" not in res:
                res = {"ok": True, **res}
            return res
        except BaseException as e:  # noqa: BLE001 - el servidor nunca debe caerse
            if isinstance(e, (KeyboardInterrupt, SystemExit)):
                raise
            err = classify(e)
            if err.code == "INTERNAL_ERROR":
                log.exception("fallo inesperado en %s", fn.__name__)
            else:
                log.info("%s -> %s: %s", fn.__name__, err.code, err.message)
            return err.to_dict()

    return wrapper
