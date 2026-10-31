"""Robustez: entradas inválidas, original intacto, idempotencia, estructuras raras, fuzz y rendimiento."""
import hashlib
import time
import zipfile
from pathlib import Path

import corpus
import pytest
from docx import Document
from lxml import etree

from apa7_mcp import integrity, pipeline, server
from apa7_mcp.errors import ApaError
from apa7_mcp.scan import tracked_count


def sha(p) -> str:
    return hashlib.sha256(Path(p).read_bytes()).hexdigest()


def canon(path: str) -> dict:
    """XML canónico de las partes principales (para comparar dos ejecuciones)."""
    out = {}
    with zipfile.ZipFile(path) as z:
        for n in z.namelist():
            if n.endswith(".xml") and (n.startswith("word/") and "/_rels" not in n):
                out[n] = etree.tostring(etree.fromstring(z.read(n)), method="c14n")
    return out


# ------------------------------------------------------------------ entradas inválidas


def test_missing_file(tmp_path):
    r = server.apa7_apply_apa(file_path=str(tmp_path / "no_existe.docx"))
    assert r["ok"] is False and r["error"]["code"] == "FILE_NOT_FOUND" and r["error"]["hint"]


@pytest.mark.parametrize("ext,code", [(".doc", "UNSUPPORTED_FORMAT"), (".pdf", "UNSUPPORTED_FORMAT"), (".txt", "UNSUPPORTED_FORMAT"),
                                      (".xyz", "UNSUPPORTED_FORMAT")])
def test_unsupported_extensions(tmp_path, ext, code):
    f = tmp_path / f"a{ext}"
    f.write_bytes(b"hola")
    r = server.apa7_validate_document(file_path=str(f))
    assert r["ok"] is False and r["error"]["code"] == code


def test_corrupt_docx(tmp_path):
    f = tmp_path / "roto.docx"
    f.write_bytes(b"esto no es un zip")
    r = server.apa7_apply_apa(file_path=str(f))
    assert r["ok"] is False and r["error"]["code"] == "CORRUPT_DOCX"


def test_ole_header_is_reported_as_encrypted_or_legacy(tmp_path):
    f = tmp_path / "cifrado.docx"
    f.write_bytes(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + b"0" * 100)
    r = server.apa7_apply_apa(file_path=str(f))
    assert r["error"]["code"] == "ENCRYPTED_OR_LEGACY"


def test_zip_without_document_xml(tmp_path):
    f = tmp_path / "zip.docx"
    with zipfile.ZipFile(f, "w") as z:
        z.writestr("hola.txt", "x")
    assert server.apa7_apply_apa(file_path=str(f))["error"]["code"] == "NOT_DOCX"


def test_directory_and_empty_path(tmp_path):
    assert server.apa7_apply_apa(file_path=str(tmp_path))["ok"] is False
    assert server.apa7_apply_apa(file_path="")["error"]["code"] == "INVALID_ARGUMENT"


def test_quoted_path_is_accepted(tmp_path):
    f = corpus.plain_bad(tmp_path / "con espacios.docx")
    r = server.apa7_validate_document(file_path=f'"{f}"')
    assert r["ok"] is True


def test_bad_blocks_and_args(tmp_path):
    f = corpus.plain_bad(tmp_path / "a.docx")
    assert server.apa7_insert_blocks(file_path=str(f), blocks=[{"type": "table", "title": "x"}])["error"]["code"] == "INVALID_ARGUMENT"
    assert server.apa7_insert_blocks(file_path=str(f), blocks=[{"type": "nada"}])["ok"] is False
    assert server.apa7_insert_blocks(file_path=str(f), blocks=[{"type": "paragraph", "text": "x"}], where="mal")["ok"] is False
    assert server.apa7_replace_text(file_path=str(f), replacements=[{"find": "", "replace": "x"}])["ok"] is False
    assert server.apa7_replace_references(file_path=str(f), entries=[])["ok"] is False
    assert server.apa7_build_document(output_path=str(tmp_path / "b.docx"), blocks=[])["ok"] is False
    r = server.apa7_insert_blocks(file_path=str(f), blocks=[{"type": "paragraph", "text": "x"}], where="after_text", text="NO EXISTE")
    assert r["ok"] is False


# ------------------------------------------------------------------ el original y la salida


def test_original_untouched_and_output_collision(tmp_path):
    f = corpus.plain_bad(tmp_path / "t.docx")
    h = sha(f)
    r1 = pipeline.apply_apa(str(f))
    r2 = pipeline.apply_apa(str(f))  # la salida ya existe → nombre-2
    r3 = pipeline.apply_apa(str(f), overwrite=True)
    assert sha(f) == h
    assert Path(r1["outputPath"]).name == "t.apa7.docx" and Path(r2["outputPath"]).name == "t.apa7-2.docx"
    assert r3["outputPath"] == r1["outputPath"]
    assert not list(tmp_path.glob(".*tmp*")), "no deben quedar temporales"


def test_output_equals_input_is_refused(tmp_path):
    f = corpus.plain_bad(tmp_path / "t.docx")
    r = server.apa7_apply_apa(file_path=str(f), output_path=str(f))
    assert r["ok"] is False and sha(f)


def test_dry_run_writes_nothing(tmp_path):
    f = corpus.plain_bad(tmp_path / "t.docx")
    r = pipeline.apply_apa(str(f), dry_run=True)
    assert r["dry_run"] and r["outputPath"] is None and r["changes"]
    assert [p.name for p in tmp_path.iterdir()] == ["t.docx"]


# ------------------------------------------------------------------ idempotencia


@pytest.mark.parametrize("maker", ["plain_bad", "sdt_doc", "two_sections_titlepg", "with_toc", "theme_fonts", "nested_tables"])
def test_idempotent(tmp_path, maker):
    src = getattr(corpus, maker)(tmp_path / "s.docx")
    r1 = pipeline.apply_apa(str(src), str(tmp_path / "o1.docx"))
    r2 = pipeline.apply_apa(r1["outputPath"], str(tmp_path / "o2.docx"))
    a, b = canon(r1["outputPath"]), canon(r2["outputPath"])
    diff = [k for k in a if a[k] != b.get(k)]
    assert not diff, f"la segunda ejecución cambió: {diff}"


# ------------------------------------------------------------------ estructuras reales


def test_tracked_changes_need_consent(tmp_path):
    f = corpus.tracked_changes(tmp_path / "c.docx")
    r = server.apa7_apply_apa(file_path=str(f))
    assert r["ok"] is False and r["error"]["code"] == "TRACKED_CHANGES"
    ok = server.apa7_apply_apa(file_path=str(f), accept_tracked_changes=True)
    assert ok["ok"] is True
    doc = Document(ok["outputPath"])
    text = "\n".join(p.text for p in doc.paragraphs)
    assert "Texto insertado" in text and "Texto borrado" not in text and tracked_count(doc) == 0


def test_content_controls_are_formatted(tmp_path):
    f = corpus.sdt_doc(tmp_path / "s.docx")
    struct = pipeline.read_structure(str(f))
    texts = [p["text"] for p in struct["paragraphs"]]
    assert any("Método" in t for t in texts) and any("Dentro del control" in t for t in texts)
    r = pipeline.apply_apa(str(f))
    doc = Document(r["outputPath"])
    inner = next(p for p in doc.element.body.iter("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p")
                 if "Dentro del control" in "".join(t.text or "" for t in p.iter("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t")))
    assert b'w:firstLine="720"' in etree.tostring(inner)


def test_sections_and_first_page_numbers(tmp_path):
    f = corpus.two_sections_titlepg(tmp_path / "s.docx")
    r = pipeline.apply_apa(str(f))
    doc = Document(r["outputPath"])
    assert pipeline.read_structure(r["outputPath"])["has_page_numbers"] is True
    assert all(not s.different_first_page_header_footer for s in doc.sections)
    assert "PAGE" in doc.sections[0].header._element.xml


def test_existing_toc_is_not_reformatted_nor_duplicated(tmp_path):
    f = corpus.with_toc(tmp_path / "t.docx")
    before = Document(str(f))
    toc_xml = [etree.tostring(p._p) for p in before.paragraphs[:3]]
    r = pipeline.apply_apa(str(f), toc=True)
    assert any("Ya existe una tabla de contenido" in c for c in r["changes"])
    after = Document(r["outputPath"])
    assert sum("instrText" in etree.tostring(p._p).decode() and "TOC" in etree.tostring(p._p).decode() for p in after.paragraphs) >= 1
    assert b'w:firstLine="720"' not in b"".join(etree.tostring(p._p) for p in after.paragraphs[:3]), "el TOC no debe llevar sangría de párrafo"
    assert toc_xml  # el TOC existe antes y después


def test_nested_tables_and_empty_and_theme_fonts(tmp_path):
    for maker in ("nested_tables", "empty_doc", "theme_fonts"):
        f = getattr(corpus, maker)(tmp_path / f"{maker}.docx")
        r = server.apa7_apply_apa(file_path=str(f))
        assert r["ok"] is True, (maker, r)
        Document(r["outputPath"])  # reabre


def test_nested_tables_preserved(tmp_path):
    f = corpus.nested_tables(tmp_path / "n.docx")
    snap = integrity.snapshot(Document(str(f)))
    r = pipeline.apply_apa(str(f))
    assert integrity.snapshot(Document(r["outputPath"]))["tables"] == snap["tables"] == 2


# ------------------------------------------------------------------ guardia de integridad


def test_integrity_problems():
    a = {"words": 1000, "tables": 2, "images": 1, "hyperlinks": 0, "equations": 0, "footnote_refs": 0, "bookmarks": 0, "sections": 1}
    assert integrity.problems(a, dict(a)) == []
    assert integrity.problems(a, {**a, "words": 990}) == []  # pérdida mínima tolerada
    assert integrity.problems(a, {**a, "words": 400})
    assert integrity.problems(a, {**a, "tables": 1})
    assert integrity.problems(a, {**a, "images": 0})
    with pytest.raises(ApaError):
        integrity.guard(a, Document())


# ------------------------------------------------------------------ fuzz y rendimiento


@pytest.mark.parametrize("seed", range(40))
def test_fuzz_never_fails_and_output_opens(tmp_path, seed):
    f = corpus.fuzz_doc(tmp_path / "f.docx", seed)
    r = server.apa7_apply_apa(file_path=str(f))
    assert r["ok"] is True, r
    Document(r["outputPath"])
    v = server.apa7_validate_document(file_path=r["outputPath"])
    assert v["ok"] is True
    ex = server.apa7_extract_text(file_path=r["outputPath"])
    assert ex["ok"] is True


def test_big_document_speed(tmp_path):
    f = corpus.big_doc(tmp_path / "big.docx", 3000)
    t = time.perf_counter()
    r = pipeline.apply_apa(str(f))
    dt = time.perf_counter() - t
    assert r["outputPath"] and dt < 20, f"{dt:.1f}s"


def test_doctor_reports_ok():
    r = server.apa7_doctor(check_network=False)
    assert r["ok"] is True and r["checks"]
