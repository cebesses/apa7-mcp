"""Perfiles/opciones: fuente, tamaño de página, interlineado, perfil junto al documento, errores y ausencia de fugas."""
import json

import corpus
from docx import Document
from docx.shared import Cm, Pt

from apa7_mcp import options as O
from apa7_mcp import pipeline, server


def test_a4_profile_forces_a4_even_from_letter(tmp_path):
    d = Document()
    d.add_paragraph("Texto de prueba suficientemente largo para ser un párrafo de cuerpo en la prueba del perfil.")
    f = tmp_path / "l.docx"
    d.save(f)
    r = pipeline.apply_apa(str(f), profile="a4")
    s = Document(r["outputPath"]).sections[0]
    assert abs(s.page_width.cm - 21.0) < 0.05 and abs(s.page_height.cm - 29.7) < 0.05


def test_letter_profile_overrides_a4(tmp_path):
    f = corpus.plain_bad(tmp_path / "a4.docx")  # el documento es A4
    r = pipeline.apply_apa(str(f), profile="apa7")
    assert abs(Document(r["outputPath"]).sections[0].page_width.cm - 21.59) < 0.05


def test_forced_font_and_spacing_from_dict(tmp_path):
    f = corpus.plain_bad(tmp_path / "x.docx")
    r = pipeline.apply_apa(str(f), profile={"font": "Arial", "size": 11, "line_spacing": 1.5, "margin_cm": 2.0})
    d = Document(r["outputPath"])
    assert d.styles["Normal"].font.name == "Arial" and d.styles["Normal"].font.size == Pt(11)
    assert abs(d.sections[0].left_margin - Cm(2.0)) < 1000  # Word guarda twips: ~0.0001 cm de redondeo
    assert d.styles["Normal"].paragraph_format.line_spacing == 1.5
    v = pipeline.validate_document(r["outputPath"], profile={"font": "Arial", "size": 11, "line_spacing": 1.5, "margin_cm": 2.0})
    assert not [i for i in v["remaining"] if i["ruleId"] in ("fmt-002", "fmt-003", "fmt-004", "fmt-005")]


def test_profile_file_next_to_document(tmp_path):
    f = corpus.plain_bad(tmp_path / "x.docx")
    (tmp_path / "apa7.profile.json").write_text(json.dumps({"extends": "a4", "font": "Georgia", "size": 11}), encoding="utf-8")
    r = pipeline.apply_apa(str(f))
    assert Document(r["outputPath"]).styles["Normal"].font.name == "Georgia"


def test_invalid_profiles_are_structured_errors(tmp_path):
    f = corpus.plain_bad(tmp_path / "x.docx")
    for prof in ("no_existe", {"fuente": "x"}, {"page_size": "A3"}, {"line_spacing": -1}, {"language": "fr"}):
        r = server.apa7_apply_apa(file_path=str(f), profile=prof)
        assert r["ok"] is False and r["error"]["code"] == "INVALID_PROFILE", (prof, r)


def test_no_state_leak_between_calls(tmp_path):
    f = corpus.plain_bad(tmp_path / "x.docx")
    before = dict(O.TARGET)
    pipeline.apply_apa(str(f), profile={"font": "Arial", "size": 11}, overwrite=True)
    assert O.TARGET == before and O.active() == O.Options()
    # una llamada sin perfil después de otra con perfil no hereda la fuente
    r = pipeline.apply_apa(str(f), overwrite=True)
    assert Document(r["outputPath"]).styles["Normal"].font.name in ("Calibri", "Times New Roman")


def test_numbered_headings_flag_silences_warning(tmp_path):
    d = Document()
    d.add_paragraph("1. Introducción", style="Heading 1")
    d.add_paragraph("Texto largo para que el documento tenga cuerpo y sea analizado correctamente en esta prueba.")
    f = tmp_path / "n.docx"
    d.save(f)
    base = pipeline.validate_document(str(f))
    assert base["ok"] if "ok" in base else True
    ids = lambda r: {i["ruleId"] for i in server.apa7_analyze_document(file_path=str(f), profile=r)["formatIssues"]}  # noqa: E731
    assert "tit-003" in ids(None) and "tit-003" not in ids({"numbered_headings": True})
