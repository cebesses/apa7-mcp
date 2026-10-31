from docx import Document
from docx.shared import Pt

from apa7_mcp import pipeline


def _messy(path):
    d = Document()
    d.add_paragraph("Un título de ejemplo")
    for _ in range(4):
        d.add_paragraph("")
    d.add_paragraph("Carlos Mendoza Rivera")
    d.add_paragraph("Universidad Ejemplo")
    d.add_page_break()
    h = d.add_heading("Uso de números", 2)
    for r in h.runs:
        r.bold = False
        r.italic = True
    d.add_paragraph("Texto de prueba suficientemente largo para contar como cuerpo del trabajo académico aquí.")
    d.add_paragraph("Palabras clave: apa, norma")
    d.add_heading("Referencias", 1)
    r = d.add_paragraph("Pérez, J. (2020). Un libro. Editorial.")
    r.paragraph_format.space_after = Pt(10)
    d.save(path)


def test_validator_flags_then_apply_fixes(tmp_path):
    f = tmp_path / "s.docx"
    _messy(f)
    before = str(pipeline.validate_document(str(f)))
    for rid in ("est-008", "tit-005", "abs-002", "ref-019"):
        assert rid in before, rid
    r = pipeline.apply_apa(str(f))
    after = str(pipeline.validate_document(r["outputPath"]))
    for rid in ("est-008", "tit-005", "abs-002", "ref-019"):
        assert rid not in after, rid
