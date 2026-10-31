from docx import Document

from apa7_mcp import pipeline


def _doc(path):
    d = Document()
    d.add_heading("Método", 1)
    d.add_heading("Participantes", 2)
    d.add_heading("Muestra inicial", 4)
    d.add_paragraph("Se evaluó a veinte personas en total durante el primer semestre del año.")
    d.add_heading("Criterios", 5)
    d.add_paragraph("Se incluyó a quienes aceptaron participar de forma voluntaria en el estudio.")
    d.save(path)


def test_levels_4_and_5_become_inline_and_idempotent(tmp_path):
    f = tmp_path / "a.docx"
    _doc(f)
    r = pipeline.apply_apa(str(f))
    out = Document(r["outputPath"])
    paras = [p for p in out.paragraphs if p.text]
    texts = [p.text for p in paras]
    assert any(t.startswith("Muestra inicial. Se evaluó") for t in texts)
    assert any(t.startswith("Criterios. Se incluyó") for t in texts)
    p4 = next(p for p in paras if p.text.startswith("Muestra inicial."))
    assert p4.runs[0].bold and not p4.runs[0].italic and not p4.runs[-1].bold
    p5 = next(p for p in paras if p.text.startswith("Criterios."))
    assert p5.runs[0].bold and p5.runs[0].italic
    r2 = pipeline.apply_apa(r["outputPath"], overwrite=False)
    assert not any("nivel 4/5" in c for c in r2["changes"])
