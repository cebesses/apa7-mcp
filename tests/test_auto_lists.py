from docx import Document

from apa7_mcp import pipeline


def _doc(path, n):
    d = Document()
    d.add_paragraph("Título del trabajo")
    d.add_page_break()
    for i in range(1, n + 1):
        d.add_paragraph(f"Como muestra la Tabla {i}, el resultado del grupo fue estable durante todo el estudio.")
        d.add_paragraph(f"Tabla {i}")
        d.add_paragraph(f"Resultados del grupo {i}")
        d.add_table(rows=1, cols=2)
    d.save(path)


def test_list_of_tables_added_from_three(tmp_path):
    _doc(tmp_path / "a.docx", 3)
    r = pipeline.apply_apa(str(tmp_path / "a.docx"))
    xml = Document(r["outputPath"]).element.xml
    assert 'TOC \h \z \c "Tabla"' in xml and xml.count("SEQ Tabla") == 3
    texts = [p.text for p in Document(r["outputPath"]).paragraphs]
    assert texts.index("Índice de tablas") < texts.index("Como muestra la Tabla 1, el resultado del grupo fue estable durante todo el estudio.")
    r2 = pipeline.apply_apa(r["outputPath"])
    assert not any("índice" in c for c in r2["changes"])


def test_no_list_with_two_tables_or_when_disabled(tmp_path):
    _doc(tmp_path / "b.docx", 2)
    assert not any("índice" in c for c in pipeline.apply_apa(str(tmp_path / "b.docx"))["changes"])
    _doc(tmp_path / "c.docx", 4)
    assert not any("índice" in c for c in pipeline.apply_apa(str(tmp_path / "c.docx"), lists=False)["changes"])
