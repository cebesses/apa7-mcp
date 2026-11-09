from docx import Document

from apa7_mcp import pipeline, server
from apa7_mcp.wordfmt import _ref_key


def _rules(path):
    return {i["rule"] for i in server.apa7_check(file_path=str(path))["issues"]}


def test_same_author_references_by_date():
    refs = ["Pérez, J. (2021). B.", "Pérez, J. (s.f.). C.", "Pérez, J. (2019b). A.", "Pérez, J. (2019a). Z.",
            "Pérez, J. (en prensa). X.", "Alba, R. (2000). M."]
    assert sorted(refs, key=_ref_key) == ["Alba, R. (2000). M.", "Pérez, J. (s.f.). C.", "Pérez, J. (2019a). Z.",
                                          "Pérez, J. (2019b). A.", "Pérez, J. (2021). B.", "Pérez, J. (en prensa). X."]


def test_uncited_table_is_flagged(tmp_path):
    f = tmp_path / "t.docx"
    d = Document()
    d.add_paragraph("Este es el cuerpo del trabajo y no menciona ningún cuadro en especial.")
    d.add_paragraph("Tabla 1")
    d.add_paragraph("Resultados")
    d.add_paragraph("Figura 1")
    d.add_paragraph("Como muestra la Figura 1, el efecto creció de manera sostenida.")
    d.save(f)
    r = server.apa7_check(file_path=str(f), detail="full")
    msg = next(i["message"] for i in r["formatIssues"] if i["ruleId"] == "tab-005")
    assert "Tabla 1" in msg and "Figura 1" not in msg


def test_author_note_on_professional_title_page(tmp_path):
    out = tmp_path / "p.docx"
    pipeline.build_document(str(out), [{"type": "title_page", "title": "Estudio", "authors": "Ana Pérez",
                                        "author_note": ["Ana Pérez https://orcid.org/0000-0000-0000-0000", "Correspondencia: ana@x.org"]}],
                            language="es")
    texts = [p.text for p in Document(str(out)).paragraphs]
    assert "Nota del autor" in texts and "Correspondencia: ana@x.org" in texts
