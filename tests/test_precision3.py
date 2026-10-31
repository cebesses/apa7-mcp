"""Fase 3: cursiva en referencias, cita en bloque automática, citas secundarias y año ambiguo."""
from docx import Document
from docx.shared import Cm

from apa7_mcp import pipeline, server

LONG_QUOTE = ("la crisis financiera ha sido motivada por lo que hay de más perverso en el mundo capitalista y ha contribuido de manera "
              "notable a un clima general de desconfianza hacia las instituciones financieras y hacia los partidos políticos tradicionales "
              "de derecha en buena parte del mundo desarrollado durante la década siguiente")


def italics(doc):
    return [r.text for p in doc.paragraphs for r in p.runs if r.italic and r.text.strip()]


def make(tmp_path, paragraphs, name="d.docx"):
    d = Document()
    for t in paragraphs:
        d.add_paragraph(t)
    f = tmp_path / name
    d.save(f)
    return f


def test_reference_italics_are_added_and_idempotent(tmp_path):
    f = make(tmp_path, [
        "Introducción",
        "Texto de cuerpo suficientemente largo para la prueba de referencias con cursivas (Pérez, 2019; Soto, 2020; Gómez, 2018; López, 2017).",
        "Referencias",
        "Gómez, L. (2018). Redes sociales y juventud [Tesis de maestría, Universidad Nacional]. Repositorio.",
        "López, A. (2017). Ética en la investigación. En M. Ruiz (Ed.), Manual de investigación (pp. 10–30). Editorial Dos.",
        "Pérez, J. (2019). Aprendizaje autónomo. Revista de Educación, 12(3), 20–30. https://doi.org/10.1000/xyz",
        "Soto, M. (2020). Deserción universitaria en Colombia (2.ª ed.). Editorial Uno.",
    ])
    r = pipeline.apply_apa(str(f))
    it = italics(Document(r["outputPath"]))
    assert "Redes sociales y juventud" in it
    assert "Manual de investigación" in it
    assert "Revista de Educación" in it and "12" in it and "(3)" not in "".join(it)
    assert "Deserción universitaria en Colombia" in it and not any("ed." in t for t in it)
    assert "ref-017" not in {i["ruleId"] for i in r["needs_agent_action"]}
    # segunda pasada: no cambia nada
    r2 = pipeline.apply_apa(r["outputPath"], overwrite=False)
    assert italics(Document(r2["outputPath"])) == it


def test_existing_italics_in_references_are_respected(tmp_path):
    d = Document()
    d.add_paragraph("Referencias").runs[0].bold = True
    p = d.add_paragraph()
    p.add_run("Soto, M. (2020). ")
    p.add_run("Título ya en cursiva").italic = True
    p.add_run(". Editorial Uno.")
    f = tmp_path / "i.docx"
    d.save(f)
    r = pipeline.apply_apa(str(f))
    assert italics(Document(r["outputPath"])) == ["Título ya en cursiva"]


def test_long_quote_becomes_block_quote(tmp_path):
    f = make(tmp_path, [
        "Texto de introducción suficientemente largo para ser un párrafo de cuerpo normal en la prueba.",
        f"“{LONG_QUOTE}” (Lynch, 2012, p. 127)",
        "Cita corta: “una frase breve” (Lynch, 2012, p. 5) dentro del texto del párrafo normal.",
    ])
    r = pipeline.apply_apa(str(f))
    d = Document(r["outputPath"])
    q = next(p for p in d.paragraphs if "crisis financiera" in p.text)
    assert not q.text.lstrip().startswith(("“", '"')) and q.text.rstrip().endswith("(Lynch, 2012, p. 127)")
    assert ". (Lynch" in q.text
    assert abs(q.paragraph_format.left_indent - Cm(1.27)) < 1000 and not q.paragraph_format.first_line_indent
    short = next(p for p in d.paragraphs if "Cita corta" in p.text)
    assert "“una frase breve”" in short.text  # las citas cortas se dejan en el texto
    assert any("block quote" in c.lower() or "cita en bloque" in c for c in r["changes"])
    assert "cit-010" not in {i["ruleId"] for i in r["needs_agent_action"]}


def test_secondary_citation_checks_the_secondary_source():
    from apa7_mcp import analyze

    rep = analyze.analyze_text("Se afirma (Pérez, 2010, citado en Soto, 2020) que sí.\nReferencias\nSoto, M. (2020). Libro. Editorial.")
    assert not [i for i in rep["citationIssues"] if i["ruleId"] == "cit-007"], rep["citationIssues"]
    assert not [i for i in rep["referenceIssues"] if i["ruleId"] == "ref-007"]


def test_same_author_year_needs_letters():
    from apa7_mcp import analyze

    rep = analyze.analyze_text("Texto (Soto, 2020).\nReferencias\nSoto, M. (2020). Uno. Ed.\nSoto, M. (2020). Dos. Ed.")
    assert "ref-018" in {i["ruleId"] for i in rep["referenceIssues"]}
    rep2 = analyze.analyze_text("Texto (Soto, 2020).\nReferencias\nSoto, M. (2020a). Uno. Ed.\nSoto, M. (2020b). Dos. Ed.")
    assert "cit-012" in {i["ruleId"] for i in rep2["citationIssues"]}
    assert "ref-018" not in {i["ruleId"] for i in rep2["referenceIssues"]}


def test_server_returns_new_issue_ids(tmp_path):
    f = make(tmp_path, ["Referencias", "Gómez, L. (2018). Un libro sin cursiva. Editorial."])
    r = server.apa7_analyze_document(file_path=str(f))
    assert "ref-017" in {i["ruleId"] for i in r["formatIssues"]}
