"""Fase 4: documentos en inglés (reglas, formato, bloques) y que el español no se contamine."""
from docx import Document

from apa7_mcp import analyze, lang as L, options as O, pipeline, server

BODY = ("Online learning has grown quickly in recent years and it is now part of the way that students study in many countries "
        "around the world. ")


def make_en(tmp_path, name="en.docx"):
    d = Document()
    d.add_paragraph().add_run("Introduction").bold = True
    d.add_paragraph(BODY + "Previous work found clear benefits (Smith and Jones, 2019, pp. 23-25) and later studies agreed (Smith et al., 2019; Lee, n.d.).")
    d.add_paragraph(BODY + "Brown & Lee (2020) argued that the effect depends on the course design and on the support students receive.")
    d.add_paragraph().add_run("Bibliography").bold = True
    for t in [
        "Smith, J. & Jones, M. (2019). Learning Online For Everyone (2 ed.). Publisher. doi:10.1000/abc.",
        "Brown, T., and Lee, K. (2020). Course design and student success. Journal of Learning, 3(2), 10-20. Retrieved from https://doi.org/10.1000/xyz",
        "Lee, K. (n.d.). A website about learning. https://example.com/page.",
    ]:
        d.add_paragraph(t)
    f = tmp_path / name
    d.save(f)
    return f


def texts(path):
    return [p.text for p in Document(path).paragraphs]


def test_language_is_autodetected(tmp_path):
    f = make_en(tmp_path)
    rep = analyze.analyze_text("\n".join(texts(f)))
    assert not [i for i in rep["citationIssues"] if i["ruleId"] == "cit-007"], rep["citationIssues"]
    assert L.detect(texts(f)) == "en"
    assert L.detect(["El efecto del método en los estudiantes de la universidad"]) == "es"


def test_english_document_is_fixed_with_english_rules(tmp_path):
    f = make_en(tmp_path)
    r = pipeline.apply_apa(str(f))
    t = texts(r["outputPath"])
    full = "\n".join(t)
    assert "References" in t and "Bibliography" not in t and "Referencias" not in full
    assert "(Smith & Jones, 2019, pp. 23–25)" in full       # & entre paréntesis y rango con raya
    assert "Brown and Lee (2020)" in full                     # and en narrativa
    assert "(Lee, n.d.; Smith et al., 2019)" in full  # citas múltiples en orden alfabético
    assert "Smith, J., & Jones, M." in full                   # coma + & en la lista
    assert "Brown, T., & Lee, K." in full                     # and → &
    assert "(2nd ed.)" in full and "(2.ª ed.)" not in full
    assert "https://doi.org/10.1000/abc" in full and "doi:" not in full
    assert "Retrieved from" not in full
    assert "10–20" in full
    assert not any(x.rstrip().endswith(".com/page.") for x in t)
    assert " y " not in full.replace("Lee, K. (n.d.)", "")    # nada se tradujo al español
    ids = {i["ruleId"] for i in r["needs_agent_action"]}
    assert "cit-007" not in ids and "ref-007" not in ids, r["needs_agent_action"]


def test_english_italics_and_order(tmp_path):
    f = make_en(tmp_path)
    r = pipeline.apply_apa(str(f))
    d = Document(r["outputPath"])
    it = [run.text for p in d.paragraphs for run in p.runs if run.italic and run.text.strip()]
    assert "Journal of Learning" in it and "3" in it and "A website about learning" in it
    refs = [x for x in texts(r["outputPath"]) if x.startswith(("Smith, J.", "Brown, T.", "Lee, K."))]
    assert refs == sorted(refs)


def test_english_is_idempotent(tmp_path):
    f = make_en(tmp_path)
    r1 = pipeline.apply_apa(str(f), str(tmp_path / "o1.docx"))
    r2 = pipeline.apply_apa(r1["outputPath"], str(tmp_path / "o2.docx"))
    assert texts(r1["outputPath"]) == texts(r2["outputPath"])


def test_explicit_language_overrides_detection(tmp_path):
    d = Document()
    d.add_paragraph("Bibliografía")
    d.add_paragraph("Soto, M. (2020). Un libro. Editorial.")
    f = tmp_path / "x.docx"
    d.save(f)
    r = pipeline.apply_apa(str(f), language="en")
    assert "References" in texts(r["outputPath"])
    r2 = pipeline.apply_apa(str(f), language="es", overwrite=False)
    assert "Referencias" in texts(r2["outputPath"])
    bad = server.apa7_apply_apa(file_path=str(f), language="fr")
    assert bad["ok"] is False and bad["error"]["code"] == "INVALID_ARGUMENT"


def test_english_blocks_use_english_labels(tmp_path):
    f = make_en(tmp_path)
    r = pipeline.insert_blocks(str(f), str(tmp_path / "b.docx"), [
        {"type": "table", "title": "Results by group", "header": ["Group", "Mean"], "rows": [["A", 4.2]], "note": "SD = standard deviation."},
        {"type": "list_of_tables"},
    ], where="before_references")
    full = "\n".join(texts(r["outputPath"]))
    assert "Table 1" in full and "Tabla" not in full and "Note. SD" in full and "List of Tables" in full


def test_build_document_in_english(tmp_path):
    out = tmp_path / "n.docx"
    pipeline.build_document(str(out), [
        {"type": "toc"},
        {"type": "heading", "level": 1, "text": "Introduction"},
        {"type": "paragraph", "text": "The purpose of this study is to examine the effects of the course design on the results of the students."},
        {"type": "figure", "title": "Mean score by group", "chart": {"type": "bar", "labels": ["A", "B"], "series": [{"name": "M", "values": [1, 2]}]}},
        {"type": "references", "entries": ["Smith, J., & Jones, M. (2019). Learning online. *Journal of Learning*, *3*(2), 1–9."]},
    ])
    full = "\n".join(texts(out))
    assert "Table of Contents" in full and "References" in full and "Figure 1" in full and "Referencias" not in full


def test_english_reference_formatting():
    from apa7_mcp import rules

    ref = rules.format_reference({"type": "articulo_revista", "authors": ["Smith", "Jones"], "initials": ["J.", "M."], "year": "2019",
                                  "title": "Learning online", "journal": "Journal of Learning", "volume": "3", "number": "2",
                                  "pages": "1–9", "url": "https://doi.org/10.1000/x"}, lang="en")["text"]
    assert ref.startswith("Smith, J., & Jones, M. (2019).")
    es = rules.format_reference({"type": "libro", "authors": ["Soto", "Pérez"], "initials": ["M.", "J."], "title": "Un libro", "publisher": "Ed."},
                                lang="es")["text"]
    assert es.startswith("Soto, M. y Pérez, J. (s.f.).")
    nd = rules.format_reference({"type": "libro", "authors": ["Lee"], "initials": ["K."], "title": "A book", "publisher": "Pub"}, lang="en")["text"]
    assert "(n.d.)" in nd


def test_english_validators():
    from apa7_mcp import rules

    with O.use_options(O.Options(language="en")):
        L.resolve(["x"])
        ids = {i["ruleId"] for i in rules.validate_reference("Smith, J. & Jones, M. (2019). Learning Online For Everyone Today (2 ed.). Publisher.")}
        assert {"ref-003", "ref-015", "ref-016"} <= ids
        assert not rules.validate_reference("Smith, J., & Jones, M. (2019). Learning online for everyone (2nd ed.). Publisher.")
        cit = [i["ruleId"] for i in rules.detect_citations("See (Smith and Jones, 2019).", "en")[1]]
        assert "cit-002" in cit
        assert not [i for i in rules.detect_citations("See (Smith & Jones, 2019).", "en")[1] if i["ruleId"] == "cit-002"]
        assert rules.detect_narrative_amp("Smith & Jones (2019) said", "en")[0]["suggestion"].startswith("Smith and Jones")
