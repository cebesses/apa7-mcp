"""Prueba de humo: Word con errores → apa7_apply_apa → validación; construcción e inserción de bloques."""
import sys
import tempfile
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm, Pt

from apa7_mcp import pipeline

tmp = Path(tempfile.mkdtemp(prefix="apa7_"))

# ---- 1. Documento con errores típicos
d = Document()
s = d.sections[0]
s.page_width, s.page_height = Cm(21), Cm(29.7)
s.left_margin = s.right_margin = s.top_margin = s.bottom_margin = Cm(1.5)
d.styles["Normal"].font.name = "Calibri"
d.styles["Normal"].font.size = Pt(11)


def para(text, bold=False, justify=True, center=False):
    p = d.add_paragraph()
    r = p.add_run(text)
    r.bold = bold
    if center:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    elif justify:
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    return p


para("Efectos del aprendizaje autónomo", bold=True, center=True)
para("Introducción", bold=True, justify=False)
for _ in range(2):
    para("El aprendizaje autónomo mejora el rendimiento académico según diversos estudios recientes (Pérez & Soto, 2019, p. 23-25). "
         "Además, se ha observado que reduce la deserción en la educación superior (Gómez, 2020).")
para("Método", bold=True, justify=False)
para("Participaron estudiantes de dos grupos y los resultados se resumen en la siguiente tabla, que compara las medias de ambos.")
t = d.add_table(rows=3, cols=3)
t.style = "Table Grid"
for i, row in enumerate([["Grupo", "Media", "DE"], ["A", "4.2", "0.8"], ["B", "3.7", "0.9"]]):
    for j, v in enumerate(row):
        t.cell(i, j).text = v
para("Referencias", bold=True, justify=False)
para("Soto, M. (2020). Deserción universitaria. Bogotá: Editorial Uno.")
para("Pérez, J. & Soto, M. (2019). Aprendizaje autónomo. Revista de Educación, 12(3), 20-30. Recuperado de https://doi.org/10.1000/xyz")
bad = tmp / "malo.docx"
d.save(bad)

before = pipeline.validate_document(str(bad))
print("ANTES  :", before["summary"], "| errores:", before["errors"], "advertencias:", before["warnings"])

r = pipeline.apply_apa(str(bad), running_head="Aprendizaje autónomo")
print("CAMBIOS:")
for c in r["changes"]:
    print("  -", c)
print("DESPUÉS:", r["after"]["summary"], "| passed:", r["passed"])
print("PENDIENTE (agente):")
for i in r["needs_agent_action"]:
    print("  -", i["ruleId"], i["message"][:100])

# ---- 2. Comprobaciones
out = Document(r["outputPath"])
sec = out.sections[0]
assert abs(sec.left_margin.cm - 2.54) < 0.05, "margen"
assert abs(sec.page_width.cm - 21.0) < 0.1, "el documento era A4: debe conservarse (A4 es válido)"
assert r["after"]["score"] > before["score"], "score no mejoró"
refs = [p.text for p in out.paragraphs if p.text.startswith(("Soto", "Pérez, J"))]
assert refs and refs[0].startswith("Pérez"), f"orden alfabético: {refs}"
assert "Recuperado" not in " ".join(refs), "Recuperado de"
assert pipeline.read_structure(r["outputPath"])["has_page_numbers"], "número de página"

# ---- 3. Construir desde cero con todo
blocks = [
    {"type": "title_page", "title": "Efectos del aprendizaje autónomo", "authors": "Ana Pérez", "affiliation": "Universidad Nacional",
     "course": "PSI-101", "instructor": "Dr. Gómez", "due_date": "28 de octubre de 2026"},
    {"type": "toc"}, {"type": "list_of_tables"}, {"type": "list_of_figures"}, {"type": "page_break"},
    {"type": "title", "text": "Efectos del aprendizaje autónomo"},
    {"type": "heading", "level": 1, "text": "Introducción"},
    {"type": "paragraph", "text": "El aprendizaje autónomo mejora el rendimiento (Pérez, 2019, p. 23)."},
    {"type": "heading", "level": 2, "text": "Método"},
    {"type": "list", "style": "numbered", "items": ["Selección", "Cuestionario", "Análisis"]},
    {"type": "list", "style": "bullet", "items": ["Grupo A", "Grupo B"]},
    {"type": "table", "title": "Resultados por grupo", "header": ["Grupo", "Media", "DE"], "rows": [["A", 4.2, 0.8], ["B", 3.7, 0.9]],
     "note": "DE = desviación estándar."},
    {"type": "figure", "title": "Rendimiento por grupo", "chart": {"type": "bar", "labels": ["A", "B"],
     "series": [{"name": "Media", "values": [4.2, 3.7]}], "x_label": "Grupo", "y_label": "Media"}, "note": "Datos del estudio."},
    {"type": "figure", "title": "Distribución", "chart": {"type": "pie", "labels": ["Sí", "No"], "series": [{"name": "r", "values": [70, 30]}]}},
    {"type": "references", "entries": ["Soto, M. (2020). *Deserción universitaria*. Editorial Uno.",
                                       "Pérez, J. (2019). Aprendizaje autónomo. *Revista de Educación*, *12*(3), 20-30. https://doi.org/10.1000/xyz"]},
]
built = tmp / "nuevo.docx"
b = pipeline.build_document(str(built), blocks)
print("BUILD  :", b)
v = pipeline.validate_document(str(built))
print("VALID  :", v["summary"], "| passed:", v["passed"], [i["ruleId"] for i in v["remaining"]])

# ---- 4. Insertar en existente
ins = pipeline.insert_blocks(r["outputPath"], str(tmp / "ins.docx"),
                             [{"type": "figure", "title": "Otra", "chart": {"type": "line", "labels": ["1", "2", "3"], "series": [{"name": "s", "values": [1, 3, 2]}]}}],
                             where="before_references")
print("INSERT :", ins)
doc = Document(ins["outputPath"])
texts = [p.text for p in doc.paragraphs]
fig = next(i for i, t in enumerate(texts) if t.startswith("Figura 1"))
assert fig < texts.index("Referencias"), "figura antes de Referencias"

# ---- 5. replace_references
rr = pipeline.replace_references(ins["outputPath"], str(tmp / "rr.docx"), ["Zeta, A. (2021). *Libro*. Ed.", "Alfa, B. (2020). *Otro*. Ed."])
assert [p.text for p in Document(rr["outputPath"]).paragraphs if p.text.startswith(("Zeta", "Alfa"))][0].startswith("Alfa")
print("OK ->", tmp)
sys.exit(0)
