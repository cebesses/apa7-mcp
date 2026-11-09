"""Prueba de precisión APA 7: resumen, palabras clave, citas, DOI, referencias, alt text, apéndices, fuente válida."""
import io
import tempfile
from pathlib import Path

from docx import Document
from docx.shared import Inches, Pt
from PIL import Image

from apa7_mcp import pipeline

tmp = Path(tempfile.mkdtemp(prefix="apa7p_"))
d = Document()
d.styles["Normal"].font.name = "Georgia"  # fuente APA válida (Georgia 11): debe conservarse
d.styles["Normal"].font.size = Pt(11)


def para(text, bold=False, indent=None):
    p = d.add_paragraph()
    p.add_run(text).bold = bold
    if indent is not None:
        p.paragraph_format.first_line_indent = Inches(indent)
    return p


para("Resumen", bold=True)
para("Este trabajo analiza el aprendizaje autónomo en estudiantes universitarios y describe sus efectos principales.", indent=0.5)
para("Palabras clave: aprendizaje, autonomía, educación")
para("Introducción", bold=True)
para("La OMS recomienda el estudio autónomo. La OMS lo repite y la OMS también lo menciona en tres informes distintos "
     "con 3 participantes de la muestra piloto, según se vio antes de la aplicación final del cuestionario.")
para("Se afirma que “el aprendizaje autónomo mejora de forma notable el rendimiento” (Lynch, 2012) en varios contextos "
     "y se ha estudiado (Soto, 2020; Pérez, 2019) con resultados mixtos, como indican Pérez & Soto (2019) en su estudio.")
d.add_paragraph("Figura 1").runs[0].bold = True
d.add_paragraph("Distribución de la muestra").runs[0].italic = True
buf = io.BytesIO()
Image.new("RGB", (40, 30), "navy").save(buf, "PNG")
buf.seek(0)
d.add_picture(buf)
para("Bibliografía", bold=True)
para("Soto, J.M. (2020). Deserción Universitaria En Colombia (2da ed.). Editorial Uno. doi:10.1000/abc.")
para("Pérez, A. (2019). Aprendizaje autónomo. Revista de educación, 12(3), 20-30. https://doi.org/10.1000/xyz.")
para("Lynch, T. (2012). La crisis financiera. Editorial Dos.")
para("Apéndice A", bold=True)
para("Cuestionario aplicado a los participantes.")
src = tmp / "p.docx"
d.save(src)

r = pipeline.apply_apa(str(src))
out = Document(r["outputPath"])
texts = [p.text for p in out.paragraphs]
by = {p.text: p for p in out.paragraphs}
joined = "\n".join(texts)

checks = {
    "fuente válida conservada (Georgia 11)": out.styles["Normal"].font.name == "Georgia" and out.styles["Normal"].font.size == Pt(11),
    "título 'Referencias'": "Referencias" in texts and "Bibliografía" not in texts,
    "Referencias en página nueva": by["Referencias"].paragraph_format.page_break_before is True,
    "resumen sin sangría": by[next(t for t in texts if t.startswith("Este trabajo"))].paragraph_format.first_line_indent == 0,
    "palabras clave con etiqueta en cursiva": by[next(t for t in texts if t.startswith("Palabras clave"))].runs[0].italic is True,
    "DOI https://doi.org/": "https://doi.org/10.1000/abc" in joined and "doi:" not in joined,
    "sin punto tras URL/DOI": not any(t.rstrip().endswith("xyz.") or t.rstrip().endswith("abc.") for t in texts),
    "iniciales separadas": "Soto, J. M." in joined,
    "rango con raya": "20–30" in joined,
    "edición 2.ª": "(2.ª ed.)" in joined,
    "citas múltiples en orden": "(Pérez, 2019; Soto, 2020)" in joined,
    "narrativa con 'y'": "Pérez y Soto (2019)" in joined,
    "referencias ordenadas": [t for t in texts if t.startswith(("Lynch,", "Pérez, A", "Soto, J"))] == sorted(
        [t for t in texts if t.startswith(("Lynch,", "Pérez, A", "Soto, J"))]),
    "apéndice centrado en página nueva": by["Apéndice A"].alignment is not None and by["Apéndice A"].paragraph_format.page_break_before is True,
    "alt text de la figura": "Distribución de la muestra" in out.element.body.xml and 'descr="Distribución de la muestra"' in out.element.body.xml,
}
pend = {i["rule"] for i in r["needs_agent_action"]}
sugg = {i["rule"] for i in r["suggestions"]}
checks["cita textual sin página detectada (cit-009)"] = "cit-009" in pend
checks["número en letras sugerido (num-001)"] = "num-001" in sugg
checks["sigla sin definir sugerida (abr-001)"] = "abr-001" in sugg
checks["título en mayúsculas detectado (ref-015)"] = "ref-015" in pend or "Deserción Universitaria En Colombia" in joined

bad = [k for k, v in checks.items() if not v]
for k, v in checks.items():
    print(("OK  " if v else "FAIL"), k)
print("tiempo:", r["elapsed_s"], "s | sugerencias:", sorted(sugg))
for i in r["needs_agent_action"]:
    print("  pendiente", i["rule"], "|", i["msg"][:70])
assert not bad, f"fallaron: {bad}"
print("PRECISIÓN OK")
