"""Pruebas: estadísticos en cursiva, puntuación, extracción de texto y verificación de fuentes (usa la red)."""
import os
import tempfile
from pathlib import Path

from docx import Document

from apa7_mcp import pipeline
from apa7_mcp.analyze import analyze_document

tmp = Path(tempfile.mkdtemp(prefix="apa7n_"))
d = Document()
d.add_paragraph().add_run("Introducción").bold = True
p = d.add_paragraph()
# el símbolo 'p' queda en un run distinto al resto (como hace Word al editar): debe detectarse igual
for txt in ("Los resultados mostraron diferencias (M = 3.2, SD = 0.8; t(23) = 2.1, ", "p", " < .05) entre los grupos de estudio y el control."):
    p.add_run(txt)
d.add_paragraph("Hubo  doble espacio ,y falta de espacio,después de la coma, con la palabra de de repetida y un (paréntesis sin cerrar.")
d.add_paragraph("Referencias").runs[0].bold = True
d.add_paragraph("Soto, M. (2020). Un libro de prueba. Editorial Uno.")
src = tmp / "n.docx"
d.save(src)

before = analyze_document(str(src))
ids_before = {i["ruleId"] for i in before["formatIssues"]}
assert "stat-001" in ids_before, f"no detectó estadísticos sin cursiva: {ids_before}"

r = pipeline.apply_apa(str(src))
out = Document(r["outputPath"])
ital = {}
for par in out.paragraphs:
    for run in par.runs:
        if run.italic and run.text.strip():
            ital[run.text] = True
text = "\n".join(par.text for par in out.paragraphs)
checks = {
    "M en cursiva": "M" in ital, "SD en cursiva": "SD" in ital, "t en cursiva": "t" in ital, "p (run aparte) en cursiva": "p" in ital,
    "texto intacto": "(M = 3.2, SD = 0.8; t(23) = 2.1, p < .05)" in text,
    "doble espacio corregido": "Hubo doble" in text,
    "coma mal ubicada corregida (' ,y' → ', y')": "espacio, y falta" in text,
    "espacio tras coma agregado": "coma, con" in text and "espacio, después" in text,
}
after_ids = {i["rule"] for i in r["needs_agent_action"]}
checks["stat-001 resuelto"] = "stat-001" not in after_ids
checks["palabra repetida avisada (txt-001)"] = "txt-001" in after_ids
checks["paréntesis sin cerrar avisado (txt-005)"] = "txt-005" in after_ids

ex = pipeline.extract_text(r["outputPath"])
secs = {x["section"] for x in ex["paragraphs"]}
checks["extract_text con secciones"] = {"cuerpo", "referencias"} <= secs and ex["total"] >= 4

no = pipeline.verify_sources(str(src), None, False)
checks["verificación exige confirmación"] = no["verificado"] is False and "confirm_external" in no["mensaje"]
try:
    if os.environ.get("APA7_OFFLINE"):
        raise RuntimeError("APA7_OFFLINE: sin red")
    v = pipeline.verify_sources(None, [
        "LeCun, Y., Bengio, Y., y Hinton, G. (2015). Deep learning. *Nature*, *521*(7553), 436–444. https://doi.org/10.1038/nature14539",
        "Pérez, J. (2019). Un artículo inventado. *Revista Imaginaria*, 3(1), 1–9. https://doi.org/10.1234/fake.9999"], True)
    st = [x["status"] for x in v["results"]]
    if "sin_conexion" in st:
        print("(sin red: se omite la comprobación en línea)")
    else:
        checks["DOI real verificado"] = st[0] == "verificada"
        checks["DOI inventado detectado"] = st[1] == "doi_inexistente" and len(v["a_revisar"]) == 1
except Exception as e:  # noqa: BLE001
    print("(verificación en línea omitida:", e, ")")

bad = [k for k, ok in checks.items() if not ok]
for k, ok in checks.items():
    print("OK  " if ok else "FAIL", k)
assert not bad, f"fallaron: {bad}"
print("NUEVAS FUNCIONES OK")
