"""Corpus de documentos difíciles generados por código (sin archivos binarios en el repositorio)."""
from __future__ import annotations

import random
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Cm, Pt

W_NS = nsdecls("w")


def _base_messy(d: Document) -> None:
    d.add_paragraph().add_run("Introducción").bold = True
    p = d.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.add_run("El aprendizaje autónomo mejora el rendimiento (Pérez & Soto, 2019, p. 23-25). Además se ha visto que reduce la deserción (Gómez, 2020).")
    d.add_paragraph().add_run("Referencias").bold = True
    d.add_paragraph("Soto, M. (2020). Deserción universitaria. Bogotá: Editorial Uno. Recuperado de https://doi.org/10.1/x")
    d.add_paragraph("Pérez, J. & Soto, M. (2019). Aprendizaje autónomo. Revista de Educación, 12(3), 20-30.")


def plain_bad(path: Path) -> Path:
    d = Document()
    s = d.sections[0]
    s.page_width, s.page_height = Cm(21), Cm(29.7)
    s.left_margin = s.right_margin = Cm(1.5)
    d.styles["Normal"].font.name = "Calibri"
    d.styles["Normal"].font.size = Pt(11)
    _base_messy(d)
    t = d.add_table(rows=3, cols=3)
    t.style = "Table Grid"
    d.save(path)
    return path


def empty_doc(path: Path) -> Path:
    Document().save(path)
    return path


def tracked_changes(path: Path) -> Path:
    d = Document()
    d.add_paragraph("Texto antes de los cambios que se mantiene intacto en el documento.")
    p = d.add_paragraph()
    p._p.append(parse_xml(f'<w:ins {W_NS} w:id="1" w:author="A" w:date="2026-01-01T00:00:00Z"><w:r><w:t xml:space="preserve">Texto insertado. </w:t></w:r></w:ins>'))
    p._p.append(parse_xml(f'<w:del {W_NS} w:id="2" w:author="A" w:date="2026-01-01T00:00:00Z"><w:r><w:delText xml:space="preserve">Texto borrado.</w:delText></w:r></w:del>'))
    d.save(path)
    return path


def sdt_doc(path: Path) -> Path:
    """Párrafos dentro de un control de contenido (w:sdt), incluido un título."""
    d = Document()
    d.add_paragraph("Párrafo normal previo al control de contenido, suficientemente largo para ser cuerpo.")
    body = d.element.body
    sect = body.find(qn("w:sectPr"))
    sdt = parse_xml(
        f'<w:sdt {W_NS}><w:sdtPr><w:alias w:val="Contenido"/></w:sdtPr><w:sdtContent>'
        '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Método</w:t></w:r></w:p>'
        '<w:p><w:r><w:t>Dentro del control de contenido hay un párrafo largo que también debe recibir formato APA completo.</w:t></w:r></w:p>'
        '</w:sdtContent></w:sdt>')
    sect.addprevious(sdt)
    d.save(path)
    return path


def two_sections_titlepg(path: Path) -> Path:
    d = Document()
    d.add_paragraph("Primera sección con texto suficiente para contar como cuerpo del documento en la prueba.")
    d.sections[0].different_first_page_header_footer = True
    d.add_section()
    d.add_paragraph("Segunda sección con más texto de ejemplo para el cuerpo del documento de la prueba.")
    d.save(path)
    return path


def nested_tables(path: Path) -> Path:
    d = Document()
    d.add_paragraph("Texto previo a la tabla anidada, con longitud suficiente para ser un párrafo de cuerpo.")
    t = d.add_table(rows=2, cols=2)
    cell = t.cell(0, 0)
    inner = cell.add_table(rows=2, cols=2)
    inner.cell(0, 0).text = "interna"
    t.cell(1, 1).text = "externa"
    d.add_paragraph("Texto posterior.")
    d.save(path)
    return path


def theme_fonts(path: Path) -> Path:
    """Fuentes de tema (sin nombre explícito) en los estilos por defecto."""
    d = Document()
    rpr = d.styles.element.find(qn("w:docDefaults")).find(qn("w:rPrDefault")).find(qn("w:rPr"))
    rf = rpr.find(qn("w:rFonts"))
    for k in list(rf.attrib):
        del rf.attrib[k]
    rf.set(qn("w:asciiTheme"), "minorHAnsi")
    rf.set(qn("w:hAnsiTheme"), "minorHAnsi")
    d.add_paragraph("Documento con fuentes de tema: el análisis no debe fallar y la corrección debe fijar la fuente.")
    d.save(path)
    return path


def big_doc(path: Path, n: int = 3000) -> Path:
    d = Document()
    for i in range(n):
        if i % 50 == 0:
            d.add_paragraph().add_run(f"Sección {i // 50 + 1}").bold = True
        d.add_paragraph("El aprendizaje autónomo mejora el rendimiento académico según diversos estudios (Pérez, 2019, p. 23). " * 2)
    d.save(path)
    return path


def with_toc(path: Path) -> Path:
    from apa7_mcp import pipeline

    pipeline.build_document(str(path), [
        {"type": "toc"}, {"type": "page_break"},
        {"type": "heading", "level": 1, "text": "Introducción"},
        {"type": "paragraph", "text": "Texto del cuerpo suficientemente largo para la prueba de la tabla de contenido."},
        {"type": "heading", "level": 2, "text": "Método"},
        {"type": "paragraph", "text": "Más texto del cuerpo para que haya contenido bajo el segundo encabezado."},
    ])
    return path


WORDS = ["aprendizaje", "estudiantes", "resultados", "método", "análisis", "educación", "muestra", "variables", "🙂", "نص", "naïve"]


def fuzz_doc(path: Path, seed: int) -> Path:
    """Documento aleatorio: títulos en negrita, listas, tablas, citas, texto con emoji/RTL/espacios raros."""
    rnd = random.Random(seed)
    d = Document()
    if rnd.random() < 0.5:
        d.sections[0].page_width, d.sections[0].page_height = Cm(21), Cm(29.7)
    for _ in range(rnd.randint(3, 40)):
        kind = rnd.choice(["head", "para", "para", "list", "quote", "table", "empty", "stat"])
        text = " ".join(rnd.choice(WORDS) for _ in range(rnd.randint(2, 40)))
        if kind == "head":
            d.add_paragraph().add_run(text[:60]).bold = True
        elif kind == "para":
            p = d.add_paragraph()
            p.alignment = rnd.choice([None, WD_ALIGN_PARAGRAPH.JUSTIFY, WD_ALIGN_PARAGRAPH.CENTER])
            p.add_run(text + rnd.choice([" (Pérez & Soto, 2019, p. 3-4).", " (Gómez, 2020).", "  ,mal  puntuado .", "."]))
        elif kind == "list":
            d.add_paragraph(text[:50], style="List Bullet")
        elif kind == "quote":
            d.add_paragraph("“" + text + "” (Lynch, 2012).")
        elif kind == "table":
            t = d.add_table(rows=rnd.randint(1, 4), cols=rnd.randint(1, 4))
            for r in t.rows:
                for c in r.cells:
                    c.text = rnd.choice(WORDS)
        elif kind == "stat":
            d.add_paragraph("Los grupos difirieron (M = 3.2, SD = 0.8; t(23) = 2.1, p < .05).")
        else:
            d.add_paragraph("")
    if rnd.random() < 0.6:
        d.add_paragraph().add_run(rnd.choice(["Referencias", "Bibliografía"])).bold = True
        for _ in range(rnd.randint(1, 6)):
            d.add_paragraph(f"{rnd.choice(['Soto','Pérez','Álvarez'])}, J.M. (20{rnd.randint(10,24)}). {rnd.choice(WORDS).title()} Y Otros. Editorial. doi:10.1000/{rnd.randint(1,99)}.")
    d.save(path)
    return path
