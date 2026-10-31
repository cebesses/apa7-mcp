import pytest
from docx import Document

from apa7_mcp import pipeline


@pytest.mark.parametrize("authors,expected", [
    ("Ana Pruebas", "Ana Pruebas"),
    (["Ana Pruebas"], "Ana Pruebas"),
    (["Ana Pruebas", "Luis Soto"], "Ana Pruebas y Luis Soto"),
    (["Ana Pruebas", "Luis Soto", "Eva Gil"], "Ana Pruebas, Luis Soto y Eva Gil"),
])
def test_title_page_authors_text_or_list(tmp_path, authors, expected):
    out = tmp_path / "t.docx"
    pipeline.build_document(str(out), [{"type": "title_page", "title": "Un título", "authors": authors,
                                        "affiliation": "Universidad X"}], language="es")
    texts = [p.text for p in Document(str(out)).paragraphs]
    assert expected in texts
