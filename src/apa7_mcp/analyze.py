"""Análisis de conformidad APA 7 de un .docx o de texto plano (lectura rápida sobre el XML)."""
from __future__ import annotations

import re
from collections import Counter

from docx import Document

from . import lang as L
from . import options as O
from . import rules, textrules
from .rules import issue
from .scan import PInfo, W, clear_cache, has_toc, eff_align, eff_first, eff_line, eff_size, scan, styles_of
from .wordfmt import (ABSTRACT, APPENDIX, CAPTION, REF_HEADING, _ref_key, has_page_numbers, lang_is_spanish, reference_italic_missing, title_page_extra_blanks, KEYWORDS, stat_unitalic,
                      reference_range)

EMU_CM = 360000
WP = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"

_SMALL_NUM: dict[str, re.Pattern] = {}


def _small_num(code: str) -> re.Pattern:
    if code not in _SMALL_NUM:
        _SMALL_NUM[code] = re.compile(rf"(?<![\w.,/–:=<>-])([0-9])\s+({L.get(code).count_nouns})\b", re.I)
    return _SMALL_NUM[code]
ACRONYM = re.compile(r"\b([A-ZÁÉÍÓÚÑ]{2,6})\b")
ACR_IGNORE = {"APA", "DOI", "URL", "PDF", "HTML", "HTTP", "HTTPS", "ISBN", "ISSN", "EEUU", "UU", "TNR", "ISO", "SD", "ANOVA"}


# ------------------------------------------------------------------ puntaje


def score(issues: list[dict]) -> int:
    s = 100.0
    info = 0.0
    for i in issues:
        if i["severity"] == "error":
            s -= 5
        elif i["severity"] == "warning":
            s -= 2
        else:
            info += 0.5
    return max(0, round(s - min(info, 5)))


def _summary(sc: int, n: int, c: int, r: int) -> str:
    return f"Conformidad APA 7: {sc}/100. {n} incidencia(s); {c} cita(s) y {r} referencia(s) detectadas."


# ------------------------------------------------------------------ texto


def _cross_check(citations: list[dict], references: list[str]) -> tuple[list[dict], list[dict]]:
    cit, ref = [], []
    if not references:
        return cit, ref
    refs_n = [rules.norm(r) for r in references]
    cited: set[int] = set()
    seen: set[str] = set()
    for c in citations:
        if not c["authors"] or c["raw"] in seen:
            continue
        # en citas narrativas la regex puede arrastrar una palabra inicial ("Según Soto"): se prueba cada palabra
        cands = [rules.norm(w) for w in c["authors"][0].split()] if c.get("narrative") else [rules.norm(c["authors"][0])]
        y = rules.norm(c["year"].split("/")[0])
        idx = next((i for i, r in enumerate(refs_n) if any(r.startswith(a) for a in cands) and (not y or y in r)), None)
        if idx is not None and re.fullmatch(r"\d{4}", c["year"]):
            clones = [r for r in refs_n if any(r.startswith(a) for a in cands) and re.search(rf"\({c['year']}[a-z]\)", r)]
            if len(clones) >= 2 and c["raw"] not in seen:
                seen.add(c["raw"])
                cit.append(issue("cit-012", "citacion", "warning",
                                 f"La cita ({c['raw']}) es ambigua: hay varias obras de ese autor en {c['year']}; use {c['year']}a, {c['year']}b…",
                                 citation=c["raw"]))
        if idx is None:
            seen.add(c["raw"])
            cit.append(issue("cit-007", "citacion", "error",
                             f"La cita ({c['raw']}) no tiene entrada en la lista de referencias.",
                             citation=c["raw"], suggestion="Agregue la referencia o corrija el apellido/año de la cita."))
        else:
            cited.add(idx)
    for grp in rules.ambiguous_year_groups(references):
        ref.append(issue("ref-018", "referencias", "warning",
                         f"Hay {len(grp)} obras del mismo autor y año: distíngalas con letras (2020a, 2020b) en la lista y en las citas.",
                         reference=grp[0], suggestion="Ordene por título y agregue a, b, c… al año."))
    for i, r in enumerate(references):
        if i not in cited:
            ref.append(issue("ref-007", "referencias", "warning", "Referencia no citada en el texto.", reference=r,
                             suggestion="Toda referencia debe citarse en el texto; elimínela o añada la cita."))
    return cit, ref


def _text_style_checks(body: str, code: str = "es") -> list[dict]:
    out: list[dict] = []
    for m in list(_small_num(code).finditer(body))[:8]:
        out.append(issue("num-001", "texto", "info",
                         "Los números del cero al nueve se escriben con letras cuando cuentan elementos (tres participantes / three participants).",
                         original=m.group(0)))
    counts = Counter(a for a in ACRONYM.findall(body) if a not in ACR_IGNORE and not re.fullmatch(r"[IVXLCDM]+", a))
    for a, n in counts.most_common():
        if n >= 3 and not re.search(rf"\(\s*{a}\s*\)|\b{a}\s*\(", body) and len(out) < 14:
            out.append(issue("abr-001", "texto", "info",
                             f"La sigla {a} aparece {n} veces sin definirse: escriba el nombre completo y luego ({a}) la primera vez.",
                             original=a))
    return out


def analyze_text(text: str, file_path: str = "(texto)", lang: str | None = None, proof_text: str | None = None) -> dict:
    lines = [l.strip() for l in text.splitlines()]
    ref_idx = next((i for i in range(len(lines) - 1, -1, -1) if REF_HEADING.match(lines[i])), None)
    body = "\n".join(lines if ref_idx is None else lines[:ref_idx])
    if ref_idx is not None:  # las referencias terminan donde empiezan los apéndices
        tail = lines[ref_idx + 1:]
        stop = next((k for k, l in enumerate(tail) if APPENDIX.match(l)), len(tail))
        references = [l for l in tail[:stop] if l]
    else:
        references = [l for l in lines if rules.looks_like_reference(l)]
    code = L.resolve(lines) if lang is None else lang
    lg = L.get(code)
    paren, cit_issues = rules.detect_citations(body, code)
    citations = paren + rules.detect_narrative(body)
    cit_issues += rules.detect_quotes(body)
    cit_issues += rules.detect_narrative_amp(body, code)
    ref_issues = [i for r in references for i in rules.validate_reference(r, code)]
    c2, r2 = _cross_check(citations, references)
    cit_issues += c2
    ref_issues += r2
    if ref_idx is None and citations:
        cit_issues.append(issue("est-001", "estructura", "error",
                                f'Hay citas pero no existe una sección titulada "{lg.ref_heading}".', citation=""))
    style = _text_style_checks(body, code) + textrules.detect_stat_zero(body)
    seen_rule: Counter = Counter()
    for it in textrules.proofread(proof_text if proof_text is not None else body):  # revisión mecánica; máx. 6 avisos por tipo para no saturar
        seen_rule[it["ruleId"]] += 1
        if seen_rule[it["ruleId"]] <= 6:
            style.append(it)
    allv = cit_issues + ref_issues + style
    sc = score(allv)
    return {"filePath": file_path, "wordCount": len(body.split()), "citations": citations, "references": references,
            "citationIssues": cit_issues, "referenceIssues": ref_issues, "formatIssues": style,
            "totalIssues": len(allv), "score": sc, "summary": _summary(sc, len(allv), len(citations), len(references))}


# ------------------------------------------------------------------ documento


def measure(doc, infos: list[PInfo]) -> dict:
    st = styles_of(doc)
    sec = doc.sections[0]
    body = [pi for pi in infos if len(pi.text) > 40 and not pi.level and not pi.listy and not CAPTION.match(pi.text)]
    top = lambda c: c.most_common(1)[0][0] if c else None  # noqa: E731
    return {
        "page_w_cm": round(sec.page_width / EMU_CM, 2),
        "page_h_cm": round(sec.page_height / EMU_CM, 2),
        "margins_cm": {k: round((getattr(sec, f"{k}_margin") or 0) / EMU_CM, 2) for k in ("top", "bottom", "left", "right")},
        "font": top(Counter((pi.font or st.eff(pi.sid)["font"]) for pi in body)),
        "size": top(Counter(eff_size(pi, st) for pi in body)),
        "spacing": top(Counter(round(eff_line(pi, st), 2) for pi in body)),
        "alignment": top(Counter(eff_align(pi, st) for pi in body)),
        "body_paragraphs": len(body),
        "first_line_indented": sum(1 for pi in body if eff_first(pi, st) > 0),
    }


def _italic_not_bold(p_el) -> bool:
    runs = [r for r in p_el.iter(W + "r") if "".join(t.text or "" for t in r.iter(W + "t")).strip()]
    def on(r, tag):
        e = r.find(W + "rPr/" + W + tag)
        return e is not None and e.get(W + "val") not in ("0", "false")
    return bool(runs) and all(on(r, "i") for r in runs) and not any(on(r, "b") for r in runs)


def _spacing_extra(p_el) -> bool:
    sp = p_el.find(W + "pPr/" + W + "spacing")
    return sp is not None and any(int(sp.get(W + k) or 0) > 0 for k in ("before", "after"))


def _structure(doc, infos: list[PInfo]) -> list[dict]:
    st = styles_of(doc)
    lg = L.current()
    out: list[dict] = []
    texts = [pi.text for pi in infos]
    n_tables = sum(1 for _ in doc.element.body.iter(W + "tbl"))
    n_imgs = sum(1 for pi in infos for _ in pi.el.iter(WP + "docPr"))
    cap_t = sum(1 for t in texts if re.match(r"^(?:Tabla|Table)\s+\d+", t, re.I))
    cap_f = sum(1 for t in texts if re.match(r"^(?:Figura|Figure)\s+\d+", t, re.I))
    if n_tables > cap_t:
        out.append(issue("tbl-001", "tablas_figuras", "warning",
                         f"{n_tables - cap_t} tabla(s) sin rótulo: '{lg.table} N' en negrita y, debajo, el título en cursiva (encima de la tabla).",
                         suggestion="Use apa7_insert_blocks o agregue los rótulos."))
    if n_imgs > cap_f:
        out.append(issue("fig-001", "tablas_figuras", "warning",
                         f"{n_imgs - cap_f} figura(s) sin rótulo: '{lg.figure} N' en negrita y el título en cursiva (encima de la figura)."))
    missing_alt = sum(1 for pi in infos for d in pi.el.iter(WP + "docPr") if not d.get("descr"))
    if missing_alt:
        out.append(issue("fig-002", "tablas_figuras", "warning",
                         f"{missing_alt} imagen(es) sin texto alternativo (accesibilidad).", autofix=True))
    if not has_page_numbers(doc):
        out.append(issue("est-002", "estructura", "warning", "Falta el número de página (arriba a la derecha).", autofix=True))

    rr = reference_range(infos)
    if rr:
        head = infos[rr[0]]
        if head.text.lower() != lg.ref_heading.lower():
            out.append(issue("est-005", "estructura", "warning", f'El título debe ser "{lg.ref_heading}" (no "{head.text}").', autofix=True))
        prev = infos[rr[0] - 1].el if rr[0] > 0 else None
        has_break = head.page_break_before or (prev is not None and any(
            b.get(W + "type") == "page" for b in prev.iter(W + "br")))
        if rr[0] > 0 and not has_break:
            out.append(issue("est-004", "estructura", "warning", "La lista de referencias empieza en una página nueva.", autofix=True))
        refs = [pi for pi in infos[rr[0] + 1: rr[1]] if pi.text]
        no_ital = reference_italic_missing(infos, rr)
        if no_ital:
            out.append(issue("ref-017", "referencias", "warning",
                             f"{len(no_ital)} referencia(s) sin cursiva: APA 7 pone en cursiva el título (libros, informes, tesis) "
                             "o la revista y el volumen (artículos).", autofix=True))
        bad = [pi for pi in refs if eff_first(pi, st) >= 0]
        if bad:
            out.append(issue("ref-008", "referencias", "warning",
                             f"{len(bad)} referencia(s) sin sangría francesa de 1.27 cm.", autofix=True))
        keys = [_ref_key(pi.text) for pi in refs]
        if keys != sorted(keys):
            out.append(issue("ref-009", "referencias", "warning", "La lista de referencias no está en orden alfabético.", autofix=True))

    levels = [pi.level for pi in infos if 1 <= pi.level <= 5]
    for prev, cur in zip(levels, levels[1:]):
        if cur > prev + 1:
            out.append(issue("tit-002", "titulos", "warning",
                             f"Se salta un nivel de título (nivel {prev} → nivel {cur}); mantenga la jerarquía."))
            break
    if not O.active().numbered_headings and any(re.match(r"^\d+(\.\d+)*[.)]\s", pi.text) for pi in infos if 1 <= pi.level <= 5):
        out.append(issue("tit-003", "titulos", "info", "Los títulos no se numeran en APA 7 salvo que la institución lo pida."))
    if not levels and sum(len(t.split()) for t in texts) > 400:
        out.append(issue("tit-001", "titulos", "info",
                         "El documento no usa estilos de título (Heading 1-5); no se puede generar tabla de contenido."))

    ab = next((k for k, pi in enumerate(infos) if ABSTRACT.match(pi.text)), None)
    if ab is not None:
        nxt = next((pi for pi in infos[ab + 1:] if pi.text), None)
        if nxt is not None and eff_first(nxt, st) > 0:
            out.append(issue("abs-001", "estructura", "warning", "La primera línea del resumen va sin sangría.", autofix=True))

    for pi in infos:
        if 1 <= pi.level <= 5 and re.fullmatch(r"(?:\d+(?:\.\d+)*[.)]?\s*)?(?:introducci[óo]n|introduction)\.?", pi.text, re.I):
            out.append(issue("est-006", "estructura", "warning",
                             f'APA 7 no usa el encabezado «{pi.text}»: el texto inicia en página nueva con el título del trabajo centrado en negrita.',
                             original=pi.text))
            break
    if has_toc(doc):
        out.append(issue("est-007", "estructura", "info",
                         "Tiene tabla de contenido: APA 7 no la exige en trabajos ni artículos; manténgala solo si la institución la pide."))
    paren_alone = [m.group(0) for pi in infos for m in re.finditer(r"\[[A-ZÁÉÍÓÚÑ]{2,8}\]", pi.text)
                   if pi.text.rfind("(", 0, m.start()) <= pi.text.rfind(")", 0, m.start())]
    if paren_alone:
        out.append(issue("abr-002", "texto", "warning",
                         f"La sigla {paren_alone[0]} va entre corchetes en texto corrido: use paréntesis, Nombre completo (SIGLA); "
                         "los corchetes solo se usan si la sigla se define dentro de un paréntesis.", original=paren_alone[0]))
    own = next((pi.text for pi in infos if re.search(r"\b(?:elaboraci[óo]n propia|fuente propia|own (?:work|elaboration))\b", pi.text, re.I)), None)
    if own:
        out.append(issue("fig-003", "tablas_figuras", "info",
                         "«Elaboración propia» sobra: en APA 7 una tabla o figura sin cita externa se entiende creada por el autor.",
                         original=own[:80]))

    for pi in infos:
        if pi.level in (4, 5) and pi.text:
            nxt = next((x for x in infos[pi.i + 1:pi.i + 2]), None)
            if nxt is not None and nxt.text and not nxt.level and pi.el.getnext() is nxt.el:
                out.append(issue("tit-004", "titulos", "warning",
                                 f"Los títulos de nivel 4 y 5 van en línea: «{pi.text[:40]}.» en negrita y el texto sigue en el mismo renglón.",
                                 original=pi.text[:60], autofix=True))
                break
    bad_h = next((pi.text for pi in infos if 1 <= pi.level <= 5 and pi.text and _italic_not_bold(pi.el)), None)
    if bad_h:
        out.append(issue("tit-005", "titulos", "warning",
                         f"El título «{bad_h[:40]}» está en cursiva sin negrita: en APA 7 todos los niveles llevan negrita.",
                         original=bad_h[:60], autofix=True))
    kw = next((pi for pi in infos if KEYWORDS.match(pi.text)), None)
    if kw is not None and eff_first(kw, st) <= 0:
        out.append(issue("abs-002", "estructura", "warning",
                         "La línea de palabras clave va con sangría de primera línea de 1.27 cm.", autofix=True))
    rr2 = reference_range(infos)
    spaced = [pi for pi in (infos[rr2[0] + 1:rr2[1]] if rr2 else []) if pi.text and _spacing_extra(pi.el)]
    if spaced:
        out.append(issue("ref-019", "referencias", "warning",
                         f"{len(spaced)} referencia(s) con espacio extra antes o después del párrafo: APA 7 usa interlineado doble continuo.",
                         autofix=True))
    n_blank = len(title_page_extra_blanks(infos))
    if n_blank:
        out.append(issue("est-008", "estructura", "warning",
                         f"Portada con {n_blank} línea(s) en blanco de más: deje una sola línea en blanco entre el título y el autor.", autofix=True))

    rr_stats = reference_range(infos)
    unital = sum(stat_unitalic(pi.el) for pi in infos if pi.text and not (rr_stats and rr_stats[0] <= pi.i < rr_stats[1]))
    if unital:
        out.append(issue("stat-001", "texto", "warning",
                         f"{unital} símbolo(s) estadístico(s) sin cursiva (p, M, SD, t, F, r, n…): APA 7 los escribe en cursiva.", autofix=True))

    hy = doc.settings.element.find(W + "autoHyphenation")
    if hy is not None and hy.get(W + "val") not in ("0", "false", "off"):
        out.append(issue("fmt-009", "formato", "warning", "Desactive la separación automática de palabras con guiones.", autofix=True))
    colored = sum(1 for c in doc.element.body.iter(W + "color") if c.get(W + "val") not in (None, "auto", "000000"))
    if colored:
        out.append(issue("fmt-010", "formato", "info", f"{colored} fragmento(s) con color de texto: APA 7 usa texto negro.", autofix=True))
    return out


def analyze_document(path: str) -> dict:
    doc = Document(path)
    clear_cache()
    infos = scan(doc)
    texts = [pi.text for pi in infos]
    L.resolve(texts)  # idioma del documento (perfil o detección automática)
    rr0 = reference_range(infos)
    proof = "\n".join(pi.raw for pi in infos if pi.raw.strip() and not (rr0 and rr0[0] <= pi.i < rr0[1]))  # texto sin normalizar
    rep = analyze_text("\n".join(texts), path, proof_text=proof)
    fmt = rules.validate_formatting(measure(doc, infos))
    rep["formatIssues"] = fmt + _structure(doc, infos) + rep["formatIssues"]
    allv = rep["citationIssues"] + rep["referenceIssues"] + rep["formatIssues"]
    rep["paragraphCount"] = len(infos)
    rep["tables"] = sum(1 for _ in doc.element.body.iter(W + "tbl"))
    rep["images"] = sum(1 for pi in infos for _ in pi.el.iter(WP + "docPr"))
    rep["totalIssues"] = len(allv)
    rep["score"] = score(allv)
    rep["summary"] = _summary(rep["score"], len(allv), len(rep["citations"]), len(rep["references"]))
    return rep
