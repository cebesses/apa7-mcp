"""Verificación de que las fuentes existen: DOI (doi.org), Crossref (artículos), Open Library (libros) y enlaces.

Envía el texto de las referencias a servicios externos; por eso la herramienta exige confirmación expresa.
"""
from __future__ import annotations

import concurrent.futures as cf
import difflib
import json
import re
import time
import urllib.error
import urllib.parse
import urllib.request

from . import rules

UA = "apa7-tools/2.0 (verificacion de referencias APA 7)"
DOI_RE = re.compile(r"10\.\d{4,9}/[^\s<>\"]+")
URL_RE = re.compile(r"https?://[^\s<>\"]+")
_cache: dict[str, tuple[int, bytes]] = {}


class Offline(Exception):
    pass


def _get(url: str, timeout: float, method: str = "GET") -> tuple[int, bytes]:
    key = f"{method} {url}"
    if key in _cache:
        return _cache[key]
    req = urllib.request.Request(url, method=method, headers={"User-Agent": UA, "Accept": "application/json"})
    out = (0, b"")
    for attempt in range(3):  # 429 / 5xx son temporales: se reintenta con espera creciente
        try:
            with urllib.request.urlopen(req, timeout=timeout) as r:
                out = (r.status, r.read() if method == "GET" else b"")
        except urllib.error.HTTPError as e:
            out = (e.code, b"")
        except (urllib.error.URLError, TimeoutError, OSError) as e:
            raise Offline(str(e)) from e
        if out[0] not in (429, 500, 502, 503, 504):
            break
        time.sleep(0.8 * (attempt + 1))
    if out[0] in (429, 500, 502, 503, 504):
        raise Offline(f"el servicio respondió HTTP {out[0]} (límite de velocidad o caída temporal); reintente en unos minutos")
    _cache[key] = out
    return out


def _clean(s: str) -> str:
    return re.sub(r"[^a-z0-9 ]", "", rules.norm(s.replace("*", ""))).strip()


def _sim(a: str, b: str) -> float:
    return difflib.SequenceMatcher(None, _clean(a), _clean(b)).ratio()


def _parts(ref: str) -> dict:
    year = re.search(r"\((\d{4})[a-z]?[,)]", ref)
    head = ref.split("(", 1)[0]
    surname = re.split(r"[,.]", head.strip(), maxsplit=1)[0].strip() if head.strip() else ""
    title = rules._title_of(ref.replace("*", ""))
    doi = DOI_RE.search(ref)
    url = URL_RE.search(ref)
    return {
        "year": int(year.group(1)) if year else None,
        "surname": surname,
        "title": title,
        "doi": doi.group(0).rstrip(".,;)") if doi else None,
        "url": url.group(0).rstrip(".,;)") if url and not doi else None,
    }


def _crossref_meta(msg: dict) -> dict:
    issued = (msg.get("issued") or {}).get("date-parts") or [[None]]
    return {
        "title": (msg.get("title") or [""])[0],
        "year": issued[0][0] if issued and issued[0] else None,
        "authors": [a.get("family", "") for a in msg.get("author", [])][:3],
        "journal": (msg.get("container-title") or [""])[0],
    }


def verify_one(ref: str, timeout: float = 10.0) -> dict:
    p = _parts(ref)
    res: dict = {"reference": ref, "status": "no_encontrada", "detail": "", "doi": p["doi"]}
    try:
        if p["doi"]:
            doi = urllib.parse.quote(p["doi"], safe="/")
            st, body = _get(f"https://api.crossref.org/works/{doi}", timeout)
            if st == 200:
                m = _crossref_meta(json.loads(body)["message"])
                title_ok = (not p["title"]) or _sim(p["title"], m["title"]) >= 0.6 or _clean(m["title"]) in _clean(ref)
                year_ok = p["year"] is None or m["year"] is None or abs(p["year"] - m["year"]) <= 1
                author_ok = not m["authors"] or not p["surname"] or any(_clean(a) and _clean(a) in _clean(ref.split("(", 1)[0]) for a in m["authors"])
                real = f"«{m['title']}» ({m['year']}), {', '.join(m['authors'])}; {m['journal']}"
                if title_ok and year_ok and author_ok:
                    res.update(status="verificada", detail=f"El DOI existe y coincide: {real}")
                else:
                    bad = [n for n, ok in (("título", title_ok), ("año", year_ok), ("autor", author_ok)) if not ok]
                    res.update(status="no_coincide", detail=f"El DOI existe pero no coincide en {', '.join(bad)}. Registro real: {real}")
                return res
            st2, body2 = _get(f"https://doi.org/api/handles/{doi}", timeout)
            if st2 == 200 and json.loads(body2).get("responseCode") == 1:
                res.update(status="verificada", detail="El DOI existe (no está en Crossref: no se pudieron comparar los metadatos).")
            else:
                res.update(status="doi_inexistente", detail="Ese DOI no existe: la referencia puede ser inventada o tener el DOI mal escrito.")
            return res

        if p["url"] and not p["title"]:
            st, _ = _get(p["url"], timeout, "HEAD")
            res.update(status="enlace_activo" if st < 400 else "enlace_roto", detail=f"HTTP {st}")
            return res

        if p["title"]:
            q = urllib.parse.urlencode({"query.bibliographic": f"{p['title']} {p['surname']} {p['year'] or ''}".strip(), "rows": 3,
                                        "select": "DOI,title,issued,author,container-title"})
            st, body = _get(f"https://api.crossref.org/works?{q}", timeout)
            if st == 200:
                for it in json.loads(body)["message"]["items"]:
                    m = _crossref_meta(it)
                    if _sim(p["title"], m["title"]) >= 0.85 and (p["year"] is None or m["year"] is None or abs(p["year"] - m["year"]) <= 1):
                        res.update(status="probable", doi=it.get("DOI"),
                                   detail=f"Crossref tiene un registro equivalente: «{m['title']}» ({m['year']}). Agregue el DOI https://doi.org/{it.get('DOI')}.")
                        return res
            q = urllib.parse.urlencode({"title": p["title"], "limit": 3, "fields": "title,author_name,first_publish_year"})
            st, body = _get(f"https://openlibrary.org/search.json?{q}", timeout)
            if st == 200:
                for d in json.loads(body).get("docs", []):
                    if _sim(p["title"], d.get("title", "")) >= 0.85:
                        res.update(status="probable", detail=f"Open Library registra «{d.get('title')}» ({d.get('first_publish_year')}), "
                                                              f"{', '.join(d.get('author_name', [])[:2])}.")
                        return res
            if p["url"]:
                st, _ = _get(p["url"], timeout, "HEAD")
                res.update(status="enlace_activo" if st < 400 else "enlace_roto", detail=f"HTTP {st}")
                return res
        res["detail"] = ("No se encontró en Crossref ni Open Library. Puede ser un libro, informe o recurso que no está en esas bases: "
                         "confirme manualmente que existe.")
    except Offline as e:
        res.update(status="sin_conexion", detail=f"No se pudo consultar: {e}")
    except (ValueError, KeyError) as e:
        res.update(status="sin_conexion", detail=f"Respuesta inesperada del servicio: {e}")
    return res


REVIEW = {"no_coincide", "doi_inexistente", "no_encontrada", "enlace_roto"}


def verify_references(entries: list[str], timeout: float = 10.0, workers: int = 3) -> dict:
    with cf.ThreadPoolExecutor(max_workers=workers) as ex:
        results = list(ex.map(lambda r: verify_one(r, timeout), entries))
    counts: dict[str, int] = {}
    for r in results:
        counts[r["status"]] = counts.get(r["status"], 0) + 1
        r["revisar"] = r["status"] in REVIEW
    return {
        "total": len(results),
        "resumen": counts,
        "a_revisar": [r for r in results if r["revisar"]],
        "results": results,
        "nota": ("«verificada» = el DOI existe y coincide con título/año/autor. «no_encontrada» no prueba que sea falsa (libros e informes "
                 "suelen no estar en Crossref); «doi_inexistente» y «no_coincide» sí son señales fuertes de referencia inventada o errónea."),
    }
