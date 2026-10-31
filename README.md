# apa7-tools

Servidor MCP en Python con herramientas **APA 7.ª edición** para documentos Word. Se conecta a cualquier agente (Claude, opencode, Cursor, Cline, Codex…): el usuario le pasa un `.docx` y recibe uno corregido y validado.

## Instalación (un comando)

```bash
pipx install git+https://github.com/TU_USUARIO/apa7-tools      # o: pip install apa7-tools  (cuando esté publicado en PyPI)
apa7-tools install opencode        # claude | claude-desktop | opencode | cursor   (--print muestra la configuración sin escribirla)
apa7-tools doctor                  # comprueba Python, dependencias, permisos, Word y red
```

Reinicie el agente y pídale, por ejemplo: *«Deja mi tesis.docx en APA 7 y dime qué falta»*. Sin instalar: `uvx --from git+https://github.com/TU_USUARIO/apa7-tools apa7-tools`.

Desde el código fuente: `python -m pip install -e .[dev]` y `python -m apa7_mcp install opencode`.

## Herramienta principal

`apa7_apply_apa(file_path)` analiza, corrige y valida. Devuelve los cambios, el puntaje antes/después, `passed`,
`needs_agent_action` (lo que requiere al usuario: citas sin referencia, citas textuales sin página, referencias incompletas,
tablas/figuras sin rótulo) y `suggestions` (recomendaciones opcionales: números en letras, siglas sin definir).

### Qué corrige automáticamente
- **Página:** márgenes 2.54 cm; Carta (se conserva A4 si ya lo era); sin guionado automático; texto en negro.
- **Fuente:** conserva Times New Roman 12, Calibri 11, Arial 11, Georgia 11, Aptos 11, Lucida 10 o Computer Modern 10 si el documento ya cumple; si no, Times New Roman 12.
- **Párrafos:** doble espacio, sin espacio extra, alineación izquierda, sangría de 1.27 cm; citas en bloque y listas se respetan.
- **Títulos:** niveles 1-5 (centrado/negrita, izquierda/negrita, negrita-cursiva, con sangría); detecta títulos sin estilo.
- **Resumen y palabras clave:** resumen sin sangría, «Palabras clave:» en cursiva. **Apéndices:** centrados, en página nueva.
- **Referencias:** título «Referencias» (no «Bibliografía») en página nueva, sangría francesa, orden alfabético letra por letra,
  DOI como `https://doi.org/…`, sin punto tras URL, sin «Recuperado de» ni ciudad de editorial, iniciales separadas, rangos con raya, `(2.ª ed.)`, «y» en vez de «&».
- **Citas:** «y» en vez de «&», p./pp., et al., citas múltiples en orden alfabético; citas textuales de 40+ palabras → **cita en bloque** (sin comillas, sangría, fuente tras el punto).
- **Cursivas en referencias:** título (libros, informes, tesis, páginas web), título del libro en capítulos, revista + volumen en artículos; respeta las que ya tienen cursivas.
- **Citas secundarias** («citado en»): se verifica la fuente secundaria. **Mismo autor y año:** avisa para distinguir con a, b, c.
- **Tablas** solo con líneas horizontales; **número de página** arriba a la derecha; **texto alternativo** en imágenes con rótulo.

### Qué solo detecta (lo resuelve el agente o la persona)
Citas sin referencia (y referencias no citadas), citas textuales sin página, citas ≥ 40 palabras sin formato de bloque, títulos de
referencias en mayúsculas (español), nombres completos en vez de iniciales, saltos de nivel de título, tablas/figuras sin rótulo.

### Estadísticos, puntuación y revisión de texto
- **Símbolos estadísticos en cursiva** (`M`, `SD`, `p`, `t`, `F`, `r`, `n`, `N`…): detecta y corrige, aunque estén en bloques de texto separados; avisa del cero inicial en `p`/`r`.
- **Puntuación mecánica:** corrige espacios dobles, espacio antes de coma/punto y falta de espacio tras la coma; avisa de palabras repetidas y signos sin emparejar.
- **Ortografía y gramática:** `apa7_extract_text` entrega los párrafos con índice para que el agente los revise y corrija con `apa7_replace_text`.

### Verificación de fuentes (`apa7_verify_references`)
Comprueba que las referencias **existen**: DOI en doi.org y Crossref (compara título, año y autor), artículos en Crossref, libros en
Open Library y enlaces. `doi_inexistente` y `no_coincide` son señales fuertes de referencia inventada; `no_encontrada` no prueba que sea
falsa (libros e informes suelen no estar en esas bases). **Envía las referencias a servicios externos**, por eso exige `confirm_external=true`.

### Fuera de alcance
Ética de investigación, derechos de autor, notas al pie, encabezados de nivel 4-5 «en línea» con el texto. La portada, el resumen y el encabezado
abreviado son opcionales y no se exigen.

## Otras herramientas

`apa7_extract_text`, `apa7_verify_references`, `apa7_validate_document`, `apa7_analyze_document`, `apa7_analyze_text`, `apa7_read_structure`, `apa7_build_document`,
`apa7_insert_blocks` (tablas, figuras, gráficos, listas, TOC, portada, referencias), `apa7_replace_text`,
`apa7_replace_references`, `apa7_format_reference`, `apa7_validate_reference`, `apa7_add_page_numbers`,
`apa7_format_tables`, `apa7_list_rules`.

## Español e inglés

El idioma se **detecta solo** del texto (o se fija con `language: "es" | "en"` en las herramientas, o en el perfil). Cambia lo que APA 7 escribe distinto:

| | Español | Inglés |
|---|---|---|
| Autores en paréntesis | `(Pérez y Soto, 2019)` | `(Smith & Jones, 2019)` |
| Cita narrativa | `Pérez y Soto (2019)` | `Smith and Jones (2019)` |
| Lista de referencias | `Pérez, J. y Soto, M.` | `Smith, J., & Jones, M.` |
| Sin fecha / en prensa | `s.f.` / `en prensa` | `n.d.` / `in press` |
| Edición | `(2.ª ed.)` | `(2nd ed.)` |
| Capítulo | `En M. Ruiz (Ed.), …` | `In M. Ruiz (Ed.), …` |
| Rótulos | Tabla, Figura, Nota., Apéndice | Table, Figure, Note., Appendix |
| Título de la lista | Referencias | References |
| Índices | Tabla de contenido, Índice de tablas | Table of Contents, List of Tables |

Se aplican en el análisis, las correcciones automáticas, `apa7_format_reference` / `apa7_validate_reference` y los bloques (`apa7_build_document`, `apa7_insert_blocks`).
Los mensajes de la herramienta siguen en español.

## Perfiles (cada institución es distinta)

`profile` (en `apa7_apply_apa`, `apa7_validate_document`, `apa7_analyze_document`, `apa7_insert_blocks`, `apa7_build_document`, `apa7_replace_references`):
un nombre (`apa7` = Carta, `a4`), la ruta a un `.json`, o un objeto. Si no se indica, se usa `apa7.profile.json` junto al documento (o la variable `APA7_PROFILE`).

```json
{ "extends": "a4", "font": "Arial", "size": 11, "line_spacing": 1.5, "margin_cm": 2.54, "first_line_cm": 1.27, "numbered_headings": true }
```
Opciones: `font`, `size`, `page_size` (`auto`|`letter`|`a4`), `margin_cm`, `line_spacing`, `first_line_cm`, `numbered_headings`, `language`.
Las opciones solo valen durante la llamada: no se filtran a la siguiente.

## Robustez

- **Errores estructurados:** toda herramienta devuelve `ok: true/false`; si falla, `error.code` + `error.hint` explican qué hacer
  (`FILE_NOT_FOUND`, `UNSUPPORTED_FORMAT` para `.doc/.pdf/…`, `ENCRYPTED_OR_LEGACY`, `CORRUPT_DOCX`, `TRACKED_CHANGES`, `INTEGRITY_ABORT`, `TIMEOUT`…). El servidor nunca se cae.
- **El original nunca se toca.** La salida se escribe en un temporal, se verifica (ZIP, XML, reapertura) y se reemplaza de forma atómica;
  si el destino existe se crea `nombre-2.docx` (o `overwrite=true`).
- **Guardia de integridad:** si una operación fuera a perder tablas, imágenes, enlaces, ecuaciones, secciones o texto, se cancela sin escribir.
- **Estructuras reales:** controles de contenido (`sdt`), tablas anidadas, varias secciones y «primera página distinta» (número de página en todas),
  tabla de contenido existente (no se reformatea ni se duplica) y control de cambios (`accept_tracked_changes=true` los acepta).
- **`dry_run=true`** simula `apa7_apply_apa` sin escribir. **`apa7_doctor`** / `python -m apa7_mcp doctor` diagnostican el entorno.
- Idempotente: aplicar `apa7_apply_apa` dos veces da el mismo documento.

## Rendimiento

Documento de 600 párrafos con 10 tablas y 60 referencias: `apply_apa` ≈ 1 s (antes 18.5 s); `analyze_document` ≈ 0.08 s (antes 3.1 s).
La lectura usa un índice de estilos en caché y XML directo; `matplotlib` solo se carga al dibujar un gráfico.

## Pruebas

```bash
pip install -e .[dev] && pytest   # 104 pruebas (APA7_OFFLINE=1 omite la red): corpus difícil, idempotencia, fuzz, errores, escenarios completos
python tests/smoke.py        # flujo completo
python tests/precision.py    # 19 verificaciones de reglas APA 7
python tests/new_features.py # estadísticos, puntuación, extracción de texto y verificación (usa la red)
python tests/mcp_client.py   # el servidor por stdio, como un agente real
```

## Límites

- Los números de página dentro de la tabla de contenido los calcula Word (pedirá actualizar campos).
- Citas y referencias se analizan con heurísticas para español e inglés; otros idiomas no están soportados.
- Solo `.docx`.
