# apa7-tools

Servidor MCP en Python que ayuda a agentes de IA a revisar y dar formato APA 7 a documentos de Microsoft Word (`.docx`). Incluye herramientas para validar contenido, aplicar formato, construir o editar documentos y revisar referencias.

Repositorio: [cebesses/apa7-mcp](https://github.com/cebesses/apa7-mcp) · Licencia MIT · Python 3.10 o posterior.

## Funciones

- **Aplicar APA 7:** revisa y formatea el documento; informa el resultado y los puntos que requieren revisión humana.
- **Validar sin modificar:** analiza un `.docx`, texto o una referencia individual.
- **Leer documentos:** extrae texto por párrafos o devuelve un esquema de títulos y rótulos.
- **Crear y editar:** crea documentos desde bloques, inserta tablas, figuras, gráficos, listas y otros bloques, o reemplaza texto y referencias.
- **Dar formato a referencias:** genera referencias APA 7 a partir de sus datos.
- **Verificar fuentes:** consulta DOI, Crossref, Open Library y enlaces. Esta función envía las referencias a servicios externos y exige consentimiento explícito.
- **Español e inglés:** detecta el idioma automáticamente; también puede indicarse en las llamadas o en un perfil.

El proyecto trabaja con `.docx`; no procesa directamente PDF ni los formatos binarios antiguos de Word (`.doc`). La corrección del contenido académico y los casos ambiguos siguen requiriendo revisión de una persona o del agente.

## Instalación

Instala la versión del repositorio directamente desde GitHub con `pipx`:

```bash
pipx install git+https://github.com/cebesses/apa7-mcp.git
```

También se puede usar `uvx` sin instalación permanente:

```bash
uvx --from git+https://github.com/cebesses/apa7-mcp.git apa7-tools --version
```

No se presupone que el paquete esté publicado en PyPI.

## Conectar un agente

Con `apa7-tools` instalado, registra el servidor MCP para uno de los agentes configurados por el instalador:

```bash
apa7-tools install claude
# Alternativas: claude-desktop, opencode, cursor
```

Para ver la configuración sin escribir archivos:

```bash
apa7-tools install cursor --print
```

Reinicia el agente después de configurarlo. El comando `apa7-tools doctor` comprueba dependencias y el entorno; `apa7-tools doctor --offline` evita la comprobación de red. Para ejecutar el servidor manualmente por stdio, usa `apa7-tools` o `python -m apa7_mcp`.

## Herramientas MCP

| Herramienta | Función |
|---|---|
| `apa7_apply_apa` | Aplica formato APA 7 y devuelve validación y asuntos pendientes. Admite `dry_run`. |
| `apa7_check` | Valida un documento, texto o referencia sin modificarlo. |
| `apa7_read` | Devuelve el esquema del documento o texto paginado con índices de párrafo. |
| `apa7_build_document` | Crea un `.docx` a partir de bloques. |
| `apa7_insert_blocks` | Inserta bloques en un `.docx` existente. |
| `apa7_replace_text` | Reemplaza fragmentos de texto exactos. |
| `apa7_replace_references` | Sustituye la lista de referencias. |
| `apa7_format_reference` | Construye una referencia APA 7 desde datos estructurados. |
| `apa7_verify_references` | Comprueba referencias con servicios externos; requiere `confirm_external=true`. |
| `apa7_help` | Consulta reglas, bloques disponibles o diagnóstico del entorno. |

### Ejemplo de uso

Una vez conectado el servidor, puedes pedirle al agente: «Revisa `tesis.docx` con APA 7, guarda una copia corregida y dime qué debo resolver manualmente». El servidor crea una salida; el original no se modifica. Confirma la ruta de salida y revisa los elementos marcados para intervención humana.

## Perfiles

Las operaciones que aceptan `profile` admiten un nombre (`apa7` o `a4`), una ruta a un archivo JSON o un objeto de opciones. Si no se especifica, se busca `apa7.profile.json` junto al documento y después la variable de entorno `APA7_PROFILE`.

Ejemplo de archivo de perfil:

```json
{
  "extends": "a4",
  "font": "Arial",
  "size": 11,
  "line_spacing": 1.5,
  "margin_cm": 2.54,
  "first_line_cm": 1.27,
  "numbered_headings": true,
  "language": "es"
}
```

Las opciones disponibles incluyen fuente y tamaño, tamaño de página (`auto`, `letter` o `a4`), márgenes, interlineado, sangría de primera línea, numeración de títulos e idioma (`auto`, `es` o `en`). Los perfiles `apa7` y `a4` seleccionan Carta y A4, respectivamente.

## Desarrollo y pruebas

```bash
python -m pip install -e ".[dev]"
python -m pytest -q
python -m apa7_mcp doctor --offline
```

Para ejecutar además las pruebas auxiliares del proyecto:

```bash
python tests/smoke.py
python tests/precision.py
python tests/mcp_client.py
```

Las pruebas que verifican servicios de referencias externos pueden requerir red; para la suite principal, `APA7_OFFLINE=1` desactiva dependencias de red.

## Limitaciones

- La verificación externa de referencias envía datos a terceros y solo se ejecuta con confirmación.
- La comprobación automática ayuda a detectar problemas, pero no sustituye la revisión de exactitud académica ni los requisitos particulares de cada institución.
- Los números de página de una tabla de contenido se actualizan en Word.
- Las referencias bibliográficas y las reglas lingüísticas se analizan mediante heurísticas; revisa los resultados antes de entregar un trabajo.

## Licencia

Este proyecto se distribuye bajo la licencia [MIT](LICENSE).
