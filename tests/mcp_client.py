"""Prueba del servidor por stdio como lo haría un agente real."""
import asyncio
import sys
import tempfile
from pathlib import Path

from docx import Document
from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client


async def main() -> None:
    tmp = Path(tempfile.mkdtemp())
    src = tmp / "t.docx"
    d = Document()
    d.add_paragraph("Introducción").runs[0].bold = True
    d.add_paragraph("Texto de prueba suficientemente largo para ser un párrafo de cuerpo normal (Pérez, 2020).")
    d.add_paragraph("Referencias").runs[0].bold = True
    d.add_paragraph("Pérez, J. (2020). Un libro. Editorial.")
    d.save(src)
    params = StdioServerParameters(command=sys.executable, args=["-m", "apa7_mcp"])
    async with stdio_client(params) as (r, w):
        async with ClientSession(r, w) as s:
            init = await s.initialize()
            tools = await s.list_tools()
            print("servidor:", init.serverInfo.name, "| herramientas:", len(tools.tools))
            res = await s.call_tool("apa7_apply_apa", {"file_path": str(src)})
            print("apply_apa isError:", res.isError)
            print(res.content[0].text[:400])


asyncio.run(main())
