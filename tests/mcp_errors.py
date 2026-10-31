"""Por stdio, como un agente real: un archivo inexistente debe devolver un error estructurado y el servidor debe seguir vivo."""
import asyncio
import json
import sys

from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client


async def main() -> None:
    params = StdioServerParameters(command=sys.executable, args=["-m", "apa7_mcp"])
    async with stdio_client(params) as (r, w):
        async with ClientSession(r, w) as s:
            await s.initialize()
            res = await s.call_tool("apa7_apply_apa", {"file_path": r"C:\no\existe.docx"})
            data = json.loads(res.content[0].text)
            assert data["ok"] is False and data["error"]["code"] == "FILE_NOT_FOUND", data
            res = await s.call_tool("apa7_insert_blocks", {"file_path": r"C:\no\existe.docx", "blocks": [{"type": "x"}]})
            assert json.loads(res.content[0].text)["ok"] is False
            # el servidor sigue respondiendo después de los errores
            res = await s.call_tool("apa7_doctor", {"check_network": False})
            assert json.loads(res.content[0].text)["ok"] is True
            tools = await s.list_tools()
            print("errores estructurados OK; herramientas:", len(tools.tools))


asyncio.run(main())
