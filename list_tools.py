"""
list_tools.py — print all MCP tool names + their input schemas.
Uses the saved token from .mcp_token.json (no OTP needed).
Run: python list_tools.py
"""
import asyncio, json, os, pathlib
from dotenv import load_dotenv
load_dotenv()

MCP_URL    = os.getenv("QME_MCP_BASE_URL", "https://retail.qand.me/api/mcp")
TOKEN_FILE = pathlib.Path(__file__).parent / ".mcp_token.json"

tok = json.loads(TOKEN_FILE.read_text())
access_token = tok["access_token"]
print(f"Using token: {access_token[:20]}…\n")

from mcp import ClientSession
from mcp.client.streamable_http import streamablehttp_client

async def main():
    headers = {"Authorization": f"Bearer {access_token}"}
    async with streamablehttp_client(MCP_URL, headers=headers) as (r, w, _):
        async with ClientSession(r, w) as session:
            await session.initialize()
            tools = await session.list_tools()
            for t in tools.tools:
                print(f"{'='*60}")
                print(f"Tool: {t.name}")
                print(f"Desc: {t.description}")
                schema = t.inputSchema if hasattr(t, 'inputSchema') else getattr(t, 'input_schema', None)
                if schema:
                    print(f"Schema: {json.dumps(schema, indent=2, ensure_ascii=False)}")
                print()

asyncio.run(main())
