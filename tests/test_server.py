"""The same calls through the MCP protocol (in-process client)."""

import anyio
from mcp import Client

from tracker_mcp import server

from .test_tracker import ru_tracker


def test_tools_over_mcp(tmp_path):
    server._tracker = ru_tracker(tmp_path / "t.xlsx")

    async def go():
        async with Client(server.mcp) as c:
            tools = {t.name: t for t in (await c.list_tools()).tools}
            assert "describe_tracker" in tools
            assert tools["summary"].annotations.read_only_hint is True
            assert tools["set_status"].annotations.read_only_hint is False

            res = await c.call_tool("followups_due", {"within_days": 7})
            assert res.structured_content["due"][0]["company"] == "Globex"

            res = await c.call_tool("set_status", {"row": 7, "status": "Скрининг", "dry_run": True})
            assert res.structured_content["dry_run"] is True

            res = await c.call_tool("set_status", {"row": 7, "status": "Ждём"})
            assert res.is_error and "Allowed" in res.content[0].text

    anyio.run(go)
