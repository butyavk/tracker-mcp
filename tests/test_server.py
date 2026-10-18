"""The same calls through the MCP protocol (in-process client), including the
server-side confirmation of writes."""

import anyio
import pytest
from mcp import Client
from mcp.types import ElicitResult

from tracker_mcp import server

from .test_tracker import ru_tracker


@pytest.fixture
def t(tmp_path):
    server._tracker = ru_tracker(tmp_path / "t.xlsx")
    return server._tracker


def status_of(t, row):
    return t.get_application(row=row)["applications"][0]["status"]


def test_tools_over_mcp(t):
    async def go():
        async with Client(server.mcp) as c:
            tools = {x.name: x for x in (await c.list_tools()).tools}
            assert {"describe_tracker", "stale_applications", "close_stale"} <= set(tools)
            assert tools["summary"].annotations.read_only_hint is True
            assert tools["set_status"].annotations.read_only_hint is False
            assert "ctx" not in tools["set_status"].input_schema["properties"]

            res = await c.call_tool("followups_due", {"within_days": 7})
            assert res.structured_content["due"][0]["company"] == "Globex"

            res = await c.call_tool("set_status", {"row": 7, "status": "Ждём"})
            assert res.is_error and "Allowed" in res.content[0].text

    anyio.run(go)


@pytest.mark.parametrize("mode", ["legacy", "auto"])
def test_write_needs_the_token_when_the_client_has_no_dialog(t, mode):
    async def go():
        async with Client(server.mcp, mode=mode) as c:
            res = await c.call_tool("set_status", {"row": 7, "status": "Скрининг"})
            assert res.is_error and "confirmation" in res.content[0].text
            assert status_of(t, 7) == "Подана"

            preview = (await c.call_tool("set_status", {"row": 7, "status": "Скрининг", "dry_run": True}))
            token = preview.structured_content["confirm"]
            res = await c.call_tool("set_status", {"row": 7, "status": "Скрининг", "confirm": "0" * 12})
            assert res.is_error and status_of(t, 7) == "Подана"

            other = await c.call_tool("set_status", {"row": 7, "status": "Интервью", "confirm": token})
            assert other.is_error and status_of(t, 7) == "Подана"   # token is for another plan

            res = await c.call_tool("set_status", {"row": 7, "status": "Скрининг", "confirm": token})
            assert not res.is_error and status_of(t, 7) == "Скрининг"

            again = await c.call_tool("set_status", {"row": 7, "status": "Интервью", "confirm": token})
            assert again.is_error                                   # the file changed: token is stale

    anyio.run(go)


@pytest.mark.parametrize("mode", ["legacy", "auto"])   # mid-call request / input-required round
@pytest.mark.parametrize("answer", ["accept", "decline"])
def test_write_asks_the_user_when_the_client_has_a_dialog(t, answer, mode):
    seen = []

    async def dialog(context, params):
        seen.append(params.message)
        return ElicitResult(action=answer, content={"confirm": True} if answer == "accept" else None)

    async def go():
        async with Client(server.mcp, elicitation_callback=dialog, mode=mode) as c:
            res = await c.call_tool("set_status", {"row": 7, "status": "Отказ", "response": "template"})
            return res

    res = anyio.run(go)
    assert "Row 7 · Globex" in seen[0] and "status: Подана → Отказ" in seen[0]
    assert "response: ∅ → 18.10: template" in seen[0]
    if answer == "accept":
        assert not res.is_error and status_of(t, 7) == "Отказ"
    else:
        assert res.is_error and "did not confirm" in res.content[0].text and status_of(t, 7) == "Подана"
