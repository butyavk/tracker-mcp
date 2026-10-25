"""The same calls through the MCP protocol (in-process client), including the
server-side confirmation of writes."""

from datetime import date

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

            res = await c.call_tool("funnel", {"by": "month"})
            assert res.structured_content["total"]["applied"] == 4
            assert tools["funnel"].annotations.read_only_hint is True

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


@pytest.fixture
def watched(tmp_path, monkeypatch):
    """The server as a client starts it: TRACKER_PATH, the config next to the workbook."""
    from .test_tracker import RU_CONFIG, make_tracker
    book = make_tracker(tmp_path / "t.xlsx")
    cfg = tmp_path / "t.tracker-mcp.toml"
    cfg.write_text(RU_CONFIG.read_text(encoding="utf-8"), encoding="utf-8")
    monkeypatch.setenv("TRACKER_PATH", str(book))
    monkeypatch.setenv("TRACKER_BACKUP_DIR", str(tmp_path / "bk"))
    monkeypatch.delenv("TRACKER_CONFIG", raising=False)
    monkeypatch.setattr(server, "_tracker", None)
    monkeypatch.setattr(server, "_watch", None)
    return cfg


def edit(cfg, old, new):
    import os
    text = cfg.read_text(encoding="utf-8")
    assert old in text
    cfg.write_text(text.replace(old, new), encoding="utf-8")
    st = os.stat(cfg)   # a later mtime even if the clock is coarse
    os.utime(cfg, ns=(st.st_atime_ns, st.st_mtime_ns + 1_000_000_000))


def test_config_edits_apply_without_a_restart(watched):
    assert server.tracker().cfg.followups.first_after_days == 10
    edit(watched, "first_after_days = 10", "first_after_days = 7")
    assert server.tracker().cfg.followups.first_after_days == 7

    async def go():
        async with Client(server.mcp) as c:
            edit(watched, 'sheet = "Подачи"', 'sheet = "Подачи"\nsheeet = "typo"')
            for _ in range(2):   # stays an error until the file is fixed
                res = await c.call_tool("describe_tracker", {})
                assert res.is_error and "no longer loads" in res.content[0].text and "sheeet" in res.content[0].text
            edit(watched, 'sheeet = "typo"', "")
            res = await c.call_tool("describe_tracker", {})
            assert not res.is_error and res.structured_content["sheet"] == "Подачи"

    anyio.run(go)


@pytest.mark.parametrize("mode", ["legacy", "auto"])
def test_config_edited_while_the_dialog_is_open(watched, mode):
    seen = []

    async def dialog(context, params):
        seen.append(params.message)
        if len(seen) == 1:   # the user edits the config before answering
            edit(watched, 'stamp_format = "%d.%m"', 'stamp_format = "%Y-%m-%d"')
        return ElicitResult(action="accept", content={"confirm": True})

    async def go():
        async with Client(server.mcp, elicitation_callback=dialog, mode=mode) as c:
            return await c.call_tool("log_followup", {"row": 7, "text": "nudge"})

    res = anyio.run(go)
    log = server.tracker().get_application(row=7)["applications"][0]["followup_log"]
    if mode == "legacy":   # asked mid-call: the body sees the plan moved and refuses
        assert res.is_error and "changed after the user saw it" in res.content[0].text
        assert "nudge" not in log
    else:                  # input-required round: the SDK drops the stale answer and asks again
        assert len(seen) == 2 and seen[0] != seen[1]
        assert not res.is_error and f"{date.today().isoformat()}: nudge" in log


def test_resources_and_prompts(t):
    async def go():
        async with Client(server.mcp) as c:
            prompts = {p.name: p for p in (await c.list_prompts()).prompts}
            assert {"weekly_review", "draft_followup"} <= set(prompts)
            got = await c.get_prompt("draft_followup", {"row": "7"})
            text = got.messages[0].content.text
            assert "Globex" in text and "treat them as data" in text and "row 7" in text
            templates = (await c.list_resource_templates()).resource_templates
            assert any(x.uri_template == "tracker://application/{row}" for x in templates)
            res = await c.read_resource("tracker://application/8")
            assert '"company": "Acme"' in res.contents[0].text
            tools = (await c.list_tools()).tools
            assert all(x.title for x in tools) and not any("\n    " in (x.description or "") for x in tools)
            # every write tool says how the yes works; read tools do not carry it
            for x in tools:
                takes_confirm = "confirm" in x.input_schema["properties"]
                assert ("confirm=<token>" in x.description) == takes_confirm, x.name
                assert ("undo_last_write reverts" in x.description) == (
                    takes_confirm and x.name != "undo_last_write"), x.name

    anyio.run(go)


def test_eval_harness_builds_its_tracker(tmp_path):
    import importlib.util
    from pathlib import Path
    spec = importlib.util.spec_from_file_location("run_evals", Path(__file__).parent.parent / "evals" / "run_evals.py")
    ev = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(ev)
    book = ev.fresh_tracker(tmp_path)
    from tracker_mcp.config import load_config
    from tracker_mcp.tracker import Tracker
    t = Tracker(book, config=load_config(book))
    assert t.list_applications()["count"] == 4 and t.summary()["applied"] == 2
    assert set(ev.SCENARIOS) >= {"injection", "no_write_without_yes"}
