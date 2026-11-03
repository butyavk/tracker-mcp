"""add_applications and set_statuses: many rows, one preview, one write, one undo."""

import json
from datetime import date

import anyio
import openpyxl
import pytest
from mcp import Client
from mcp.types import ElicitResult

from tracker_mcp import server
from tracker_mcp.tracker import TrackerError

from .test_tracker import ru_tracker


@pytest.fixture
def t(tmp_path):
    return ru_tracker(tmp_path / "tracker.xlsx")


def rows_of(t):
    return {a["row"]: a for a in t.list_applications(limit=1000)["applications"]}


NEW = [
    {"company": "Vandelay", "role": "Head of Engineering", "link": "https://example.com/1",
     "notes": "remote", "extra": {"source": "scan"}},
    {"company": "Soylent", "role": "Engineering Manager, Payments"},
    {"company": "Wonka", "role": "Director of Engineering", "extra": {"Гейт — спросить первым": "after 24.11"}},
]


# ---------------------------------------------------------------- add_applications

def test_add_applications_appends_every_row_in_one_write(t):
    before = t.version()
    plan = t.add_applications(NEW, dry_run=True)
    assert plan["row"] == [11, 12, 13] and t.version() == before
    assert set(plan["labels"]) == {11, 12, 13}

    res = t.add_applications(NEW, expected_version=before)
    assert res["row"] == [11, 12, 13] and not res["dry_run"]
    got = rows_of(t)
    assert [got[r]["company"] for r in (11, 12, 13)] == ["Vandelay", "Soylent", "Wonka"]
    assert all(got[r]["status"] == "Не подавался" for r in (11, 12, 13))
    full = {a["row"]: a for a in t.get_application("Vandelay")["applications"]}
    assert full[11]["source"] == "scan" and full[11]["link"] == "https://example.com/1"
    assert t.get_application("Wonka")["applications"][0]["gate"] == "after 24.11"

    ws = openpyxl.load_workbook(t.path)["Подачи"]
    assert [ws.cell(r, 1).value for r in (11, 12, 13)] == [6, 7, 8]          # numbers go on
    assert ws.auto_filter.ref == "A5:U13"                                     # filter reaches the last row
    assert any("O6:O13" in str(dv.sqref) for dv in ws.data_validations.dataValidation)

    entries = t._audit_entries()
    assert len(entries) == 1 and entries[0]["tool"] == "add_applications"     # one write, one backup


def test_add_applications_is_one_undo_step(t):
    t.add_applications(NEW)
    t.undo_last_write()
    assert set(rows_of(t)) == {6, 7, 8, 9, 10}
    assert openpyxl.load_workbook(t.path)["Подачи"].auto_filter.ref == "A5:U10"


def test_add_applications_refuses_a_duplicate_and_writes_nothing(t):
    v0 = t.version()
    items = NEW + [{"company": "Acme", "role": "Engineering Manager, Agents"}]
    with pytest.raises(TrackerError, match=r"Item 4 \(Acme — Engineering Manager, Agents\).*Row 8.*Nothing"):
        t.add_applications(items)
    assert t.version() == v0 and not t._audit_entries()


def test_add_applications_sees_a_duplicate_inside_the_batch(t):
    twice = [NEW[1], {"company": "soylent", "role": "engineering manager, payments"}]
    with pytest.raises(TrackerError, match=r"Item 2 .*Row 11 already has"):
        t.add_applications(twice)
    res = t.add_applications([NEW[0], {**NEW[0], "allow_duplicate": True}])
    assert res["row"] == [11, 12]


def test_add_applications_same_link_is_a_duplicate(t):
    with pytest.raises(TrackerError, match="Item 2"):
        t.add_applications([NEW[0], {"company": "Other", "role": "EM", "link": "https://example.com/1"}])


@pytest.mark.parametrize("items, error", [
    ([], "No applications"),
    ([{"company": "", "role": "EM"}], "Item 1: company and role"),
    ([{"company": "X", "role": "EM", "extra": {"No such column": "1"}}], "Item 1"),
])
def test_add_applications_checks_its_items(t, items, error):
    with pytest.raises(TrackerError, match=error):
        t.add_applications(items)
    assert not t._audit_entries()


def test_batch_limit(t):
    many = [{"company": f"Co{i}", "role": "EM"} for i in range(t.BATCH_LIMIT + 1)]
    with pytest.raises(TrackerError, match="limit is 50"):
        t.add_applications(many)
    res = t.add_applications(many[:t.BATCH_LIMIT], dry_run=True)
    assert len(res["row"]) == t.BATCH_LIMIT


def test_add_applications_warns_about_the_same_company(t):
    res = t.add_applications([{"company": "Hooli", "role": "Head of Engineering"}], dry_run=True)
    assert res["warnings"] and res["warnings"][0].startswith("Item 1 (Hooli — Head of Engineering): Row 10")


# ---------------------------------------------------------------- set_statuses

def test_set_statuses_matches_set_status_row_by_row(tmp_path):
    (tmp_path / "a").mkdir()
    (tmp_path / "b").mkdir()
    a, b = ru_tracker(tmp_path / "a" / "t.xlsx"), ru_tracker(tmp_path / "b" / "t.xlsx")   # separate backups
    items = [{"row": 7, "status": "Отказ", "response": "template", "on": date(2026, 10, 17)},
             {"row": 8, "status": "Скрининг", "note": "call on Tuesday"},
             {"row": 9, "status": "Отозвал"}]
    batch = a.set_statuses(items)
    single = [b.set_status(i["row"], i["status"], response=i.get("response"), note=i.get("note"), on=i.get("on"))
              for i in items]
    assert batch["changes"] == [c for s in single for c in s["changes"]]
    assert batch["row"] == [7, 8, 9]
    assert rows_of(a) == rows_of(b)
    assert len(a._audit_entries()) == 1 and len(b._audit_entries()) == 3


def test_set_statuses_is_all_or_nothing(t):
    v0 = t.version()
    with pytest.raises(TrackerError, match=r"Row 8: Unknown status «Ждём».*Nothing was written"):
        t.set_statuses([{"row": 7, "status": "Отказ"}, {"row": 8, "status": "Ждём"}])
    with pytest.raises(TrackerError, match="Row 99"):
        t.set_statuses([{"row": 7, "status": "Отказ"}, {"row": 99, "status": "Отказ"}])
    with pytest.raises(TrackerError, match="For a new submission use mark_applied"):
        t.set_statuses([{"row": 9, "status": "Подана"}])
    assert t.version() == v0 and rows_of(t)[7]["status"] == "Подана"


@pytest.mark.parametrize("items, error", [
    ([], "No status changes"),
    ([{"row": 7, "status": "Отказ"}, {"row": 7, "status": "Скрининг"}], "Row 7 appears twice"),
    ([{"row": "7", "status": "Отказ"}], "row number"),
    ([{"row": 7, "status": " "}], "status is required"),
    ([{"row": 7, "status": "Отказ", "on": date(2026, 10, 19)}], "Row 7: 2026-10-19 is in the future"),
])
def test_set_statuses_checks_its_items(t, items, error):
    with pytest.raises(TrackerError, match=error):
        t.set_statuses(items)


def test_set_statuses_undo_reverts_the_whole_batch(t):
    t.set_statuses([{"row": 7, "status": "Отказ"}, {"row": 8, "status": "Отказ"}])
    t.undo_last_write()
    got = rows_of(t)
    assert got[7]["status"] == "Подана" and got[8]["status"] == "Подана"


def test_reply_dates_come_from_set_statuses_too(t):
    from tracker_mcp.doctor import reply_dates_from_audit
    t.set_statuses([{"row": 7, "status": "Отказ"}])
    labels = {7: t._label(*t._load()[1:3], 7)}
    assert reply_dates_from_audit(t, labels) == {7: date.today()}


# ---------------------------------------------------------------- over MCP

@pytest.fixture
def served(t):
    server._tracker = t
    yield t
    server._tracker = None


def test_batch_tools_need_one_token(served):
    t = served

    async def go():
        async with Client(server.mcp) as c:
            tools = {x.name: x for x in (await c.list_tools()).tools}
            assert "add_applications" in tools and "set_statuses" in tools
            assert tools["set_statuses"].annotations.read_only_hint is False
            assert tools["add_applications"].input_schema["properties"]["items"]["maxItems"] == 50

            args = {"items": NEW}
            res = await c.call_tool("add_applications", args)
            assert res.is_error and "confirmation" in res.content[0].text
            preview = await c.call_tool("add_applications", {**args, "dry_run": True})
            data = preview.structured_content
            assert data["preview"].startswith("add_applications: change 3 row(s)")
            for name in ("Vandelay", "Soylent", "Wonka"):
                assert name in data["preview"]
            res = await c.call_tool("add_applications", {**args, "confirm": data["confirm"]})
            assert not res.is_error and res.structured_content["row"] == [11, 12, 13]

            items = {"items": [{"row": 7, "status": "Отказ", "on": "2026-10-17"}, {"row": 11, "status": "Отозвал"}]}
            preview = await c.call_tool("set_statuses", {**items, "dry_run": True})
            token, text = preview.structured_content["confirm"], preview.structured_content["preview"]
            assert "Row 7" in text and "Row 11" in text
            res = await c.call_tool("set_statuses", {**items, "confirm": token})
            assert not res.is_error

    anyio.run(go)
    got = rows_of(t)
    assert got[7]["status"] == "Отказ" and got[11]["status"] == "Отозвал"
    assert [e["tool"] for e in t._audit_entries()] == ["add_applications", "set_statuses"]


def test_batch_asks_once_in_the_clients_dialog(served):
    t = served
    seen = []

    async def dialog(context, params):
        seen.append(params.message)
        return ElicitResult(action="accept", content={"confirm": True})

    async def go():
        async with Client(server.mcp, elicitation_callback=dialog) as c:
            res = await c.call_tool("set_statuses", {"items": [{"row": 7, "status": "Отказ"},
                                                               {"row": 8, "status": "Отказ"}]})
            assert not res.is_error

    anyio.run(go)
    assert len(seen) == 1 and "Row 7" in seen[0] and "Row 8" in seen[0]
    assert rows_of(t)[8]["status"] == "Отказ"
    assert json.dumps(t._audit_entries()[0]["row"]) == "[7, 8]"
