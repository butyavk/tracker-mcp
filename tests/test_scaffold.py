"""Getting started (`init`, `new`) and the funnel."""

import subprocess
import sys
import tomllib
from datetime import date

import openpyxl
import pytest
from openpyxl.worksheet.datavalidation import DataValidation

from tracker_mcp.config import load_config
from tracker_mcp.scaffold import guess, init, new_tracker
from tracker_mcp.tracker import Tracker, TrackerError, dropdown_values

from .test_tracker import RU_CONFIG, TODAY, make_tracker, ru_tracker


def test_init_rebuilds_my_hand_written_config(tmp_path):
    book = make_tracker(tmp_path / "t.xlsx")
    target, g, _ = init(book)
    assert target == tmp_path / "t.tracker-mcp.toml"
    hand = Tracker(book, config=load_config(book, RU_CONFIG)).describe()
    auto = Tracker(book, config=load_config(book)).describe()
    custom = {"source", "geography", "money", "folder", "prepared", "gate"}
    assert {k: v for k, v in hand["columns"].items() if k not in custom} == \
           {k: v for k, v in auto["columns"].items() if k != "source"}
    assert auto["columns"]["source"] == "D: Источник"
    assert (auto["header_row"], auto["data_rows"]) == (hand["header_row"], hand["data_rows"])

    h, a = load_config(book, RU_CONFIG), load_config(book)
    assert set(a.statuses.closed) == set(h.statuses.closed) - {"Закрыта"}  # not in this dropdown
    assert (a.statuses.not_applied, a.statuses.applied, a.statuses.no_response) == \
           (h.statuses.not_applied, h.statuses.applied, h.statuses.no_response)
    assert a.statuses.rejected == ["Отказ"] and a.statuses.stages == h.statuses.stages
    assert (a.values, a.messages, a.summary_sheet) == (h.values, h.messages, "Сводка")
    assert ("Гейт — спросить первым" in [hd for _, hd in g.unmapped]) and not g.checks


def test_init_on_a_sheet_without_dropdowns(tmp_path):
    path = tmp_path / "jobs.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Tab"
    ws.append(["Company", "Position", "Stage", "URL", "Applied", "Follow up", "Log", "Salary"])
    ws.append(["Acme", "EM", "Sent", "https://acme/1", "2026-10-01", None, None, "200k"])
    ws.append(["Globex", "Head of Eng", "Draft", None, None, None, None, None])
    wb.save(path)
    target, g, text = init(path)
    cfg = load_config(path)
    assert cfg.columns["applied_on"] == "Applied"          # a column of dates is the date of applying
    assert "applied" not in cfg.columns and cfg.columns["role"] == "Position"
    assert cfg.statuses.fallback == ["Sent", "Draft"] and cfg.statuses.applied == "Sent"
    assert cfg.statuses.not_applied == "Draft"
    assert any("no_response" in c for c in g.checks) and "# CHECK" in text
    assert "# salary = \"Salary\"" in text
    assert Tracker(path, config=cfg).summary()["applied"] == 1


def test_init_does_not_overwrite(tmp_path):
    book = make_tracker(tmp_path / "t.xlsx")
    init(book)
    with pytest.raises(TrackerError, match="already exists"):
        init(book)
    _, _, text = init(book, write=False)
    assert tomllib.loads(text)["sheet"] == "Подачи"
    with pytest.raises(TrackerError, match="No header row"):
        wb = openpyxl.Workbook()
        wb.active.append(["Name", "Phone"])
        wb.save(tmp_path / "contacts.xlsx")
        guess(tmp_path / "contacts.xlsx")


def test_dropdown_from_a_range_on_another_sheet(tmp_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Company", "Role", "Status"])
    lists = wb.create_sheet("Lists")
    for i, s in enumerate(["To apply", "Applied", "Ghosted", "Rejected"], 1):
        lists.cell(i, 1, s)
    dv = DataValidation(type="list", formula1="Lists!$A$1:$A$4")
    dv.add("C2:C100")
    ws.add_data_validation(dv)
    wb.save(tmp_path / "r.xlsx")
    ws = openpyxl.load_workbook(tmp_path / "r.xlsx").active
    assert dropdown_values(ws, 3) == ["To apply", "Applied", "Ghosted", "Rejected"]
    g = guess(tmp_path / "r.xlsx")
    assert g.meaning["no_response"] == ["Ghosted"] and g.meaning["not_applied"] == ["To apply"]


def test_new_tracker_works_end_to_end(tmp_path):
    book, cfg_path = new_tracker(tmp_path / "jobs.xlsx")
    cfg = load_config(book)
    assert cfg.columns["reached"] == "Furthest stage" and cfg.summary_sheet == "Summary"
    t = Tracker(book, config=cfg, today=lambda: TODAY)
    for co in ("Acme", "Globex", "Initech"):
        t.add_application(co, "Engineering Manager", extra={"source": "LinkedIn"})
    ws = openpyxl.load_workbook(book)["Applications"]
    assert ws["A2"].value == 1 and not ws["B2"].font.b          # not styled like the header
    t.mark_applied(2, channel="Ashby (direct form)")
    t.mark_applied(3, channel="ashby")
    t.mark_applied(4, channel="Greenhouse")
    t.set_status(2, "Interview")
    t.set_status(2, "Rejected")
    t.set_status(3, "Screening")
    got = t.get_application(row=2)["applications"][0]
    assert (got["status"], got["reached"]) == ("Rejected", "Interview")  # the stage stays reached

    f = t.funnel()
    assert f["total"] == {**f["total"], "applied": 3, "replied": 2, "screening": 2, "interview": 1, "offer": 0}
    assert [(g["group"], g["applied"], g["interview"]) for g in f["groups"]] == [("Ashby", 2, 1),
                                                                               ("Greenhouse", 1, 0)]
    assert f["total"]["rates"]["replied"] == 66.7
    with pytest.raises(TrackerError, match="already exists"):
        new_tracker(book)


def test_funnel_on_my_layout(tmp_path):
    t = ru_tracker(tmp_path / "t.xlsx")
    t.set_status(7, "Тех.этап")
    t.set_status(7, "Отказ")          # no `reached` column here: counts as a reply only
    f = t.funnel("month")
    assert [g["group"] for g in f["groups"]] == ["2026-09", "2026-10"]
    assert f["total"]["applied"] == 4 and f["total"]["replied"] == 2 and f["total"]["interview"] == 0
    assert "current status only" in f["note"]
    assert t.funnel(since=date(2026, 10, 1))["total"]["applied"] == 3
    with pytest.raises(TrackerError, match="Cannot group by"):
        t.funnel("salary")


def test_cli(tmp_path):
    book = make_tracker(tmp_path / "t.xlsx")
    run = lambda *a: subprocess.run([sys.executable, "-m", "tracker_mcp.server", *a],  # noqa: E731
                                    capture_output=True, text=True)
    out = run("init", str(book))
    assert out.returncode == 0, out.stderr
    assert "Компания" in out.stdout and "Next:" in out.stdout
    again = run("init", str(book))
    assert again.returncode == 2 and "already exists" in again.stderr
    assert "header_row = 5" in run("init", str(book), "--print").stdout
    created = run("new", str(tmp_path / "fresh.xlsx"))
    assert created.returncode == 0 and (tmp_path / "fresh.tracker-mcp.toml").exists()
