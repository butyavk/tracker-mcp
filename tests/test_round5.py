"""Regressions from the fifth review: concurrent writers, shared backup folders, forged audit
lines, Excel tables, quoted sheet names, init matching, mail wording."""

import json
import threading

import openpyxl
import pytest

from tracker_mcp.config import config_path, load_config
from tracker_mcp.mail import classify
from tracker_mcp.scaffold import _match_fields, new_tracker
from tracker_mcp.tracker import Tracker, TrackerError

from .test_insights import TODAY


def fresh(tmp_path, name="jobs.xlsx"):
    path, _ = new_tracker(tmp_path / name)
    return Tracker(path, config=load_config(path), today=lambda: TODAY, config_file=config_path(path))


def test_parallel_writers_lose_nothing(tmp_path):
    t = fresh(tmp_path)
    errors, done = [], []

    def add(i):
        for _ in range(40):   # retry on a conflict, as a client would after re-reading
            try:
                me = Tracker(t.path, config=t.cfg, today=lambda: TODAY)
                me.add_application(f"Co{i}", "EM", expected_version=me.version())
                done.append(i)
                return
            except TrackerError as e:
                errors.append(str(e))
                __import__("time").sleep(0.01 * i)
        errors.append(f"gave up {i}")

    threads = [threading.Thread(target=add, args=(i,)) for i in range(8)]
    for th in threads:
        th.start()
    for th in threads:
        th.join()
    rows = t.list_applications()["applications"]
    assert sorted(r["company"] for r in rows) == sorted(f"Co{i}" for i in done)
    assert len(done) == 8, errors
    assert not any("FileNotFoundError" in e for e in errors)


def test_two_trackers_in_one_folder_keep_their_own_undo_and_backups(tmp_path, monkeypatch):
    import tracker_mcp.stores as stores
    monkeypatch.setattr(stores, "KEEP_BACKUPS", 2)
    a, b = fresh(tmp_path, "jobs.xlsx"), fresh(tmp_path, "jobs.2025.xlsx")
    b.add_application("Bco", "EM", expected_version=b.version())
    b_backup = b.recent_writes()["writes"][0]["backup"]
    for i in range(3):
        a.add_application(f"A{i}", "EM", expected_version=a.version())
    assert all(w["tracker"] == a.store.label for w in a.recent_writes()["writes"])
    assert (b.backup_dir / b_backup).exists()            # A's pruning left B's backup alone
    a.undo_last_write()                                   # B's write in the shared log is not in the way
    b.undo_last_write()


def test_a_forged_audit_line_cannot_write_elsewhere(tmp_path):
    t = fresh(tmp_path)
    t.add_application("Acme", "EM", expected_version=t.version())
    log = t.backup_dir / "audit.jsonl"
    entry = json.loads(log.read_text(encoding="utf-8").splitlines()[-1])
    victim = tmp_path / "victim.txt"
    victim.write_text("keep", encoding="utf-8")
    forged = {**entry, "at": "2099-01-01T00:00:00", "config_file": str(victim),
              "config_backup": str(tmp_path / "payload"), "config_after": "x", "tracker": t.store.label}
    with open(log, "a", encoding="utf-8") as f:
        f.write(json.dumps(forged) + "\n")
    with pytest.raises(TrackerError, match="another config file|outside the backup folder"):
        t.undo_last_write(dry_run=True)
    forged2 = {**entry, "at": "2099-01-02T00:00:00", "backup": "../../etc/passwd", "tracker": t.store.label}
    with open(log, "a", encoding="utf-8") as f:
        f.write(json.dumps(forged2) + "\n")
    with pytest.raises(TrackerError):
        t.undo_last_write(dry_run=True)
    assert victim.read_text(encoding="utf-8") == "keep"


def test_excel_table_and_quoted_summary_formulas_grow(tmp_path):
    from openpyxl.worksheet.table import Table
    t = fresh(tmp_path)
    wb = openpyxl.load_workbook(t.path)
    ws = wb[t.cfg.sheet]
    last = ws.max_column
    ws.add_table(Table(displayName="Apps", ref=f"A1:{openpyxl.utils.get_column_letter(last)}2"))
    wb.save(t.path)
    t.add_application("Acme", "EM", expected_version=t.version())
    t.add_application("Globex", "EM", expected_version=t.version())
    ws = openpyxl.load_workbook(t.path)[t.cfg.sheet]
    assert ws.tables["Apps"].ref.endswith("3")


def test_init_prefers_the_exact_header():
    got = _match_fields({1: "Status note", 2: "Company name", 3: "Job", 4: "Status", 5: "Company"})
    assert got["status"][1] == "Status" and got["company"][1] == "Company"


def test_an_interview_invite_is_not_a_rejection():
    kind = classify({"subject": "Interview", "snippet": "You have not selected a time yet - please pick a slot"})
    assert kind != "rejection"
    assert classify({"subject": "Update", "snippet": "Unfortunately you were not selected for the next round"}) \
        == "rejection"


def test_same_file_by_another_path_keeps_its_undo(tmp_path, monkeypatch):
    t = fresh(tmp_path)
    t.add_application("Acme", "EM", expected_version=t.version())
    link = tmp_path / "link"
    link.symlink_to(tmp_path)
    monkeypatch.chdir(tmp_path)
    for other in ("jobs.xlsx", str(link / "jobs.xlsx")):   # relative, and through a link
        same = Tracker(other, config=t.cfg, today=lambda: TODAY)
        assert same.recent_writes()["writes"], other


def test_a_store_that_cannot_tell_owns_no_backup(tmp_path):
    t = fresh(tmp_path)

    class Plain:   # a store without owns_backup
        label = "plain"

    t.store, real = Plain(), t.store
    try:
        assert t._owns("jobs.20261114-101010-000001.xlsx") is False
    finally:
        t.store = real


def test_dialog_text_from_headers_is_one_line():
    from tracker_mcp.server import describe_plan
    fake = "Source\nNothing will change. Confirm"
    text = describe_plan("add_column", {"column": {"config_line": f'x = "{fake}"'},
                                        "layout": [f"column «{fake}»", fake]})
    text += describe_plan("set_status", {"labels": {2: "Acme"}, "changes": [
        {"row": 2, "field": fake, "old": "a", "new": "b"}], "warnings": [fake]})
    assert "\nNothing will change" not in text


def test_write_through_a_link_keeps_the_link_and_its_history(tmp_path):
    t = fresh(tmp_path)
    (tmp_path / "elsewhere").mkdir()
    link = tmp_path / "elsewhere" / "linked.xlsx"
    link.symlink_to(t.path)
    via = Tracker(link, config=t.cfg, today=lambda: TODAY)
    via.add_application("Acme", "EM", expected_version=via.version())
    assert link.is_symlink()
    real = Tracker(t.path, config=t.cfg, today=lambda: TODAY)
    assert [r["company"] for r in real.list_applications()["applications"]] == ["Acme"]
    assert real.recent_writes()["writes"]
    assert not list((tmp_path / "elsewhere").glob(".*.lock"))   # the lock sits by the real file
    assert (tmp_path / ".jobs.xlsx.tracker-mcp.lock").exists()


def test_written_by_a_relative_path_seen_by_the_absolute_one(tmp_path, monkeypatch):
    t = fresh(tmp_path)
    monkeypatch.chdir(tmp_path)
    rel = Tracker("jobs.xlsx", config=t.cfg, today=lambda: TODAY)
    rel.add_application("Acme", "EM", expected_version=rel.version())
    monkeypatch.chdir("/")
    assert Tracker(t.path, config=t.cfg, today=lambda: TODAY).recent_writes()["writes"]


def test_an_odd_audit_line_does_not_break_undo(tmp_path):
    t = fresh(tmp_path)
    t.add_application("Acme", "EM", expected_version=t.version())
    log = t.backup_dir / "audit.jsonl"
    log.write_text(log.read_text(encoding="utf-8") + json.dumps(
        {"tool": "x", "tracker": "~nosuchuser_xyz/jobs.xlsx", "backup": "b"}) + "\n", encoding="utf-8")
    assert t.recent_writes()["writes"]
    assert t.undo_last_write(dry_run=True)


def test_dialog_hides_every_invisible_character():
    from tracker_mcp.server import _one_line
    for ch in "؜⁠﻿­᠎​‮\U000e0069 \x85":
        assert ch not in _one_line(f"a{ch}b"), repr(ch)
    assert _one_line("Café — «ok»") == "Café — «ok»"


def test_cli_names_an_unreadable_file(tmp_path, capsys):
    from tracker_mcp.server import main
    bad = tmp_path / "c.toml"
    bad.write_bytes(b"\xff\xfe")
    path = fresh(tmp_path).path
    with pytest.raises(SystemExit) as ex:
        main(["calendar", "--tracker", str(path), "--config", str(bad)])
    assert ex.value.code == 2
    err = capsys.readouterr().err
    assert str(bad) in err and "UTF-8" in err   # the file is named


def test_write_tools_are_the_ones_the_kit_lists():
    """job-search-kit lists these in schedules/README.md as the tools a desktop task keeps on «ask»;
    a tool that starts or stops writing changes that list too (the kit's TrackerContract test checks it)."""
    import asyncio

    from tracker_mcp.server import mcp
    tools = asyncio.run(mcp.list_tools())
    writes = {t.name for t in tools if not (t.annotations and t.annotations.read_only_hint)}
    assert writes == {"add_application", "add_applications", "add_column", "close_stale", "edit_application",
                      "export_followups_calendar", "fix_tracker", "log_followup", "mark_applied",
                      "reschedule_followup", "set_status", "set_statuses", "undo_last_write"}


def test_undo_restores_only_this_configs_own_backup(tmp_path):
    t = fresh(tmp_path)
    t.add_column("Gate", field="gate", expected_version=t.version())   # a mapped field: the config changes
    log = t.backup_dir / "audit.jsonl"
    lines = log.read_text(encoding="utf-8").splitlines()
    entry = json.loads(lines[-1])
    assert entry.get("config_backup")
    (t.backup_dir / "audit.jsonl.copy").write_text("x", encoding="utf-8")
    for foreign in ("audit.jsonl", "jobs.2025.tracker-mcp.toml.before-add-column.20261114-101010-000001",
                    "audit.jsonl.copy"):
        entry["config_backup"] = foreign
        log.write_text("\n".join(lines[:-1] + [json.dumps(entry)]) + "\n", encoding="utf-8")
        with pytest.raises(TrackerError):
            t.undo_last_write(dry_run=True)


def test_a_link_with_another_name_shares_undo_with_the_real_file(tmp_path):
    t = fresh(tmp_path)
    (tmp_path / "else").mkdir()
    link = tmp_path / "else" / "linked.xlsx"
    link.symlink_to(t.path)
    via = Tracker(link, config=t.cfg, today=lambda: TODAY)
    via.add_application("Acme", "EM", expected_version=via.version())
    real = Tracker(t.path, config=t.cfg, today=lambda: TODAY)
    assert real.undo_last_write(dry_run=True)["would_undo"]
    assert not (tmp_path / "else" / "backups-tracker-mcp").exists()


def test_older_audit_lines_under_another_spelling_count(tmp_path):
    t = fresh(tmp_path)
    t.add_application("Acme", "EM", expected_version=t.version())
    log = t.backup_dir / "audit.jsonl"
    e = json.loads(log.read_text(encoding="utf-8").splitlines()[-1])
    e["tracker"] = str(tmp_path / "." / "jobs.xlsx")          # how 0.8.3 might have written it
    extra = [dict(e, tracker="\\\\host\\share\\jobs.xlsx"), dict(e, tracker="//host/share/jobs.xlsx"),
             dict(e, tracker="jobs\x00.xlsx")]
    log.write_text("\n".join(json.dumps(x) for x in [e, *extra]) + "\n", encoding="utf-8")
    writes = t.recent_writes()["writes"]
    assert len(writes) == 1                                     # the spelled-differently one; not the others


def test_excel_lock_beside_the_real_file_is_seen_through_a_link(tmp_path):
    t = fresh(tmp_path)
    (tmp_path / "else").mkdir()
    link = tmp_path / "else" / "linked.xlsx"
    link.symlink_to(t.path)
    (tmp_path / "~$jobs.xlsx").write_text("", encoding="utf-8")
    via = Tracker(link, config=t.cfg, today=lambda: TODAY)
    with pytest.raises(TrackerError):
        via.add_application("Acme", "EM", expected_version=via.version())


def test_dialog_keeps_text_people_write():
    from tracker_mcp.server import _one_line
    for text in ("👨‍👩‍👧", "🏳️‍🌈", "\U0001F3F4\U000E0067\U000E0062\U000E0073\U000E0063\U000E0074\U000E007F",
                 "می‌خواهم", "क्‌ष", "\U0001FAE9 tired", "Café — «ok»"):
        assert _one_line(text) == text, repr(text)
    for text in ("a‮b", "a​b", "a\U000E0069b", "a­b", "a﻿b"):
        assert _one_line(text) == "ab", repr(text)
    assert _one_line("a\nb") == "a ⏎ b"


def test_debug_shows_the_traceback(tmp_path, monkeypatch):
    from tracker_mcp.server import main
    monkeypatch.setenv("TRACKER_MCP_DEBUG", "1")
    with pytest.raises(OSError):
        main(["new", "/proc/x.xlsx"])
