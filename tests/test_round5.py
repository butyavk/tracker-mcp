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
