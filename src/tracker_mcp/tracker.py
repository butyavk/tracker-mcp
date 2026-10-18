"""Read and safely write a job-application tracker kept in an .xlsx file.

The workbook stays the source of truth: I open it in Excel, and other scripts
edit it too. So the server never caches it. Every call re-reads the file, and
every write goes through `Tracker._write`, which:

1. refuses while the file is open in Excel or LibreOffice (lock file present);
2. refuses if the file changed since the caller last read it: `expected_version`
   works like an HTTP If-Match, the version is a hash of the file's bytes;
3. saves to a temp file, re-opens it and checks every intended cell;
4. backs up the current file, then swaps the new one in with an atomic rename;
5. appends an audit record (field, old -> new, backup name).

Which sheet, which columns, which status words and which follow-up rule: all of
that comes from the config (config.py). Columns are found by their header text,
not by letter, so inserting a column in Excel does not break the server.
"""

from __future__ import annotations

import hashlib
import json
import os
import re
import shutil
from collections import Counter
from collections.abc import Callable
from copy import copy
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from io import BytesIO
from pathlib import Path
from typing import Any

import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.cell_range import MultiCellRange
from openpyxl.worksheet.worksheet import Worksheet

from .config import Config

EXCEL_EPOCH = date(1899, 12, 30)
KEEP_BACKUPS = 200


class TrackerError(Exception):
    """A request the tracker cannot apply; the message says why."""


class ConflictError(TrackerError):
    """The file changed (or is open) under us; re-read and retry."""


# ---------------------------------------------------------------- values

def parse_date(value: Any) -> date | None:
    """Dates in a hand-kept sheet are mixed: datetimes, ISO strings, Excel serials."""
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, (int, float)) and 30000 < value < 80000:
        return EXCEL_EPOCH + timedelta(days=int(value))
    if isinstance(value, str):
        s = value.strip()
        if s.isdigit() and 30000 < int(s) < 80000:
            return EXCEL_EPOCH + timedelta(days=int(s))
        try:
            if m := re.match(r"(\d{4})-(\d{2})-(\d{2})", s):
                return date(*map(int, m.groups()))
            if m := re.match(r"(\d{1,2})\.(\d{1,2})\.(\d{4})", s):
                d, mo, y = map(int, m.groups())
                return date(y, mo, d)
        except ValueError:
            return None
    return None


def weekday_or_monday(d: date) -> date:
    return d + timedelta(days=(7 - d.weekday()) % 7) if d.weekday() >= 5 else d


def _hash(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()[:16]


@dataclass
class Change:
    row: int
    field: str
    old: Any
    new: Any  # the raw value written: str, int or date

    def as_dict(self, show: Callable[[Any, str], Any]) -> dict[str, Any]:
        return {"row": self.row, "field": self.field, "old": self.old, "new": show(self.new, self.field)}


@dataclass
class Layout:
    header_row: int
    cols: dict[str, int]       # field -> column number
    headers: dict[int, str]    # column number -> header text
    last_row: int              # last row with a company name

    def rows(self) -> range:
        return range(self.header_row + 1, self.last_row + 1)


# ---------------------------------------------------------------- tracker

class Tracker:
    def __init__(self, path: str | Path, *, config: Config, backup_dir: str | Path | None = None,
                 today: Callable[[], date] = date.today):
        self.path = Path(path).expanduser()
        self.cfg = config
        self.backup_dir = (Path(backup_dir).expanduser() if backup_dir
                           else self.path.parent / "backups-tracker-mcp")
        self.today = today
        st = config.statuses
        self.closed = set(st.closed)
        self.waiting = set(st.waiting or [st.applied])
        self.aliases = {k.casefold(): v for k, v in st.aliases.items()}

    # ---- values

    def _show(self, value: Any, key: str = "") -> Any:
        """Cell value -> JSON-friendly value."""
        if key in self.cfg.date_columns:
            d = parse_date(value)
            if d:
                return d.isoformat()
        if isinstance(value, (datetime, date)):
            return (value.date() if isinstance(value, datetime) else value).isoformat()
        if isinstance(value, str):
            return value.strip() or None
        return value

    def _append(self, old: Any, text: str) -> str:
        old = str(old).strip() if old is not None else ""
        return f"{old}{self.cfg.values.separator}{text}" if old else text

    def _stamp(self, text: str) -> str:
        return f"{self.today().strftime(self.cfg.values.stamp_format)}: {text.strip()}"

    def _fmt(self, d: date) -> str:
        return d.strftime(self.cfg.values.stamp_format)

    def _next_date(self, start: date, days: int) -> date:
        d = start + timedelta(days=days)
        return weekday_or_monday(d) if self.cfg.followups.weekend_to_monday else d

    def _is_applied(self, rec: dict[str, Any]) -> bool:
        if self.cfg.has("applied"):
            return rec.get("applied") == self.cfg.values.yes
        return rec.get("status") not in (None, self.cfg.statuses.not_applied)

    def _need(self, *fields: str) -> None:
        missing = [f for f in fields if not self.cfg.has(f)]
        if missing:
            raise TrackerError(f"This needs the column(s) {', '.join(missing)}, "
                               "which the config does not define.")

    # ---- reading

    def version(self) -> str:
        return _hash(self.path.read_bytes())

    def _layout(self, ws: Worksheet) -> Layout:
        cfg = self.cfg
        header_row = cfg.header_row or next(
            (r for r in range(1, 31) for c in range(1, ws.max_column + 1)
             if str(ws.cell(r, c).value or "").strip() == cfg.columns["company"]), None)
        if header_row is None:
            raise TrackerError(f"Header row not found: no cell «{cfg.columns['company']}» "
                               f"in the first 30 rows of «{cfg.sheet}»")
        headers = {c: str(ws.cell(header_row, c).value).strip()
                   for c in range(1, ws.max_column + 1) if ws.cell(header_row, c).value}
        cols = {}
        for key, title in cfg.columns.items():
            found = ([c for c, h in headers.items() if h == title]
                     or [c for c, h in headers.items() if h.startswith(title)])
            if not found:
                raise TrackerError(f"Column «{title}» ({key}) not found in row {header_row}")
            cols[key] = found[0]
        last_row = header_row
        for r in range(header_row + 1, ws.max_row + 1):
            if ws.cell(r, cols["company"]).value not in (None, ""):
                last_row = r
        return Layout(header_row, cols, headers, last_row)

    def _load(self) -> tuple[Workbook, Worksheet, Layout, str]:
        """Read the file once: the workbook and its version come from the same bytes."""
        if not self.path.exists():
            raise TrackerError(f"Tracker file not found: {self.path}")
        data = self.path.read_bytes()
        wb = openpyxl.load_workbook(BytesIO(data))
        if self.cfg.sheet not in wb.sheetnames:
            raise TrackerError(f"Sheet «{self.cfg.sheet}» not found")
        ws = wb[self.cfg.sheet]
        return wb, ws, self._layout(ws), _hash(data)

    def _record(self, ws: Worksheet, lay: Layout, r: int, full: bool = False) -> dict[str, Any]:
        rec: dict[str, Any] = {"row": r}
        for key, c in lay.cols.items():
            rec[key] = self._show(ws.cell(r, c).value, key)
        if full:
            known = set(lay.cols.values())
            other = {h: self._show(ws.cell(r, c).value) for c, h in lay.headers.items()
                     if c not in known and ws.cell(r, c).value not in (None, "")}
            if other:
                rec["other"] = other
        return rec

    def records(self) -> tuple[list[dict[str, Any]], str]:
        _, ws, lay, version = self._load()
        recs = [self._record(ws, lay, r) for r in lay.rows()
                if ws.cell(r, lay.cols["company"]).value not in (None, "")]
        return recs, version

    @staticmethod
    def _label(ws: Worksheet, lay: Layout, r: int) -> str:
        company = str(ws.cell(r, lay.cols["company"]).value or "").strip()
        role = str(ws.cell(r, lay.cols["role"]).value or "").strip()
        return f"{company} — {role}" if role else company

    @staticmethod
    def _brief(rec: dict[str, Any]) -> dict[str, Any]:
        keys = ("row", "num", "company", "role", "status", "applied_on", "followup_on")
        return {k: rec[k] for k in keys if rec.get(k) not in (None, "")}

    def describe(self) -> dict[str, Any]:
        """How the server reads this workbook: columns, statuses, follow-up rule."""
        _, ws, lay, version = self._load()
        st, fu = self.cfg.statuses, self.cfg.followups
        return {
            "file": str(self.path),
            "sheet": self.cfg.sheet,
            "header_row": lay.header_row,
            "data_rows": f"{lay.header_row + 1}–{lay.last_row}",
            "columns": {k: f"{get_column_letter(c)}: {lay.headers[c]}" for k, c in lay.cols.items()},
            "custom_fields": self.cfg.custom_fields(),
            "statuses_allowed": self._allowed_statuses(ws, lay),
            "status_not_applied": st.not_applied,
            "status_applied": st.applied,
            "statuses_closed": st.closed,
            "followups": {"first_after_days": fu.first_after_days, "next_after_days": fu.next_after_days,
                          "weekend_to_monday": fu.weekend_to_monday},
            "version": version,
        }

    def list_applications(self, status: str | None = None, query: str | None = None,
                          applied_since: date | None = None, applied_until: date | None = None,
                          open_only: bool = False, limit: int = 100) -> dict[str, Any]:
        recs, version = self.records()
        wanted = {s.strip().casefold() for s in status.split(",")} if status else None
        q = query.casefold() if query else None
        out = []
        for rec in recs:
            st = (rec["status"] or "").casefold()
            if wanted and st not in wanted:
                continue
            if open_only and rec["status"] in self.closed:
                continue
            if q and q not in f"{rec['company']} {rec['role']}".casefold():
                continue
            if applied_since or applied_until:
                d = parse_date(rec.get("applied_on"))
                if (not self._is_applied(rec) or d is None
                        or (applied_since and d < applied_since)
                        or (applied_until and d > applied_until)):
                    continue
            out.append(self._brief(rec))
        return {"count": len(out), "applications": out[:limit],
                "truncated": len(out) > limit, "version": version}

    def get_application(self, query: str | None = None, row: int | None = None) -> dict[str, Any]:
        if query is None and row is None:
            raise TrackerError("Pass a company/role text or a row number")
        _, ws, lay, version = self._load()
        if row is not None:
            self._check_row(ws, lay, row)
            rows = [row]
        else:
            q = query.casefold()
            rows = [r for r in lay.rows()
                    if q in f"{ws.cell(r, lay.cols['company']).value or ''} "
                            f"{ws.cell(r, lay.cols['role']).value or ''}".casefold()]
        return {"count": len(rows), "applications": [self._record(ws, lay, r, full=True) for r in rows],
                "version": version}

    def followups_due(self, within_days: int = 7) -> dict[str, Any]:
        self._need("followup_on")
        recs, version = self.records()
        today = self.today()
        horizon = today + timedelta(days=within_days)
        due, undated = [], []
        for rec in recs:
            if rec["status"] in self.closed or rec["status"] == self.cfg.statuses.not_applied:
                continue
            d = parse_date(rec["followup_on"])
            if d is None:
                if self._is_applied(rec) and not rec["followup_on"]:
                    undated.append(self._brief(rec))
                continue
            if d <= horizon:
                log = rec.get("followup_log") or ""
                due.append({**self._brief(rec), "days_left": (d - today).days, "overdue": d < today,
                            "contact": rec.get("followup_contact"),
                            "last_log": log[-300:] if log else None})
        due.sort(key=lambda x: x["followup_on"])
        return {"today": today.isoformat(), "within_days": within_days, "due": due,
                "applied_without_followup_date": undated, "version": version}

    def summary(self, since: date | None = None, until: date | None = None) -> dict[str, Any]:
        recs, version = self.records()
        applied = [r for r in recs if self._is_applied(r)]
        months = Counter()
        undated = 0
        for r in applied:
            d = parse_date(r.get("applied_on"))
            if d:
                months[d.strftime("%Y-%m")] += 1
            else:
                undated += 1
        out: dict[str, Any] = {
            "rows": len(recs),
            "applied": len(applied),
            "not_applied": len(recs) - len(applied),
            "by_status": dict(Counter(r["status"] or "—" for r in recs).most_common()),
            "applied_by_month": dict(sorted(months.items())),
            "applied_without_date": undated,
            "version": version,
        }
        if since or until:
            cohort = [r for r in applied if (d := parse_date(r.get("applied_on")))
                      and (not since or d >= since) and (not until or d <= until)]
            out["period"] = {
                "since": since.isoformat() if since else None,
                "until": until.isoformat() if until else None,
                "applied": len(cohort),
                "by_status_now": dict(Counter(r["status"] or "—" for r in cohort).most_common()),
                "companies": [f"{r['company']} — {r['role']}" for r in cohort],
            }
        return out

    def _stale_rows(self, ws: Worksheet, lay: Layout, days: int) -> list[dict[str, Any]]:
        """Still waiting, submitted at least `days` ago, and no follow-up planned ahead."""
        self._need("applied_on")
        today = self.today()
        out = []
        for r in lay.rows():
            if ws.cell(r, lay.cols["company"]).value in (None, ""):
                continue
            rec = self._record(ws, lay, r)
            applied = parse_date(rec.get("applied_on"))
            fu = parse_date(rec.get("followup_on"))
            if (rec["status"] in self.waiting and self._is_applied(rec) and applied
                    and (today - applied).days >= days and not (fu and fu >= today)):
                log = rec.get("followup_log") or ""
                out.append({**self._brief(rec), "days_since_applied": (today - applied).days,
                            "last_log": log[-200:] if log else None})
        return out

    def stale_applications(self, days: int = 30) -> dict[str, Any]:
        _, ws, lay, version = self._load()
        rows = self._stale_rows(ws, lay, days)
        return {"days": days, "count": len(rows), "applications": rows,
                "would_set_status": self.cfg.statuses.no_response, "version": version}

    def close_stale(self, days: int = 30, *, rows: list[int] | None = None,
                    expected_version: str | None = None, dry_run: bool = False) -> dict[str, Any]:
        """Set the no-response status on every stale row (or the listed subset) in one write."""
        cfg = self.cfg

        def mutate(wb: Workbook, ws: Worksheet, lay: Layout):
            stale = [x["row"] for x in self._stale_rows(ws, lay, days)]
            if rows is not None:
                extra = sorted(set(rows) - set(stale))
                if extra:
                    raise TrackerError(f"Rows {extra} are not stale by the rule (waiting status, submitted "
                                       f"{days}+ days ago, no follow-up planned ahead).")
                stale = [r for r in stale if r in set(rows)]
            if not stale:
                return [], [], []
            status = self._normalize_status(ws, lay, cfg.statuses.no_response)
            text = cfg.messages.stale_closed.format(days=days)
            changes: list[Change] = []
            for r in stale:
                self._set(ws, lay, r, "status", status, changes)
                if cfg.has("followup_on"):
                    self._set(ws, lay, r, "followup_on", cfg.values.no_date, changes)
                self._log(ws, lay, r, text, changes)
            return stale, changes, []

        return self._write("close_stale", mutate, expected_version, dry_run)

    def recent_writes(self, limit: int = 10) -> dict[str, Any]:
        entries = self._audit_entries()
        return {"writes": entries[-limit:][::-1]}

    # ---- writing: the tools

    def add_application(self, company: str, role: str, *, link: str | None = None,
                        notes: str | None = None, extra: dict[str, str] | None = None,
                        allow_duplicate: bool = False, expected_version: str | None = None,
                        dry_run: bool = False) -> dict[str, Any]:
        company, role = company.strip(), role.strip()
        if not company or not role:
            raise TrackerError("company and role are required")
        cfg = self.cfg

        def mutate(wb: Workbook, ws: Worksheet, lay: Layout):
            extra_cols = self._resolve_extra(lay, extra or {})
            if not allow_duplicate:
                for r in lay.rows():
                    same = (str(ws.cell(r, lay.cols["company"]).value or "").strip().casefold() == company.casefold()
                            and str(ws.cell(r, lay.cols["role"]).value or "").strip().casefold() == role.casefold())
                    same_link = (link and cfg.has("link")
                                 and str(ws.cell(r, lay.cols["link"]).value or "").strip() == link.strip())
                    if same or same_link:
                        raise TrackerError(f"Row {r} already has this application. "
                                           "Pass allow_duplicate=true to add anyway.")
            r = lay.last_row + 1
            self._copy_row_style(ws, lay, src=lay.last_row, dst=r)
            values: dict[str, Any] = {"company": company, "role": role, "link": link, "notes": notes}
            if cfg.has("num"):
                nums = [int(v) for x in lay.rows()
                        if str(v := ws.cell(x, lay.cols["num"]).value or "").strip().isdigit()]
                values = {"num": max(nums, default=0) + 1, **values}
            values.update({"applied": cfg.values.no, "status": cfg.statuses.not_applied})
            changes: list[Change] = []
            for key, val in values.items():
                if val not in (None, "") and cfg.has(key):
                    self._set(ws, lay, r, key, val, changes)
            for col, val in extra_cols.items():
                self._set_col(ws, lay, r, col, val, changes)
            layout_notes = self._grow_layout(wb, ws, lay, r)
            return r, changes, layout_notes

        return self._write("add_application", mutate, expected_version, dry_run)

    def mark_applied(self, row: int, *, applied_on: date | None = None, channel: str | None = None,
                     followup_on: date | None = None, note: str | None = None,
                     expected_version: str | None = None, dry_run: bool = False) -> dict[str, Any]:
        cfg, msg = self.cfg, self.cfg.messages

        def mutate(wb: Workbook, ws: Worksheet, lay: Layout):
            self._check_row(ws, lay, row)
            rec = self._record(ws, lay, row)
            if self._is_applied(rec) and (cfg.has("applied") or rec.get("applied_on")):
                raise TrackerError(f"Row {row} is already marked as applied ({rec.get('applied_on')}).")
            day = applied_on or self.today()
            changes: list[Change] = []
            if cfg.has("applied"):
                self._set(ws, lay, row, "applied", cfg.values.yes, changes)
            if cfg.has("applied_on"):
                self._set(ws, lay, row, "applied_on", day, changes)
            if channel and cfg.has("channel"):
                self._set(ws, lay, row, "channel", channel, changes)
            if ws.cell(row, lay.cols["status"]).value in (None, "", cfg.statuses.not_applied):
                self._set(ws, lay, row, "status", cfg.statuses.applied, changes)
            if cfg.has("followup_on"):
                current_fu = ws.cell(row, lay.cols["followup_on"]).value
                if followup_on or current_fu in (None, ""):
                    fu = followup_on or self._next_date(day, cfg.followups.first_after_days)
                    self._set(ws, lay, row, "followup_on", fu, changes)
                    self._log(ws, lay, row, msg.applied_with_followup.format(date=self._fmt(fu)), changes)
                else:
                    self._log(ws, lay, row, msg.applied, changes)
            if note and cfg.has("notes"):
                self._set(ws, lay, row, "notes",
                          self._append(ws.cell(row, lay.cols["notes"]).value, self._stamp(note)), changes)
            return row, changes, []

        return self._write("mark_applied", mutate, expected_version, dry_run)

    def set_status(self, row: int, status: str, *, response: str | None = None,
                   note: str | None = None, expected_version: str | None = None,
                   dry_run: bool = False) -> dict[str, Any]:
        cfg = self.cfg

        def mutate(wb: Workbook, ws: Worksheet, lay: Layout):
            self._check_row(ws, lay, row)
            st = self._normalize_status(ws, lay, status)
            if st == cfg.statuses.applied and not self._is_applied(self._record(ws, lay, row)):
                raise TrackerError("For a new submission use mark_applied: it also sets "
                                   "the date and the follow-up date.")
            changes: list[Change] = []
            self._set(ws, lay, row, "status", st, changes)
            for key, text in (("response", response), ("notes", note)):
                if text and cfg.has(key):
                    self._set(ws, lay, row, key,
                              self._append(ws.cell(row, lay.cols[key]).value, self._stamp(text)), changes)
            if cfg.has("followup_on"):
                fu = parse_date(ws.cell(row, lay.cols["followup_on"]).value)
                if st in self.closed and fu and fu >= self.today():
                    self._set(ws, lay, row, "followup_on", cfg.values.no_date, changes)
                    self._log(ws, lay, row,
                              cfg.messages.followup_cancelled.format(date=self._fmt(fu), status=st), changes)
            return row, changes, []

        return self._write("set_status", mutate, expected_version, dry_run)

    def log_followup(self, row: int, text: str, *, next_on: date | None = None, final: bool = False,
                     contact: str | None = None, expected_version: str | None = None,
                     dry_run: bool = False) -> dict[str, Any]:
        self._need("followup_on", "followup_log")
        cfg = self.cfg

        def mutate(wb: Workbook, ws: Worksheet, lay: Layout):
            self._check_row(ws, lay, row)
            changes: list[Change] = []
            self._log(ws, lay, row, text, changes)
            nxt = cfg.values.no_date if final else (
                next_on or self._next_date(self.today(), cfg.followups.next_after_days))
            self._set(ws, lay, row, "followup_on", nxt, changes)
            if contact and cfg.has("followup_contact"):
                self._set(ws, lay, row, "followup_contact", contact, changes)
            return row, changes, []

        return self._write("log_followup", mutate, expected_version, dry_run)

    def undo_last_write(self, expected_version: str | None = None, dry_run: bool = False) -> dict[str, Any]:
        """Restore the backup of the last write, only if nothing changed since."""
        entries = self._audit_entries()
        entry = next((e for e in reversed(entries) if e["tool"] != "undo_last_write"), None)
        if entry is None:
            raise TrackerError("No writes to undo")
        if entries[-1]["tool"] == "undo_last_write" and entries[-1].get("undid") == entry["at"]:
            raise TrackerError("The last write is already undone")
        current = self.version()
        if current != entry["version_after"] or (expected_version and expected_version != current):
            raise ConflictError("The file changed after that write; undo would lose those changes. "
                                f"Restore by hand from {entry['backup']} if needed.")
        if dry_run:
            return {"dry_run": True, "would_undo": entry, "version": current}
        self._refuse_if_open()
        backup = self.backup_dir / entry["backup"]
        safety = self._backup()
        tmp = self.path.with_name(f".{self.path.stem}.tracker-mcp-tmp.xlsx")
        shutil.copy2(backup, tmp)
        os.replace(tmp, self.path)
        after = self.version()
        self._audit({"tool": "undo_last_write", "undid": entry["at"], "row": entry["row"],
                     "changes": [], "backup": safety.name,
                     "version_before": current, "version_after": after})
        return {"undone": entry, "backup_of_state_before_undo": str(safety), "version": after}

    # ---- writing: the safe path

    def _write(self, tool: str, mutate: Callable[[Workbook, Worksheet, Layout], tuple],
               expected_version: str | None, dry_run: bool) -> dict[str, Any]:
        if not dry_run:
            self._refuse_if_open()
        wb, ws, lay, before = self._load()
        if expected_version and expected_version != before:
            raise ConflictError(f"The tracker changed since you read it (now {before}, "
                                f"you had {expected_version}). Re-read and retry.")
        row, changes, layout_notes = mutate(wb, ws, lay)
        result: dict[str, Any] = {"row": row, "dry_run": dry_run,
                                  "changes": [c.as_dict(self._show) for c in changes],
                                  "labels": {r: self._label(ws, lay, r) for r in sorted({c.row for c in changes})}}
        if layout_notes:
            result["layout"] = layout_notes
        if dry_run or not changes:
            result["version"] = before
            if not changes:
                result["note"] = "Nothing to change"
            return result

        wb.calculation.fullCalcOnLoad = True  # Excel recomputes formulas on open
        tmp = self.path.with_name(f".{self.path.stem}.tracker-mcp-tmp.xlsx")
        try:
            wb.save(tmp)
            self._verify(tmp, changes)
            if self.version() != before:  # someone saved while we worked
                raise ConflictError("The tracker changed while writing; nothing was saved. Retry.")
            backup = self._backup()
            shutil.copymode(self.path, tmp)
            os.replace(tmp, self.path)
        finally:
            tmp.unlink(missing_ok=True)
        after = self.version()
        self._audit({"tool": tool, "row": row, "changes": result["changes"],
                     "layout": layout_notes, "backup": backup.name,
                     "version_before": before, "version_after": after})
        result.update(backup=str(backup), version=after)
        return result

    def _refuse_if_open(self) -> None:
        for lock in (f"~${self.path.name}", f".~lock.{self.path.name}#"):
            if (self.path.parent / lock).exists():
                raise ConflictError(f"The tracker is open in Excel/LibreOffice (lock file {lock}). "
                                    "Close it and retry; if it is closed, the lock file is stale.")

    def _verify(self, saved: Path, changes: list[Change]) -> None:
        ws = openpyxl.load_workbook(saved)[self.cfg.sheet]
        lay = self._layout(ws)
        for ch in changes:
            col = lay.cols[ch.field] if ch.field in lay.cols else self._column_of(lay, ch.field)
            got = ws.cell(ch.row, col).value
            ok = (parse_date(got) == ch.new) if isinstance(ch.new, date) else (got == ch.new)
            if not ok:
                raise TrackerError(f"Verification failed for row {ch.row} {ch.field}: "
                                   f"wrote {ch.new!r}, read back {got!r}. Nothing was saved.")

    def _backup(self) -> Path:
        self.backup_dir.mkdir(parents=True, exist_ok=True)
        stamp = datetime.now().strftime("%Y%m%d-%H%M%S-%f")
        dst = self.backup_dir / f"{self.path.stem}.{stamp}{self.path.suffix}"
        shutil.copy2(self.path, dst)
        old = sorted(self.backup_dir.glob(f"{self.path.stem}.*{self.path.suffix}"))
        for p in old[:-KEEP_BACKUPS]:
            p.unlink()
        return dst

    def _audit(self, entry: dict[str, Any]) -> None:
        self.backup_dir.mkdir(parents=True, exist_ok=True)
        entry = {"at": datetime.now().isoformat(timespec="seconds"), **entry}
        with open(self.backup_dir / "audit.jsonl", "a", encoding="utf-8") as f:
            f.write(json.dumps(entry, ensure_ascii=False, default=str) + "\n")

    def _audit_entries(self) -> list[dict[str, Any]]:
        log = self.backup_dir / "audit.jsonl"
        if not log.exists():
            return []
        return [json.loads(line) for line in log.read_text(encoding="utf-8").splitlines() if line.strip()]

    # ---- writing: helpers

    def _log(self, ws: Worksheet, lay: Layout, row: int, text: str, changes: list[Change]) -> None:
        if not self.cfg.has("followup_log"):
            return
        old = ws.cell(row, lay.cols["followup_log"]).value
        self._set(ws, lay, row, "followup_log", self._append(old, self._stamp(text)), changes)

    def _set(self, ws: Worksheet, lay: Layout, row: int, key: str, new: Any, changes: list[Change]) -> None:
        self._set_cell(ws.cell(row, lay.cols[key]), row, key, new, changes)

    def _set_col(self, ws: Worksheet, lay: Layout, row: int, col: int, new: Any, changes: list[Change]) -> None:
        key = next((k for k, c in lay.cols.items() if c == col), lay.headers[col])
        self._set_cell(ws.cell(row, col), row, key, new, changes)

    def _set_cell(self, cell, row: int, key: str, new: Any, changes: list[Change]) -> None:
        old = cell.value
        if isinstance(new, date):
            if parse_date(old) == new and isinstance(old, datetime):
                return
            cell.value = datetime(new.year, new.month, new.day)
            cell.number_format = "yyyy-mm-dd"
        else:
            if old == new:
                return
            cell.value = new
        changes.append(Change(row, key, self._show(old, key), new))

    @staticmethod
    def _column_of(lay: Layout, header: str) -> int:
        return next(c for c, h in lay.headers.items() if h == header)

    def _resolve_extra(self, lay: Layout, extra: dict[str, str]) -> dict[int, str]:
        """add_application's `extra`: keys are config field names or header texts."""
        out = {}
        reserved = {"num", "company", "role", "status", "applied"}
        for key, val in extra.items():
            if key in reserved:
                raise TrackerError(f"«{key}» is set by the server, not through extra")
            if key in lay.cols:
                col = lay.cols[key]
            else:
                col = next((c for c, h in lay.headers.items() if h == key), None)
                if col is None:
                    raise TrackerError(f"Unknown column «{key}». Use a field from describe_tracker "
                                       "or an exact header text.")
            out[col] = val
        return out

    def _check_row(self, ws: Worksheet, lay: Layout, row: int) -> None:
        if not (lay.header_row < row <= lay.last_row) or ws.cell(row, lay.cols["company"]).value in (None, ""):
            raise TrackerError(f"Row {row} is not an application row "
                               f"(data rows are {lay.header_row + 1}–{lay.last_row}).")

    def _normalize_status(self, ws: Worksheet, lay: Layout, value: str) -> str:
        allowed = self._allowed_statuses(ws, lay)
        v = value.strip()
        by_fold = {s.casefold(): s for s in allowed}
        if v.casefold() in by_fold:
            return by_fold[v.casefold()]
        alias = self.aliases.get(v.casefold())
        if alias in allowed:
            return alias
        raise TrackerError(f"Unknown status «{value}». Allowed (the dropdown in column "
                           f"{get_column_letter(lay.cols['status'])}): {', '.join(allowed)}")

    def _allowed_statuses(self, ws: Worksheet, lay: Layout) -> list[str]:
        col = lay.cols["status"]
        for dv in ws.data_validations.dataValidation:
            if dv.type == "list" and dv.formula1 and any(
                    rng.min_col <= col <= rng.max_col for rng in dv.sqref.ranges):
                return [s.strip() for s in dv.formula1.strip('"').split(",") if s.strip()]
        return list(self.cfg.statuses.fallback)

    @staticmethod
    def _copy_row_style(ws: Worksheet, lay: Layout, src: int, dst: int) -> None:
        for c in lay.headers:
            a, b = ws.cell(src, c), ws.cell(dst, c)
            if a.has_style:
                b.font, b.fill, b.border = copy(a.font), copy(a.fill), copy(a.border)
                b.alignment, b.protection = copy(a.alignment), copy(a.protection)
                b.number_format = a.number_format
        if ws.row_dimensions[src].height:
            ws.row_dimensions[dst].height = ws.row_dimensions[src].height

    def _grow_layout(self, wb: Workbook, ws: Worksheet, lay: Layout, new_row: int) -> list[str]:
        """Stretch dropdowns, the filter and the summary sheet's formulas to cover new_row."""
        notes = []
        for dv in ws.data_validations.dataValidation:
            ranges = list(dv.sqref.ranges)
            grown = False
            for rng in ranges:
                if rng.min_row <= lay.header_row + 1 and rng.max_row < new_row:
                    rng.expand(down=new_row - rng.max_row)
                    grown = True
            if grown:
                dv.sqref = MultiCellRange(" ".join(r.coord for r in ranges))
                notes.append(f"dropdown {dv.sqref} now reaches row {new_row}")
        if ws.auto_filter.ref:
            m = re.fullmatch(r"([A-Z]+)(\d+):([A-Z]+)(\d+)", ws.auto_filter.ref)
            if m and int(m.group(4)) < new_row:
                ws.auto_filter.ref = f"{m.group(1)}{m.group(2)}:{m.group(3)}{new_row}"
                notes.append(f"filter now {ws.auto_filter.ref}")
        summary = self.cfg.summary_sheet
        if summary and summary in wb.sheetnames:
            sheet = self.cfg.sheet
            ref = re.escape(sheet) if re.fullmatch(r"\w+", sheet) else re.escape(f"'{sheet}'")
            pattern = re.compile(rf"({ref}!\$?[A-Z]+\$?\d+:\$?[A-Z]+\$?)(\d+)")
            grown_cells = []
            for row in wb[summary].iter_rows():
                for cell in row:
                    f = cell.value
                    if isinstance(f, str) and f.startswith("="):
                        new_f = pattern.sub(lambda m: m.group(1) + str(max(int(m.group(2)), new_row + 100))
                                            if int(m.group(2)) < new_row else m.group(0), f)
                        if new_f != f:
                            cell.value = new_f
                            grown_cells.append(cell.coordinate)
            if grown_cells:
                notes.append(f"«{summary}» formulas extended to row {new_row + 100}: "
                             f"{', '.join(grown_cells)}")
        return notes
