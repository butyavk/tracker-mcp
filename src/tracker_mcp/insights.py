"""Reply times and follow-up calendars: what the tracker's dates say over time.

reply_times needs the day each reply came. A `replied_on` column has it (set_status
fills it on the first reply); without one, the days this server recorded a reply
stand in, which covers only replies recorded through it.
"""

from __future__ import annotations

import hashlib
import statistics
from datetime import UTC, date, datetime, timedelta
from pathlib import Path
from typing import TYPE_CHECKING, Any

from .doctor import reply_dates_from_audit
from .tracker import TrackerError

if TYPE_CHECKING:
    from .tracker import Tracker

SILENCE_STEPS = (7, 14, 21, 30, 45, 60)


def _stats(days: list[int]) -> dict[str, Any]:
    if not days:
        return {"replies_with_dates": 0}
    q = statistics.quantiles(days, n=4, method="inclusive") if len(days) > 1 else [days[0]] * 3
    return {"replies_with_dates": len(days), "median_days": statistics.median(days),
            "p75_days": round(q[2], 1), "max_days": max(days)}


def reply_times(t: Tracker, by: str | None = None, since: date | None = None) -> dict[str, Any]:
    t._need("applied_on")
    recs, version = t.records(full=bool(by and not t.cfg.has(by) and by != "month"))
    if by and by != "month" and not t.cfg.has(by):
        for r in recs:
            r[by] = r.get("other", {}).get(by)
    audit = {} if t.cfg.has("replied_on") else reply_dates_from_audit(t)
    today = t.today()
    rows: list[tuple[dict[str, Any], date, date | None, bool]] = []   # rec, applied, replied, any reply
    unknown = 0
    for rec in recs:
        applied = t._date(rec.get("applied_on"))
        if not t._is_applied(rec) or applied is None or (since and applied < since):
            continue
        replied = t._date(rec.get("replied_on")) if t.cfg.has("replied_on") else audit.get(rec["row"])
        any_reply = t._furthest(rec) >= 0
        if any_reply and replied is None:
            unknown += 1
            continue
        if replied and replied < applied:   # a typo in a date: leave the row out
            unknown += 1
            continue
        rows.append((rec, applied, replied if any_reply else None, any_reply))

    def block(items: list[tuple[dict[str, Any], date, date | None, bool]]) -> dict[str, Any]:
        days = [(rep - app).days for _, app, rep, _ in items if rep]
        return {"applied": len(items), **_stats(days)}

    silence = []
    for d in SILENCE_STEPS:
        # applications that had been silent for d days: replied after day d, or never
        # replied and submitted at least d days ago
        still = [x for x in rows if (x[2] and (x[2] - x[1]).days > d) or (not x[3] and (today - x[1]).days >= d)]
        later = sum(1 for x in still if x[2])
        if still:
            silence.append({"after_days_silent": d, "applications": len(still), "replied_later": later,
                            "share": round(100 * later / len(still), 1)})
    out: dict[str, Any] = {"total": block(rows), "silence": silence}
    if by:
        groups: dict[str, list] = {}
        for x in rows:
            groups.setdefault(t._group(x[0], by), []).append(x)
        out["by"] = by
        out["groups"] = [{"group": k, **block(v)} for k, v in
                         sorted(groups.items(), key=lambda kv: -len(kv[1]))][:12]
    out["replies_without_a_date"] = unknown
    out["source"] = ("the replied_on column" if t.cfg.has("replied_on") else
                     "the days this server recorded a reply (add a reply-date column for all replies; "
                     "check_tracker can fill it from the response column)")
    out["how_to_read"] = ("silence: of the applications that had heard nothing for N days, how many still "
                          "got a reply later. Where that share is near zero, waiting longer rarely helps.")
    out["version"] = version
    return out


# ---------------------------------------------------------------- calendar

def followup_events(t: Tracker, days: int = 60) -> tuple[list[dict[str, Any]], str]:
    """Open applications with a follow-up date up to `days` ahead. An overdue one is put
    on today, so the calendar shows it where it is still actionable."""
    t._need("followup_on")
    recs, version = t.records()
    today = t.today()
    events = []
    for rec in recs:
        if rec["status"] in t.closed or rec["status"] == t.cfg.statuses.not_applied:
            continue
        due = t._date(rec.get("followup_on"))
        if due is None or due > today + timedelta(days=days):
            continue
        overdue = due < today
        title = f"Follow-up: {rec['company']}" + (f" — {rec['role']}" if rec.get("role") else "")
        lines = [f"Due {due.isoformat()}" + (" (overdue)" if overdue else ""),
                 f"Status: {rec.get('status')}"]
        if rec.get("followup_contact"):
            lines.append(f"Contact: {rec['followup_contact']}")
        if rec.get("link"):
            lines.append(str(rec["link"]))
        if rec.get("followup_log"):
            lines.append(f"Log: …{str(rec['followup_log'])[-200:]}")
        key = f"{t.store.label}|{rec['company']}|{rec.get('role')}"
        events.append({"uid": hashlib.sha256(key.encode()).hexdigest()[:24] + "@tracker-mcp",
                       "date": (today if overdue else due).isoformat(), "due": due.isoformat(),
                       "overdue": overdue, "row": rec["row"], "title": ("Overdue " if overdue else "") + title,
                       "description": "\n".join(lines)})
    events.sort(key=lambda e: (e["date"], e["title"]))
    return events, version


def _ics_text(s: str) -> str:
    return s.replace("\\", "\\\\").replace(";", "\\;").replace(",", "\\,").replace("\n", "\\n")


def _fold(line: str) -> list[str]:
    """RFC 5545: lines of at most 75 octets, continued with a leading space."""
    out, cur = [], ""
    for ch in line:
        if len((cur + ch).encode()) > (75 if not out else 74):
            out.append(cur)
            cur = ""
        cur += ch
    out.append(cur)
    return [out[0], *(" " + x for x in out[1:])]


def to_ics(events: list[dict[str, Any]], name: str = "Job follow-ups") -> str:
    stamp = datetime.now(UTC).strftime("%Y%m%dT%H%M%SZ")
    lines = ["BEGIN:VCALENDAR", "VERSION:2.0", "PRODID:-//tracker-mcp//follow-ups//EN", "CALSCALE:GREGORIAN",
             "METHOD:PUBLISH", f"X-WR-CALNAME:{_ics_text(name)}"]
    for e in events:
        start = date.fromisoformat(e["date"])
        lines += ["BEGIN:VEVENT", f"UID:{e['uid']}", f"DTSTAMP:{stamp}",
                  f"DTSTART;VALUE=DATE:{start:%Y%m%d}", f"DTEND;VALUE=DATE:{start + timedelta(days=1):%Y%m%d}",
                  f"SUMMARY:{_ics_text(e['title'])}", f"DESCRIPTION:{_ics_text(e['description'])}",
                  "TRANSP:TRANSPARENT", "BEGIN:VALARM", "ACTION:DISPLAY", f"DESCRIPTION:{_ics_text(e['title'])}",
                  "TRIGGER;RELATED=START:PT9H", "END:VALARM", "END:VEVENT"]
    lines.append("END:VCALENDAR")
    return "\r\n".join(part for line in lines for part in _fold(line)) + "\r\n"


def default_calendar_path(t: Tracker) -> Path:
    if t.path is not None:
        return t.path.with_name(f"{t.path.stem}.followups.ics")
    return t.backup_dir.parent / f"{t.backup_dir.name}.followups.ics"


def export_calendar(t: Tracker, path: str | Path | None = None, days: int = 60) -> dict[str, Any]:
    events, version = followup_events(t, days)
    target = Path(path).expanduser() if path else default_calendar_path(t)
    if target.suffix.lower() != ".ics":
        raise TrackerError("The calendar file must end in .ics")
    target.parent.mkdir(parents=True, exist_ok=True)
    tmp = target.with_name(target.name + ".tmp")
    tmp.write_text(to_ics(events), encoding="utf-8", newline="")
    tmp.replace(target)
    return {"file": str(target), "events": len(events),
            "overdue_on_today": sum(e["overdue"] for e in events),
            "first": [{k: e[k] for k in ("date", "title")} for e in events[:10]],
            "how_to_use": ("Open the file to import it into Calendar, Outlook or Google Calendar (Settings → "
                           "Import). Each event keeps the same ID between exports; import into a calendar of "
                           "its own, so an old copy is easy to clear before importing a fresh one."),
            "version": version}
