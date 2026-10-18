"""MCP server over a job-application tracker in an .xlsx file (stdio transport).

Thin layer: each tool validates its arguments (pydantic, from the type hints)
and calls one method of `Tracker`. File logic lives in tracker.py, everything
specific to one workbook lives in its config (config.py).

    TRACKER_PATH=~/tracker.xlsx tracker-mcp           # serve over stdio
    TRACKER_PATH=~/tracker.xlsx tracker-mcp --check   # print how the config maps the file
"""

from __future__ import annotations

import argparse
import json
import os
import sys
from datetime import date
from typing import Annotated, Any

from mcp.server import MCPServer
from mcp.server.mcpserver.exceptions import ToolError
from mcp.types import ToolAnnotations
from pydantic import Field

from .config import ConfigError, load_config
from .tracker import Tracker, TrackerError

INSTRUCTIONS = """\
A job-application tracker kept in an Excel file.
- Call describe_tracker first: it lists the columns, the allowed statuses and the
  follow-up rule of this particular workbook. Column names and statuses may be in
  any language; use them exactly as the file has them.
- Rows are identified by their sheet row number (`row`), as in Excel.
- Before any write, call the write tool with dry_run=true, show the user the
  changes, and write only after they agree. Pass `version` from your last read as
  expected_version so a concurrent edit is detected instead of overwritten.
- Every write makes a backup; undo_last_write restores it if nothing changed since.
"""

mcp = MCPServer("tracker", instructions=INSTRUCTIONS)

READ = ToolAnnotations(read_only_hint=True, open_world_hint=False)
ADD = ToolAnnotations(read_only_hint=False, destructive_hint=False, idempotent_hint=False,
                      open_world_hint=False)
EDIT = ToolAnnotations(read_only_hint=False, destructive_hint=True, idempotent_hint=False,
                       open_world_hint=False)

Row = Annotated[int, Field(description="Sheet row number of the application (see list/get results)")]
Version = Annotated[str | None, Field(description="`version` from your last read; the write is "
                                                  "refused if the file changed since")]
DryRun = Annotated[bool, Field(description="Only show what would change; write nothing")]

_tracker: Tracker | None = None


def make_tracker() -> Tracker:
    path = os.environ["TRACKER_PATH"]
    return Tracker(path, config=load_config(path, os.environ.get("TRACKER_CONFIG")),
                   backup_dir=os.environ.get("TRACKER_BACKUP_DIR"))


def tracker() -> Tracker:
    global _tracker
    if _tracker is None:
        _tracker = make_tracker()
    return _tracker


def call(method: str, *args: Any, **kwargs: Any) -> dict[str, Any]:
    """Run a Tracker method. Expected refusals (conflict, bad status, duplicate)
    go back to the model as a readable tool error; anything else is a crash and
    the SDK hides its text."""
    try:
        return getattr(tracker(), method)(*args, **kwargs)
    except TrackerError as e:
        raise ToolError(str(e)) from e


# ---------------------------------------------------------------- read

@mcp.tool(annotations=READ)
def describe_tracker() -> dict[str, Any]:
    """How this workbook is laid out: field -> column, custom fields, allowed
    statuses, which statuses are closed, and the follow-up rule."""
    return call("describe")


@mcp.tool(annotations=READ)
def list_applications(
    status: Annotated[str | None, Field(description="Filter by status; comma-separated for several")] = None,
    query: Annotated[str | None, Field(description="Substring of company or role, case-insensitive")] = None,
    applied_since: Annotated[date | None, Field(description="Only applications submitted on/after")] = None,
    applied_until: Annotated[date | None, Field(description="Only applications submitted on/before")] = None,
    open_only: Annotated[bool, Field(description="Hide closed statuses (rejected, withdrawn, ...)")] = False,
    limit: int = 100,
) -> dict[str, Any]:
    """List applications in brief: row, company, role, status, submit and follow-up dates."""
    return call("list_applications", status, query, applied_since, applied_until, open_only, limit)


@mcp.tool(annotations=READ)
def get_application(
    query: Annotated[str | None, Field(description="Company or role text, e.g. 'Acme'")] = None,
    row: Annotated[int | None, Field(description="Exact sheet row number")] = None,
) -> dict[str, Any]:
    """Everything recorded for an application: all columns, notes, follow-up log."""
    return call("get_application", query, row)


@mcp.tool(annotations=READ)
def followups_due(
    within_days: Annotated[int, Field(ge=0, le=60, description="Look this many days ahead")] = 7,
) -> dict[str, Any]:
    """Follow-ups due by today + within_days (overdue included), plus submitted
    applications that have no follow-up date at all."""
    return call("followups_due", within_days)


@mcp.tool(annotations=READ)
def summary(
    since: Annotated[date | None, Field(description="Period start for the submitted cohort")] = None,
    until: Annotated[date | None, Field(description="Period end for the submitted cohort")] = None,
) -> dict[str, Any]:
    """Counts by status and by month. With since/until: how many were submitted in that
    period and where they stand now."""
    return call("summary", since, until)


@mcp.tool(annotations=READ)
def recent_writes(limit: int = 10) -> dict[str, Any]:
    """The audit log of this server's writes, newest first."""
    return call("recent_writes", limit)


# ---------------------------------------------------------------- write

@mcp.tool(annotations=ADD)
def add_application(
    company: str,
    role: str,
    link: Annotated[str | None, Field(description="Application URL")] = None,
    notes: str | None = None,
    extra: Annotated[dict[str, str] | None, Field(
        description="Other columns: a field name from describe_tracker or an exact header text -> value")] = None,
    allow_duplicate: bool = False,
    expected_version: Version = None,
    dry_run: DryRun = False,
) -> dict[str, Any]:
    """Add a new row with the not-applied status. Refuses a duplicate (same company+role or link)."""
    return call("add_application", company, role, link=link, notes=notes, extra=extra,
                allow_duplicate=allow_duplicate, expected_version=expected_version, dry_run=dry_run)


@mcp.tool(annotations=EDIT)
def mark_applied(
    row: Row,
    applied_on: Annotated[date | None, Field(description="Submit date; default today")] = None,
    channel: Annotated[str | None, Field(description="How it was sent, e.g. 'Ashby (direct form)'")] = None,
    followup_on: Annotated[date | None, Field(
        description="Default: from the follow-up rule (describe_tracker)")] = None,
    note: str | None = None,
    expected_version: Version = None,
    dry_run: DryRun = False,
) -> dict[str, Any]:
    """Record a submission: applied flag, date, channel, the applied status, follow-up 1 date."""
    return call("mark_applied", row, applied_on=applied_on, channel=channel, followup_on=followup_on,
                note=note, expected_version=expected_version, dry_run=dry_run)


@mcp.tool(annotations=EDIT)
def set_status(
    row: Row,
    status: Annotated[str, Field(description="A value from the status dropdown")],
    response: Annotated[str | None, Field(
        description="What the company replied; appended to the response column")] = None,
    note: Annotated[str | None, Field(description="Appended to the notes column")] = None,
    expected_version: Version = None,
    dry_run: DryRun = False,
) -> dict[str, Any]:
    """Change the status. A closed status (see describe_tracker) also cancels a pending
    follow-up date."""
    return call("set_status", row, status, response=response, note=note,
                expected_version=expected_version, dry_run=dry_run)


@mcp.tool(annotations=EDIT)
def log_followup(
    row: Row,
    text: Annotated[str, Field(description="What was sent and to whom; appended to the follow-up log")],
    next_on: Annotated[date | None, Field(description="Next follow-up date; default from the follow-up rule")] = None,
    final: Annotated[bool, Field(description="No more follow-ups: clears the follow-up date")] = False,
    contact: Annotated[str | None, Field(description="Replace the follow-up contact / channel")] = None,
    expected_version: Version = None,
    dry_run: DryRun = False,
) -> dict[str, Any]:
    """Record a sent follow-up and schedule the next one."""
    return call("log_followup", row, text, next_on=next_on, final=final, contact=contact,
                expected_version=expected_version, dry_run=dry_run)


@mcp.tool(annotations=EDIT)
def undo_last_write(expected_version: Version = None) -> dict[str, Any]:
    """Restore the file to its state before this server's last write. Refused if the
    file has changed since that write (it would lose someone else's edits)."""
    return call("undo_last_write", expected_version)


def main() -> None:
    parser = argparse.ArgumentParser(prog="tracker-mcp")
    parser.add_argument("--check", action="store_true",
                        help="print how the config maps the workbook, then exit")
    args = parser.parse_args()
    path = os.environ.get("TRACKER_PATH")
    if not path or not os.path.exists(os.path.expanduser(path)):
        print(f"tracker-mcp: set TRACKER_PATH to the .xlsx tracker (got {path!r})", file=sys.stderr)
        sys.exit(2)
    global _tracker
    try:  # fail at startup, with the reason in the client's log, not on the first call
        _tracker = make_tracker()
        layout = _tracker.describe()
    except (ConfigError, TrackerError) as e:
        print(f"tracker-mcp: {e}", file=sys.stderr)
        sys.exit(2)
    if args.check:
        print(json.dumps(layout, ensure_ascii=False, indent=2))
        return
    mcp.run()  # stdio


if __name__ == "__main__":
    main()
