"""MCP server over a job-application tracker in an .xlsx file (stdio transport).

Thin layer: each tool validates its arguments (pydantic, from the type hints)
and calls one method of `Tracker`. File logic lives in tracker.py, everything
specific to one workbook lives in its config (config.py). The one policy that
lives here is confirmation: no write happens without the user's yes (`guarded`).

    TRACKER_PATH=~/tracker.xlsx tracker-mcp           # serve over stdio
    TRACKER_PATH=~/tracker.xlsx tracker-mcp --check   # print how the config maps the file
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import sys
from collections import OrderedDict
from datetime import date
from typing import Annotated, Any

from mcp.server import MCPServer
from mcp.server.elicitation import AcceptedElicitation, ElicitationResult
from mcp.server.mcpserver import Context, Elicit, Resolve
from mcp.server.mcpserver.exceptions import ToolError
from mcp.types import ToolAnnotations
from pydantic import BaseModel, Field

from .config import ConfigError, load_config
from .tracker import Tracker, TrackerError

INSTRUCTIONS = """\
A job-application tracker kept in an Excel file.
- Call describe_tracker first: it lists the columns, the allowed statuses and the
  follow-up rule of this particular workbook. Column names and statuses may be in
  any language; use them exactly as the file has them.
- Rows are identified by their sheet row number (`row`), as in Excel.
- Every write needs the user's yes, and the server enforces it. If your client can
  show a confirmation dialog, the server asks the user there. Otherwise call the
  write tool with dry_run=true, show the user the changes, and after they agree call
  it again with the same arguments plus confirm=<token from the preview>.
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
DryRun = Annotated[bool, Field(description="Only show what would change, plus the confirm token; write nothing")]
Confirm = Annotated[str | None, Field(description="Token from the dry run the user approved. Not needed "
                                                  "when the client shows the server's confirmation dialog")]

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


# ---------------------------------------------------------------- confirmation
#
# No write happens without the user's yes, and the server enforces it rather than
# trusting the model to ask. Each write tool has a resolver that runs before the
# tool body: it computes the plan (a dry run) and, when the client can show a
# dialog, returns an `Elicit` request, so the SDK asks the user with the plan as
# the message (mid-call on older protocol versions, as an input-required round on
# 2026-07-28+). A client without a dialog must echo the token of a dry run instead.
# Either way the body writes with the plan's version as If-Match, so what gets
# written is exactly what the user approved, or a conflict.

class WriteConfirmation(BaseModel):
    confirm: bool = Field(description="Apply these changes to the tracker")


SKIP, BY_TOKEN = "skip", "token"   # resolver outcomes when no dialog is shown
_plans: OrderedDict[str, dict[str, Any]] = OrderedDict()   # request id -> plan, resolver -> body


def plan_token(plan: dict[str, Any]) -> str:
    """Fingerprint of exactly what will be written, at exactly this file version."""
    basis = {k: plan.get(k) for k in ("row", "changes", "version", "would_undo")}
    data = json.dumps(basis, sort_keys=True, ensure_ascii=False, default=str)
    return hashlib.sha256(data.encode()).hexdigest()[:12]


def _has_work(plan: dict[str, Any]) -> bool:
    return bool(plan.get("changes") or plan.get("would_undo"))


def _short(value: Any, limit: int = 120) -> str:
    text = "∅" if value in (None, "") else str(value)
    return text if len(text) <= limit else text[: limit - 1] + "…"


def describe_plan(tool: str, plan: dict[str, Any]) -> str:
    """The text the user sees in the confirmation dialog."""
    if plan.get("would_undo"):
        e = plan["would_undo"]
        text = f"Undo {e['tool']} from {e['at']} (rows {e['row']}): restore the backup {e['backup']}?"
    else:
        labels = plan.get("labels", {})
        lines = [f"{tool}: change {len(labels)} row(s) of the tracker?"]
        for row, label in labels.items():
            lines.append(f"\nRow {row} · {label}")
            for ch in plan["changes"]:
                if str(ch["row"]) != str(row):
                    continue
                old, new = ch["old"], ch["new"]
                if isinstance(old, str) and isinstance(new, str) and old and new.startswith(old):
                    lines.append(f"  {ch['field']}: + {_short(new[len(old):].lstrip(' |'))}")
                else:
                    lines.append(f"  {ch['field']}: {_short(old)} → {_short(new)}")
        text = "\n".join(lines)
    return f"{text}\n\n(plan {plan_token(plan)})"


def ask(ctx: Context, method: str, args: tuple, kwargs: dict[str, Any], *,
        expected_version: str | None, dry_run: bool) -> Elicit[WriteConfirmation] | str:
    """Resolver body: make the plan and decide how the user approves it."""
    plan = call(method, *args, dry_run=True, expected_version=expected_version, **kwargs)
    _plans[ctx.request_id] = plan
    while len(_plans) > 50:
        _plans.popitem(last=False)
    if dry_run or not _has_work(plan) or not tracker().cfg.safety.confirm_writes:
        return SKIP
    caps = ctx.client_capabilities
    if caps is not None and caps.elicitation is not None:
        return Elicit(describe_plan(method, plan), WriteConfirmation)
    return BY_TOKEN


def finish(ctx: Context, method: str, args: tuple, kwargs: dict[str, Any], approval: Any, *,
           dry_run: bool, confirm: str | None) -> dict[str, Any]:
    """Tool body: return the preview, or write the approved plan."""
    plan = _plans.pop(ctx.request_id, None)
    if plan is None:  # not expected: the resolver runs first in the same request
        raise ToolError("Internal: no plan for this call; retry.")
    if not _has_work(plan):
        return plan
    token = plan_token(plan)
    if dry_run:
        return {**plan, "confirm": token}
    if not isinstance(approval, AcceptedElicitation):
        raise ToolError("The user did not confirm; nothing was written.")
    answer = approval.data
    if isinstance(answer, WriteConfirmation) and not answer.confirm:
        raise ToolError("The user did not confirm; nothing was written.")
    if answer == BY_TOKEN and confirm != token:
        raise ToolError("Not written: this needs the user's confirmation. Call the tool with dry_run=true, "
                        "show the user the changes, and after they agree repeat the call with "
                        "confirm=<token from that preview>. A token from another preview, or one made "
                        "before the file changed, does not match.")
    return call(method, *args, dry_run=False, expected_version=plan["version"], **kwargs)


Approval = ElicitationResult[WriteConfirmation]


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
def stale_applications(
    days: Annotated[int, Field(ge=1, description="Submitted at least this many days ago")] = 30,
) -> dict[str, Any]:
    """Applications still waiting for a reply: a waiting status, submitted `days`+ days
    ago, and no follow-up planned for today or later. These are what close_stale closes."""
    return call("stale_applications", days)


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
#
# Each write tool = a resolver (`_ask_*`: plan + how to confirm) and a body
# (`finish`: preview or write). The resolver receives the tool's own arguments.

def _ask_add(ctx: Context, company: str, role: str, link: str | None = None, notes: str | None = None,
             extra: dict[str, str] | None = None, allow_duplicate: bool = False,
             expected_version: str | None = None, dry_run: bool = False):
    return ask(ctx, "add_application", (company, role),
               dict(link=link, notes=notes, extra=extra, allow_duplicate=allow_duplicate),
               expected_version=expected_version, dry_run=dry_run)


@mcp.tool(annotations=ADD)
def add_application(
    ctx: Context,
    approval: Annotated[Approval, Resolve(_ask_add)],
    company: str,
    role: str,
    link: Annotated[str | None, Field(description="Application URL")] = None,
    notes: str | None = None,
    extra: Annotated[dict[str, str] | None, Field(
        description="Other columns: a field name from describe_tracker or an exact header text -> value")] = None,
    allow_duplicate: bool = False,
    expected_version: Version = None,
    dry_run: DryRun = False,
    confirm: Confirm = None,
) -> dict[str, Any]:
    """Add a new row with the not-applied status. Refuses a duplicate (same company+role or link)."""
    return finish(ctx, "add_application", (company, role),
                  dict(link=link, notes=notes, extra=extra, allow_duplicate=allow_duplicate),
                  approval, dry_run=dry_run, confirm=confirm)


def _ask_applied(ctx: Context, row: int, applied_on: date | None = None, channel: str | None = None,
                 followup_on: date | None = None, note: str | None = None,
                 expected_version: str | None = None, dry_run: bool = False):
    return ask(ctx, "mark_applied", (row,),
               dict(applied_on=applied_on, channel=channel, followup_on=followup_on, note=note),
               expected_version=expected_version, dry_run=dry_run)


@mcp.tool(annotations=EDIT)
def mark_applied(
    ctx: Context,
    approval: Annotated[Approval, Resolve(_ask_applied)],
    row: Row,
    applied_on: Annotated[date | None, Field(description="Submit date; default today")] = None,
    channel: Annotated[str | None, Field(description="How it was sent, e.g. 'Ashby (direct form)'")] = None,
    followup_on: Annotated[date | None, Field(
        description="Default: from the follow-up rule (describe_tracker)")] = None,
    note: str | None = None,
    expected_version: Version = None,
    dry_run: DryRun = False,
    confirm: Confirm = None,
) -> dict[str, Any]:
    """Record a submission: applied flag, date, channel, the applied status, follow-up 1 date."""
    return finish(ctx, "mark_applied", (row,),
                  dict(applied_on=applied_on, channel=channel, followup_on=followup_on, note=note),
                  approval, dry_run=dry_run, confirm=confirm)


def _ask_status(ctx: Context, row: int, status: str, response: str | None = None, note: str | None = None,
                expected_version: str | None = None, dry_run: bool = False):
    return ask(ctx, "set_status", (row, status), dict(response=response, note=note),
               expected_version=expected_version, dry_run=dry_run)


@mcp.tool(annotations=EDIT)
def set_status(
    ctx: Context,
    approval: Annotated[Approval, Resolve(_ask_status)],
    row: Row,
    status: Annotated[str, Field(description="A value from the status dropdown")],
    response: Annotated[str | None, Field(
        description="What the company replied; appended to the response column")] = None,
    note: Annotated[str | None, Field(description="Appended to the notes column")] = None,
    expected_version: Version = None,
    dry_run: DryRun = False,
    confirm: Confirm = None,
) -> dict[str, Any]:
    """Change the status. A closed status (see describe_tracker) also cancels a pending
    follow-up date."""
    return finish(ctx, "set_status", (row, status), dict(response=response, note=note),
                  approval, dry_run=dry_run, confirm=confirm)


def _ask_followup(ctx: Context, row: int, text: str, next_on: date | None = None, final: bool = False,
                  contact: str | None = None, expected_version: str | None = None, dry_run: bool = False):
    return ask(ctx, "log_followup", (row, text), dict(next_on=next_on, final=final, contact=contact),
               expected_version=expected_version, dry_run=dry_run)


@mcp.tool(annotations=EDIT)
def log_followup(
    ctx: Context,
    approval: Annotated[Approval, Resolve(_ask_followup)],
    row: Row,
    text: Annotated[str, Field(description="What was sent and to whom; appended to the follow-up log")],
    next_on: Annotated[date | None, Field(description="Next follow-up date; default from the follow-up rule")] = None,
    final: Annotated[bool, Field(description="No more follow-ups: clears the follow-up date")] = False,
    contact: Annotated[str | None, Field(description="Replace the follow-up contact / channel")] = None,
    expected_version: Version = None,
    dry_run: DryRun = False,
    confirm: Confirm = None,
) -> dict[str, Any]:
    """Record a sent follow-up and schedule the next one."""
    return finish(ctx, "log_followup", (row, text), dict(next_on=next_on, final=final, contact=contact),
                  approval, dry_run=dry_run, confirm=confirm)


def _ask_stale(ctx: Context, days: int = 30, rows: list[int] | None = None,
               expected_version: str | None = None, dry_run: bool = False):
    return ask(ctx, "close_stale", (days,), dict(rows=rows),
               expected_version=expected_version, dry_run=dry_run)


@mcp.tool(annotations=EDIT)
def close_stale(
    ctx: Context,
    approval: Annotated[Approval, Resolve(_ask_stale)],
    days: Annotated[int, Field(ge=1, description="Submitted at least this many days ago")] = 30,
    rows: Annotated[list[int] | None, Field(
        description="Only these rows (each must be stale); default: all from stale_applications")] = None,
    expected_version: Version = None,
    dry_run: DryRun = False,
    confirm: Confirm = None,
) -> dict[str, Any]:
    """Close stale applications in one write: the no-response status, no follow-up date,
    a line in the follow-up log. One backup and one undo step for the whole batch."""
    return finish(ctx, "close_stale", (days,), dict(rows=rows), approval, dry_run=dry_run, confirm=confirm)


def _ask_undo(ctx: Context, expected_version: str | None = None, dry_run: bool = False):
    return ask(ctx, "undo_last_write", (), {}, expected_version=expected_version, dry_run=dry_run)


@mcp.tool(annotations=EDIT)
def undo_last_write(
    ctx: Context,
    approval: Annotated[Approval, Resolve(_ask_undo)],
    expected_version: Version = None,
    dry_run: DryRun = False,
    confirm: Confirm = None,
) -> dict[str, Any]:
    """Restore the file to its state before this server's last write. Refused if the
    file has changed since that write (it would lose someone else's edits)."""
    return finish(ctx, "undo_last_write", (), {}, approval, dry_run=dry_run, confirm=confirm)


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
