"""What the server needs to know about a particular tracker workbook.

Nothing about one person's tracker lives in the code: the sheet, the column
headers, the status words, the follow-up rule and the texts the server writes
into the follow-up log all come from a TOML file. By default the file sits next
to the workbook: `my-tracker.xlsx` -> `my-tracker.tracker-mcp.toml`.
`TRACKER_CONFIG` points to another one. `examples/tracker-mcp.example.toml`
describes every key.
"""

from __future__ import annotations

import json
import re
import tomllib
from pathlib import Path

from pydantic import BaseModel, ConfigDict, Field, ValidationError, field_validator

# Fields the server's logic understands. Any other key in [columns] is a custom
# field: it is returned under its own key and can be set through add_application.
REQUIRED_FIELDS = ("company", "role", "status")
KNOWN_FIELDS = REQUIRED_FIELDS + (
    "num", "link", "applied", "applied_on", "channel", "response", "notes",
    "followup_on", "followup_contact", "followup_log", "reached", "replied_on",
)


class _Strict(BaseModel):
    model_config = ConfigDict(extra="forbid", frozen=True)


class Statuses(_Strict):
    not_applied: str = "Not applied"
    applied: str = "Applied"
    closed: list[str] = ["Rejected", "Withdrawn", "No response", "Offer", "Closed"]
    # Still waiting for the company: close_stale only touches these.
    waiting: list[str] | None = None   # default: [applied]
    no_response: str = "No response"   # what close_stale sets
    # For the funnel. A rejection is a reply. Stages run in this order, and a status
    # in a later stage means the application got through the earlier ones.
    rejected: list[str] = ["Rejected"]
    stages: dict[str, list[str]] = {"screening": ["Screening"], "interview": ["Interview"], "offer": ["Offer"]}
    # Used only when the status column has no dropdown in the workbook.
    fallback: list[str] = ["Not applied", "Applied", "Screening", "Interview", "Offer",
                           "Rejected", "Withdrawn", "No response", "Closed"]
    # Alternative spellings the model may use -> the value from the dropdown.
    aliases: dict[str, str] = {}


class Values(_Strict):
    yes: str = "Yes"                 # «applied» column, when there is one
    no: str = "No"
    no_date: str = "—"               # written into the follow-up date when none is due
    separator: str = " | "           # between entries appended to a text cell
    stamp_format: str = "%Y-%m-%d"   # date prefix of appended entries (strftime)
    # Extra formats for dates kept as text, tried after ISO and dd.mm.yyyy,
    # e.g. ["%m/%d/%Y"] for US dates or ["%d/%m/%Y"] for UK ones. `init` detects them.
    date_formats: list[str] = []


class Followups(_Strict):
    first_after_days: int = Field(10, ge=0)   # submitted -> follow-up 1
    next_after_days: int = Field(14, ge=0)    # follow-up sent -> the next one
    weekend_to_monday: bool = True


class Messages(_Strict):
    """Texts appended to the follow-up log. {date} and {status} are filled in."""
    applied: str = "applied"
    applied_with_followup: str = "applied, follow-up 1 on {date}"
    followup_cancelled: str = "follow-up {date} cancelled: {status}"
    stale_closed: str = "closed: no reply for {days} days"
    followup_moved: str = "follow-up moved to {date}: {reason}"
    followup_dropped: str = "no follow-up planned: {reason}"


class Funnel(_Strict):
    # group name -> substrings (any case). A value of the grouped column goes to the
    # first group with a matching substring, otherwise to its own text up to the
    # first bracket or dash: "Ashby (direct form)" -> "Ashby".
    groups: dict[str, list[str]] = {}
    # The same, for one column only: [funnel.by.source] "LinkedIn" = ["linkedin"].
    # Used instead of `groups` when grouping by that column.
    by: dict[str, dict[str, list[str]]] = {}


class Rules(_Strict):
    # add_application and mark_applied warn (they do not refuse) when the same company
    # already has an application submitted this many days around the new one. 0: off.
    same_company_days: int = Field(28, ge=0)
    # A waiting application counts as silent from its last touch: the submit date or
    # the newest date that starts an entry of the follow-up log, whichever is later.
    # "applied": from the submit date only, as before 0.4.
    silence_from: str = Field("last_touch", pattern="^(last_touch|applied)$")


class Safety(_Strict):
    # Every write needs the user's yes: through the client's confirmation dialog
    # (MCP elicitation) when it has one, otherwise through the token from a dry run.
    confirm_writes: bool = True


class Config(_Strict):
    sheet: str
    summary_sheet: str | None = None   # its formulas are stretched when rows are added
    header_row: int | None = Field(None, ge=1)  # found automatically if not set
    # field -> header text. Exact match wins, otherwise the first header
    # (left to right) that starts with this text.
    columns: dict[str, str]
    date_columns: list[str] = ["applied_on", "followup_on"]
    statuses: Statuses = Statuses()
    values: Values = Values()
    followups: Followups = Followups()
    messages: Messages = Messages()
    funnel: Funnel = Funnel()
    rules: Rules = Rules()
    safety: Safety = Safety()

    @field_validator("columns")
    @classmethod
    def _required_columns(cls, cols: dict[str, str]) -> dict[str, str]:
        missing = [f for f in REQUIRED_FIELDS if f not in cols]
        if missing:
            raise ValueError(f"[columns] must define {', '.join(missing)}")
        return cols

    def has(self, field: str) -> bool:
        return field in self.columns

    def is_date(self, field: str) -> bool:
        """A column of dates: the configured ones, and the reply date when there is one."""
        return field in self.date_columns or field == "replied_on"

    def custom_fields(self) -> list[str]:
        return [f for f in self.columns if f not in KNOWN_FIELDS]


class ConfigError(Exception):
    pass


def config_path(tracker: str | Path, explicit: str | Path | None = None) -> Path:
    """The config file for a tracker: `explicit`, else next to the file
    (my-tracker.xlsx -> my-tracker.tracker-mcp.toml), else for a Google Sheet
    ~/.tracker-mcp/<sheet id>.tracker-mcp.toml."""
    if explicit:
        return Path(explicit).expanduser()
    t = str(tracker)
    if t.startswith(("https://docs.google.com/spreadsheets/", "gsheet:")):
        from .gsheets import default_config_path
        return default_config_path(t)
    path = Path(t).expanduser()
    return path.with_name(f"{path.stem}.tracker-mcp.toml")


def load_config(tracker: str | Path, explicit: str | Path | None = None) -> Config:
    path = config_path(tracker, explicit)
    if not path.exists():
        raise ConfigError(f"No config at {path}. Run `tracker-mcp init {tracker}` to write one from "
                          "your tracker's headers, or `tracker-mcp example` for a commented template.")
    return load_config_file(path)


def load_config_file(path: Path) -> Config:
    return parse_config(path.read_text(encoding="utf-8"), str(path))


def parse_config(text: str, where: str = "config") -> Config:
    try:
        return Config.model_validate(tomllib.loads(text))
    except tomllib.TOMLDecodeError as e:
        raise ConfigError(f"{where}: not valid TOML: {e}") from e
    except ValidationError as e:
        raise ConfigError(f"{where}: {e}") from e


def with_column(text: str, field: str, header: str) -> str:
    """The config text with `field = "header"` added at the end of [columns], keeping
    everything else (comments, order) as it was."""
    lines = text.splitlines(keepends=True)
    start = next((i for i, line in enumerate(lines) if re.fullmatch(r"\[columns\]\s*(#.*)?", line.strip())), None)
    if start is None:
        raise ConfigError("The config has no [columns] section.")
    end = next((i for i in range(start + 1, len(lines)) if lines[i].lstrip().startswith("[")), len(lines))
    while end > start + 1 and not lines[end - 1].strip():   # before the blank lines that end the section
        end -= 1
    if lines[end - 1] and not lines[end - 1].endswith("\n"):
        lines[end - 1] += "\n"
    value = json.dumps(header, ensure_ascii=False)   # a valid TOML basic string too
    lines.insert(end, f"{field} = {value}\n")
    return "".join(lines)
