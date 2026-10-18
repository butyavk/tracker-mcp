"""What the server needs to know about a particular tracker workbook.

Nothing about one person's tracker lives in the code: the sheet, the column
headers, the status words, the follow-up rule and the texts the server writes
into the follow-up log all come from a TOML file. By default the file sits next
to the workbook: `my-tracker.xlsx` -> `my-tracker.tracker-mcp.toml`.
`TRACKER_CONFIG` points to another one. `examples/tracker-mcp.example.toml`
describes every key.
"""

from __future__ import annotations

import tomllib
from pathlib import Path

from pydantic import BaseModel, ConfigDict, Field, ValidationError, field_validator

# Fields the server's logic understands. Any other key in [columns] is a custom
# field: it is returned under its own key and can be set through add_application.
REQUIRED_FIELDS = ("company", "role", "status")
KNOWN_FIELDS = REQUIRED_FIELDS + (
    "num", "link", "applied", "applied_on", "channel", "response", "notes",
    "followup_on", "followup_contact", "followup_log",
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

    def custom_fields(self) -> list[str]:
        return [f for f in self.columns if f not in KNOWN_FIELDS]


class ConfigError(Exception):
    pass


def config_path(tracker: str | Path, explicit: str | Path | None = None) -> Path:
    if explicit:
        return Path(explicit).expanduser()
    tracker = Path(tracker).expanduser()
    return tracker.with_name(f"{tracker.stem}.tracker-mcp.toml")


def load_config(tracker: str | Path, explicit: str | Path | None = None) -> Config:
    path = config_path(tracker, explicit)
    if not path.exists():
        raise ConfigError(f"No config at {path}. Copy examples/tracker-mcp.example.toml "
                          "there and set your sheet, headers and statuses.")
    try:
        return Config.model_validate(tomllib.loads(path.read_text(encoding="utf-8")))
    except tomllib.TOMLDecodeError as e:
        raise ConfigError(f"{path}: not valid TOML: {e}") from e
    except ValidationError as e:
        raise ConfigError(f"{path}: {e}") from e
