"""Where a tracker lives: an .xlsx file, a .csv file or a Google Sheet.

A store hands the tracker an in-memory workbook (openpyxl) together with a
version, and writes a changed workbook back. The tracker's logic only ever sees
the workbook, so every store gets the same tools and the same safety rules:

- `load()` returns the workbook, its version and any warnings about parts of
  the source that a save would lose;
- `save()` refuses if the version moved since `load()`, makes a backup, writes,
  and lets the tracker verify what was written before the write counts;
- `restore()` puts a backup back (undo).

The Google Sheets store lives in gsheets.py; `open_store` picks a store by target.
"""

from __future__ import annotations

import csv
import hashlib
import io
import os
import re
import shutil
import tempfile
import threading
import time
import warnings
import zipfile
from collections.abc import Callable
from contextlib import contextmanager
from datetime import date, datetime, timedelta
from io import BytesIO
from pathlib import Path
from typing import Protocol

import openpyxl
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl.utils.exceptions import InvalidFileException
from openpyxl.workbook.workbook import Workbook

KEEP_BACKUPS = 200
MAX_TEXT = 32767   # characters openpyxl keeps in one cell


class TrackerError(Exception):
    """A request the tracker cannot apply; the message says why."""


class ConflictError(TrackerError):
    """The source changed (or is open elsewhere) under us; re-read and retry."""


class UnverifiedWrite(TrackerError):
    """The write was sent but could not be read back to check it."""

    def __init__(self, backup: str, message: str):
        super().__init__(message)
        self.backup = backup


def digest(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()[:16]


_now = datetime.now   # the clock for backup names; tests freeze it


def backup_path(folder: Path, stem: str, suffix: str) -> Path:
    """<stem>.<yyyymmdd-hhmmss-micro><suffix>, never an existing file: Windows' clock can
    return the same microsecond twice, and a backup taken right after another one (undo
    right after a write) would overwrite it. Names still sort by time."""
    now = _now()
    while True:
        path = folder / f"{stem}.{now:%Y%m%d-%H%M%S-%f}{suffix}"
        if not path.exists():
            return path
        now += timedelta(microseconds=1)


Verify = Callable[[Workbook], None]


class Store(Protocol):
    kind: str                  # "xlsx", "csv" or "gsheet"
    label: str                 # how messages name the tracker
    backup_dir: Path           # backups and audit.jsonl
    single_sheet: bool         # True when the source has exactly one sheet (CSV)

    def load(self) -> tuple[Workbook, str, list[str]]: ...
    def version(self) -> str: ...
    def check_writable(self) -> None: ...
    def save(self, wb: Workbook, before: str, verify: Verify) -> tuple[str, str]: ...
    def restore(self, backup: str) -> tuple[str, str]: ...


def is_google_sheet(target: str) -> bool:
    return target.startswith(("https://docs.google.com/spreadsheets/", "gsheet:"))


def open_store(target: str | Path, backup_dir: str | Path | None = None) -> Store:
    t = str(target)
    if is_google_sheet(t):
        from .gsheets import SheetsStore
        return SheetsStore(t, backup_dir)
    suffix = Path(t).suffix.lower()
    if suffix == ".csv":
        return CsvStore(t, backup_dir)
    if suffix == ".xlsx":
        return XlsxStore(t, backup_dir)
    # .xlsm would lose its macros and its file type on save; .xls and .ods are other formats.
    raise TrackerError(f"tracker-mcp works with .xlsx and .csv files and Google Sheets, not «{suffix or t}». "
                       "Save a copy of the tracker as .xlsx (Excel or Numbers: File → Save As / Export) "
                       "and point the server to it.")


# ---------------------------------------------------------------- files

_THREAD_LOCKS: dict[str, threading.Lock] = {}
_THREAD_LOCKS_GUARD = threading.Lock()


@contextmanager
def file_lock(path: Path):
    """One writer at a time for one tracker: a lock between threads of this server and an OS
    lock between server processes, held from the last version check to the replace."""
    real = path.resolve()   # the same file through a link or another path takes the same lock
    key = str(real)
    with _THREAD_LOCKS_GUARD:
        tl = _THREAD_LOCKS.setdefault(key, threading.Lock())
    with tl:
        lock_file = real.with_name(f".{real.name}.tracker-mcp.lock")
        with open(lock_file, "a+b") as fh:
            if os.name == "nt":
                import msvcrt
                fh.seek(0)
                for _ in range(100):   # about ten seconds
                    try:
                        msvcrt.locking(fh.fileno(), msvcrt.LK_NBLCK, 1)
                        break
                    except OSError:
                        time.sleep(0.1)
                else:
                    raise ConflictError("Another write to this tracker is still running; retry in a moment.")
                try:
                    yield
                finally:
                    fh.seek(0)
                    msvcrt.locking(fh.fileno(), msvcrt.LK_UNLCK, 1)
            else:
                import fcntl
                fcntl.flock(fh.fileno(), fcntl.LOCK_EX)
                try:
                    yield
                finally:
                    fcntl.flock(fh.fileno(), fcntl.LOCK_UN)


def backup_pattern(stem: str, suffix: str) -> re.Pattern:
    """The backups of exactly this tracker: «jobs.2025.xlsx» is not a backup of «jobs.xlsx»."""
    return re.compile(re.escape(stem) + r"\.\d{8}-\d{6}-\d{6}" + re.escape(suffix) + r"$")


class _FileStore:
    """What .xlsx and .csv share: a local file, lock files, backups next to it,
    and the write path temp file -> verify -> version check -> backup -> rename."""

    kind = "file"
    single_sheet = False

    def __init__(self, path: str | Path, backup_dir: str | Path | None = None):
        self.path = Path(path).expanduser()
        self.label = str(self.path)
        # backups and the audit live by the real file, under its name, so every path to it (a link,
        # a relative path, `~`) shares one history; backups made before 0.8.4 next to a link stay there
        self.backup_dir = (Path(backup_dir).expanduser() if backup_dir
                           else self.path.resolve().parent / "backups-tracker-mcp")

    def _bytes(self) -> bytes:
        if not self.path.exists():
            raise TrackerError(f"Tracker file not found: {self.path}")
        return self.path.read_bytes()

    def version(self) -> str:
        return digest(self._bytes())

    def check_writable(self) -> None:
        real = self.path.resolve()   # Excel puts its lock file beside the file it opened: the link or the real one
        for folder, name in {(self.path.parent, self.path.name), (real.parent, real.name)}:
            for lock in (f"~${name}", f".~lock.{name}#"):
                if (folder / lock).exists():
                    raise ConflictError(f"The tracker is open in Excel/LibreOffice (lock file {lock}). "
                                        "Close it and retry; if it is closed, the lock file is stale.")

    def _render(self, wb: Workbook) -> bytes:
        raise NotImplementedError

    def _parse(self, data: bytes) -> Workbook:
        raise NotImplementedError

    def _tmp(self) -> Path:
        # a temp file of its own per write: two writers never share one
        # next to the real file: a write through a link replaces the file, not the link
        real = self.path.resolve()
        fd, name = tempfile.mkstemp(dir=real.parent, prefix=f".{real.stem}.tracker-mcp-", suffix=real.suffix)
        os.close(fd)
        return Path(name)

    def owns_backup(self, name: str) -> bool:
        real = self.path.resolve()   # by the real file's name; a link's own name for older backups
        return any(backup_pattern(p.stem, p.suffix).fullmatch(name) for p in {real, self.path})

    def _replace(self, tmp: Path) -> None:
        try:
            os.replace(tmp, self.path.resolve())   # the real file: a link stays a link
        except OSError as e:  # Windows: OneDrive, an antivirus or Excel holding the file
            raise ConflictError(f"Could not replace {self.path.name}: {e.strerror or e}. Another program "
                                "(Excel, OneDrive, an antivirus) may be holding it. Nothing was changed; "
                                "retry in a moment.") from e

    def save(self, wb: Workbook, before: str, verify: Verify) -> tuple[str, str]:
        tmp = self._tmp()
        try:
            tmp.write_bytes(self._render(wb))
            verify(self._parse(tmp.read_bytes()))
            with file_lock(self.path):
                if self.version() != before:  # someone saved while we worked
                    raise ConflictError("The tracker changed while writing; nothing was saved. Retry.")
                backup = self._backup()
                shutil.copymode(self.path, tmp)
                self._replace(tmp)
                after = self.version()
        finally:
            tmp.unlink(missing_ok=True)
        return backup.name, after

    def restore(self, backup: str) -> tuple[str, str]:
        self.check_writable()
        if not self.owns_backup(backup):   # a bare name of this tracker's backup, nothing else
            raise TrackerError(f"«{backup}» is not a backup of {self.path.name}; nothing restored.")
        source = self.backup_dir / backup
        if not source.exists():
            raise TrackerError(f"Backup {backup} is gone; nothing to restore.")
        tmp = self._tmp()
        try:
            shutil.copy2(source, tmp)
            with file_lock(self.path):
                safety = self._backup()
                self._replace(tmp)
                after = self.version()
        finally:
            tmp.unlink(missing_ok=True)
        return safety.name, after

    def _backup(self) -> Path:
        self.backup_dir.mkdir(parents=True, exist_ok=True)
        real = self.path.resolve()
        dst = backup_path(self.backup_dir, real.stem, real.suffix)
        shutil.copy2(self.path, dst)
        pat = backup_pattern(real.stem, real.suffix)
        old = sorted(p for p in self.backup_dir.iterdir() if pat.fullmatch(p.name))
        for p in old[:-KEEP_BACKUPS]:
            p.unlink()
        return dst


# Parts of an .xlsx that openpyxl writes back (or that Excel rebuilds by itself). Anything else in a
# workbook is dropped by a save, so its presence refuses writes: allow, do not block.
_KEPT_PARTS = ("[Content_Types].xml", "_rels/", "docProps/", "xl/workbook.xml", "xl/_rels/", "xl/styles.xml",
               "xl/sharedStrings.xml", "xl/theme/", "xl/worksheets/", "xl/comments", "xl/drawings/", "xl/charts/",
               "xl/media/", "xl/tables/", "xl/externalLinks/", "xl/pivotTables/", "xl/pivotCache/",
               "xl/calcChain.xml", "xl/printerSettings/", "customXml/")
_PART_NAMES = {"xl/threadedComments/": "threaded comments (Excel 365)", "xl/persons/": "threaded comments (Excel 365)",
               "xl/featurePropertyBag/": "checkboxes in cells", "xl/richData/": "pictures or data types in cells",
               "xl/metadata.xml": "dynamic-array formulas", "xl/slicers/": "slicers", "xl/slicerCaches/": "slicers",
               "xl/timelines/": "timelines", "xl/ctrlProps/": "form controls", "xl/activeX/": "ActiveX controls",
               "xl/embeddings/": "embedded objects", "xl/vbaProject.bin": "macros",
               "xl/connections.xml": "data connections",
               "xl/queryTables/": "query tables", "xl/webextensions/": "add-ins"}


def _parts_a_save_drops(z: zipfile.ZipFile, names: list[str]) -> list[str]:
    """What an openpyxl save would silently lose from this workbook, in words for the user."""
    out = []
    for n in names:
        if not n.startswith(_KEPT_PARTS):
            what = next((v for k, v in _PART_NAMES.items() if n.startswith(k)), f"the part {n}")
            if what not in out:
                out.append(what)
    def text(n: str) -> str:
        try:
            return z.read(n).decode("utf-8", "replace")
        except KeyError:
            return ""
    if any(re.search(r"<xdr:(?:sp|grpSp|cxnSp)\b", text(n)) for n in names if n.startswith("xl/drawings/drawing")):
        out.append("shapes (text boxes, arrows, grouped drawings)")
    if "xfComplement" in text("xl/styles.xml"):
        out.append("checkboxes in cells")
    if re.search(r"<(?:\w+:)?r>", text("xl/sharedStrings.xml")) or any(
            re.search(r"<is>\s*<(?:\w+:)?r>", text(n)) for n in names if n.startswith("xl/worksheets/sheet")):
        out.append("mixed formatting inside a cell (part bold, part plain)")
    return [f"{w} (a save would drop them)" for w in dict.fromkeys(out)]


class XlsxStore(_FileStore):
    kind = "xlsx"

    def _parse(self, data: bytes) -> Workbook:
        try:
            return openpyxl.load_workbook(BytesIO(data))
        except (zipfile.BadZipFile, InvalidFileException, KeyError) as e:
            raise TrackerError(f"{self.path.name} is not a readable .xlsx file ({e}). If it was saved "
                               "in another format, save it again as an Excel Workbook (.xlsx).") from e

    def load(self) -> tuple[Workbook, str, list[str]]:
        data = self._bytes()
        with warnings.catch_warnings(record=True) as caught:
            warnings.simplefilter("always")
            wb = self._parse(data)
        # Parts of the file openpyxl cannot keep (e.g. dropdowns that newer Excel stores
        # in an extension): a save would silently drop them, so writes are refused.
        lossy = sorted({str(w.message) for w in caught if "will be removed" in str(w.message)})
        # pictures need Pillow to survive an openpyxl save; without it they vanish on the first write
        try:
            z = zipfile.ZipFile(BytesIO(data))
            names = z.namelist()
        except zipfile.BadZipFile:
            z, names = None, []
        lossy += _parts_a_save_drops(z, names) if z is not None else []
        if any(n.startswith("xl/media/") for n in names):
            try:
                import PIL  # noqa: F401
            except ImportError:
                lossy.append("pictures in the workbook (they would be dropped on save; `pip install pillow` "
                             "in the server's environment keeps them)")
        return wb, digest(data), lossy

    def _render(self, wb: Workbook) -> bytes:
        wb.calculation.fullCalcOnLoad = True  # Excel recomputes formulas on open
        out = BytesIO()
        wb.save(out)
        return out.getvalue()


class CsvStore(_FileStore):
    """A CSV export (Notion, Huntr, Teal, Google Sheets...). One sheet, no dropdowns:
    statuses come from the config. Dates are written as yyyy-mm-dd text."""

    kind = "csv"
    single_sheet = True

    def __init__(self, path: str | Path, backup_dir: str | Path | None = None):
        super().__init__(path, backup_dir)
        self.delimiter, self.bom, self.encoding, self.newline = ",", False, "utf-8", "\r\n"
        self.lossy: list[str] = []
        self.sheet = re.sub(r"[\[\]:*?/\\]", "_", self.path.stem)[:31] or "Sheet1"

    def _parse(self, data: bytes) -> Workbook:
        self.bom = data.startswith(b"\xef\xbb\xbf")
        body = data[3:] if self.bom else data
        try:
            text, self.encoding = body.decode("utf-8"), "utf-8"
        except UnicodeDecodeError:
            text, self.encoding = _cp1252(body, self.path.name), "cp1252"   # Excel's plain "CSV" on Windows
        end = text.find("\n")   # the first row's line ending decides
        self.newline = "\r\n" if end < 0 or (end > 0 and text[end - 1] == "\r") else "\n"
        try:
            self.delimiter = csv.Sniffer().sniff(text[:4096], delimiters=",;\t").delimiter
        except csv.Error:
            self.delimiter = ","
        wb = Workbook()
        ws = wb.active
        ws.title = self.sheet
        self.lossy = []
        for row in csv.reader(io.StringIO(text, newline=""), delimiter=self.delimiter):
            if any(len(v) > MAX_TEXT for v in row):
                self.lossy.append(f"row {ws.max_row + 1} has a cell with more than {MAX_TEXT} characters, "
                                  "and the rest would be cut off")
            ws.append([ILLEGAL_CHARACTERS_RE.sub(" ", v) if v != "" else None for v in row])
        return wb

    def load(self) -> tuple[Workbook, str, list[str]]:
        data = self._bytes()
        return self._parse(data), digest(data), self.lossy

    def _render(self, wb: Workbook) -> bytes:
        out = io.StringIO(newline="")
        writer = csv.writer(out, delimiter=self.delimiter, lineterminator=self.newline)
        for row in wb.worksheets[0].iter_rows(values_only=True):
            writer.writerow([_csv_text(v) for v in row])
        try:
            body = out.getvalue().encode(self.encoding)
        except UnicodeEncodeError as e:
            raise TrackerError(f"«{e.object[e.start:e.end]}» cannot be written to this CSV file: it is saved "
                               f"as {self.encoding}. Save it as «CSV UTF-8» in Excel and retry.") from e
        return (b"\xef\xbb\xbf" if self.bom else b"") + body


def _cp1252(body: bytes, name: str) -> str:
    """Windows-1252, the encoding of Excel's plain "CSV" in Western Europe and the US.
    Text in another code page (Cyrillic, Greek...) would read as gibberish, so it is
    refused rather than guessed."""
    try:
        text = body.decode("cp1252")
    except UnicodeDecodeError:
        text = ""
    letters = [ch for ch in text if ch.isalpha()]
    accented = sum(1 for ch in letters if "\u00c0" <= ch <= "\u00ff")
    if not text or (letters and accented / len(letters) > 0.3):
        raise TrackerError(f"{name} is not saved as UTF-8, and its encoding cannot be told for sure. "
                           "Save it as «CSV UTF-8» (Excel: File → Save As) and retry.")
    return text


def _csv_text(value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.date().isoformat() if value.time() == datetime.min.time() else value.isoformat(" ")
    if isinstance(value, date):
        return value.isoformat()
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    return str(value)
