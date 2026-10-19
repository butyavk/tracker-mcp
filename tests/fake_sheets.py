"""An in-memory stand-in for the Google Sheets API: the subset gsheets.py uses.

Values and number formats are kept apart, as in Sheets: a write whose field mask
leaves out the number format keeps the cell's format (a date format stays on a
cell even when a plain number is written into it).
"""

from __future__ import annotations

import json
import urllib.parse
from datetime import datetime, timedelta

EPOCH = datetime(1899, 12, 30)
DATE = {"type": "DATE", "pattern": "yyyy-mm-dd"}


class FakeSheets:
    def __init__(self, sheets: dict[str, list[list]], dropdowns: dict[str, dict[int, list[str]]] | None = None):
        # title -> cells {(row, col): raw value}, dates {(row, col) formatted as a date},
        # kinds {(row, col): "stringValue" | "formulaValue" | ...}; rows and cols are 1-based
        self.sheets: dict[str, dict] = {}
        for i, (title, rows) in enumerate(sheets.items()):
            cells, dates = {}, set()
            for r, row in enumerate(rows, 1):
                for c, v in enumerate(row, 1):
                    if v is None:
                        continue
                    if isinstance(v, datetime):
                        cells[(r, c)] = (v - EPOCH).days
                        dates.add((r, c))
                    else:
                        cells[(r, c)] = v
            self.sheets[title] = {"id": 100 + i, "rows": 20, "cols": 26, "cells": cells, "dates": dates,
                                  "kinds": {}, "lists": (dropdowns or {}).get(title, {})}
        self.calls: list[tuple[str, str]] = []
        self.requests: list[dict] = []
        self.fail_next_post: Exception | None = None   # raised after the next batchUpdate is applied

    def value(self, title: str, r: int, c: int):
        s = self.sheets[title]
        v = s["cells"].get((r, c))
        if isinstance(v, (int, float)) and not isinstance(v, bool) and (r, c) in s["dates"]:
            return EPOCH + timedelta(days=v)
        return v

    def is_date(self, title: str, r: int, c: int) -> bool:
        return (r, c) in self.sheets[title]["dates"]

    def set(self, title: str, r: int, c: int, value, kind: str | None = None) -> None:
        """Someone edits the sheet in the browser. kind "stringValue" makes "=..." text."""
        s = self.sheets[title]
        s["cells"][(r, c)] = value
        if kind:
            s["kinds"][(r, c)] = kind
        else:
            s["kinds"].pop((r, c), None)

    def kind(self, title: str, r: int, c: int) -> str | None:
        """How the last write stored the cell: "stringValue", "formulaValue", ..."""
        return self.sheets[title]["kinds"].get((r, c))

    def __call__(self, method: str, url: str, body: dict | None) -> dict:
        self.calls.append((method, url))
        if method == "GET":
            assert "includeGridData=true" in url and urllib.parse.unquote(url)
            return json.loads(json.dumps(self._doc()))
        assert url.endswith(":batchUpdate")
        for req in body["requests"]:
            self.requests.append(req)
            if "appendDimension" in req:
                sheet = self._by_id(req["appendDimension"]["sheetId"])
                dim = "cols" if req["appendDimension"]["dimension"] == "COLUMNS" else "rows"
                sheet[dim] += req["appendDimension"]["length"]
                continue
            u = req["updateCells"]
            sheet = self._by_id(u["range"]["sheetId"])
            r, c = u["range"]["startRowIndex"] + 1, u["range"]["startColumnIndex"] + 1
            assert r <= sheet["rows"], "write beyond the grid without appendDimension"
            assert c <= sheet["cols"], "write beyond the last column without appendDimension"
            data = u["rows"][0]["values"][0]
            mask = u["fields"].split(",")
            if "userEnteredValue" in mask:
                v = data.get("userEnteredValue")
                if v:
                    sheet["cells"][(r, c)] = next(iter(v.values()))
                    sheet["kinds"][(r, c)] = next(iter(v))
                else:
                    sheet["cells"].pop((r, c), None)
                    sheet["kinds"].pop((r, c), None)
            if "userEnteredFormat.numberFormat" in mask:
                if data.get("userEnteredFormat", {}).get("numberFormat", {}).get("type") == "DATE":
                    sheet["dates"].add((r, c))
                else:
                    sheet["dates"].discard((r, c))
        if self.fail_next_post:
            e, self.fail_next_post = self.fail_next_post, None
            raise e
        return {}

    def _by_id(self, sid: int) -> dict:
        return next(s for s in self.sheets.values() if s["id"] == sid)

    def _doc(self) -> dict:
        out = []
        for title, s in self.sheets.items():
            nrows = max([r for r, _ in s["cells"]] + [1])
            ncols = max([c for _, c in s["cells"]] + list(s["lists"]) + [1])
            rows = []
            for r in range(1, nrows + 1):
                values = []
                for c in range(1, ncols + 1):
                    cell: dict = {}
                    if (r, c) in s["dates"]:
                        cell["userEnteredFormat"] = {"numberFormat": DATE}
                    if (r, c) in s["cells"]:
                        raw = s["cells"][(r, c)]
                        if isinstance(raw, bool):
                            cell["userEnteredValue"] = {"boolValue": raw}
                        elif isinstance(raw, (int, float)):
                            cell["userEnteredValue"] = {"numberValue": raw}
                            if (r, c) in s["dates"]:
                                cell["effectiveFormat"] = {"numberFormat": DATE}
                        elif isinstance(raw, str) and raw.startswith("=") and \
                                s["kinds"].get((r, c)) != "stringValue":
                            cell["userEnteredValue"] = {"formulaValue": raw}
                        else:
                            cell["userEnteredValue"] = {"stringValue": str(raw)}
                    if c in s["lists"] and r > 1:
                        cell["dataValidation"] = {"condition": {"type": "ONE_OF_LIST", "values": [
                            {"userEnteredValue": v} for v in s["lists"][c]]}}
                    values.append(cell)
                rows.append({"values": values})
            out.append({"properties": {"sheetId": s["id"], "title": title,
                                       "gridProperties": {"rowCount": s["rows"], "columnCount": s["cols"]}},
                        "data": [{"rowData": rows}]})
        return {"properties": {"title": "Jobs"}, "sheets": out}
