"""An in-memory stand-in for the Google Sheets API: the subset gsheets.py uses."""

from __future__ import annotations

import json
import urllib.parse
from datetime import datetime, timedelta

EPOCH = datetime(1899, 12, 30)


class FakeSheets:
    def __init__(self, sheets: dict[str, list[list]], dropdowns: dict[str, dict[int, list[str]]] | None = None):
        # title -> {(row, col): (value, is_date)}; rows and cols are 1-based
        self.sheets: dict[str, dict] = {}
        for i, (title, rows) in enumerate(sheets.items()):
            cells = {}
            for r, row in enumerate(rows, 1):
                for c, v in enumerate(row, 1):
                    if v is None:
                        continue
                    if isinstance(v, datetime):
                        cells[(r, c)] = ((v - EPOCH).days, True)
                    else:
                        cells[(r, c)] = (v, False)
            self.sheets[title] = {"id": 100 + i, "rows": 20, "cols": 26, "cells": cells,
                                  "lists": (dropdowns or {}).get(title, {})}
        self.calls: list[tuple[str, str]] = []
        self.requests: list[dict] = []

    def value(self, title: str, r: int, c: int):
        v = self.sheets[title]["cells"].get((r, c))
        if v is None:
            return None
        return EPOCH + timedelta(days=v[0]) if v[1] else v[0]

    def set(self, title: str, r: int, c: int, value, kind: str | None = None) -> None:
        """Someone edits the sheet in the browser. kind "stringValue" makes "=..." text."""
        self.sheets[title]["cells"][(r, c)] = (value, False)
        if kind:
            self.sheets[title].setdefault("kinds", {})[(r, c)] = kind
        else:
            self.sheets[title].get("kinds", {}).pop((r, c), None)

    def kind(self, title: str, r: int, c: int) -> str | None:
        """How the last write stored the cell: "stringValue", "formulaValue", ..."""
        return self.sheets[title].get("kinds", {}).get((r, c))

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
                sheet["rows"] += req["appendDimension"]["length"]
            else:
                u = req["updateCells"]
                sheet = self._by_id(u["range"]["sheetId"])
                r, c = u["range"]["startRowIndex"] + 1, u["range"]["startColumnIndex"] + 1
                assert r <= sheet["rows"], "write beyond the grid without appendDimension"
                data = u["rows"][0]["values"][0]
                v = data.get("userEnteredValue")
                if not v:
                    sheet["cells"].pop((r, c), None)
                    continue
                is_date = data.get("userEnteredFormat", {}).get("numberFormat", {}).get("type") == "DATE"
                raw = next(iter(v.values()))
                sheet["cells"][(r, c)] = (raw, is_date)
                sheet.setdefault("kinds", {})[(r, c)] = next(iter(v))
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
                    if (r, c) in s["cells"]:
                        raw, is_date = s["cells"][(r, c)]
                        if isinstance(raw, bool):
                            cell["userEnteredValue"] = {"boolValue": raw}
                        elif isinstance(raw, (int, float)):
                            cell["userEnteredValue"] = {"numberValue": raw}
                            if is_date:
                                cell["effectiveFormat"] = {"numberFormat": {"type": "DATE"}}
                        elif isinstance(raw, str) and raw.startswith("=") and \
                                s.get("kinds", {}).get((r, c)) != "stringValue":
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
