"""Real-workbook API acceptance checks; LibreOffice is never mocked.

Run from the repository root with ``python -m unittest discover -s tests -p
'test_app.py' -v``. Data is isolated in /tmp, and the uploaded source is read only.
The complete lifecycle uses five actual spreadsheet recalculations.
"""
from __future__ import annotations

import hashlib
import importlib
from io import BytesIO
import json
import os
from pathlib import Path
import re
import sys
import tempfile
import unittest
import xml.etree.ElementTree as ET
from zipfile import ZipFile

from fastapi.testclient import TestClient


_sources = list(Path('/workspace/attachments').rglob('2027年预算基础表.xlsx'))
SOURCE = Path(os.environ.get('TEST_WORKBOOK') or (_sources[0] if _sources else 'example.xlsx'))
NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
DOC_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CELL_RE = re.compile(rb'<c\b[^>]*\br="([^"]+)"[^>]*(?:/>|>(.*?)</c>)', re.S)
FORMULA_RE = re.compile(rb'<f\b[^>]*(?:/>|>.*?</f>)', re.S)


def digest(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def formulas(path_or_bytes):
    """Independent, byte-for-byte signatures including shared-formula tags."""
    source = BytesIO(path_or_bytes) if isinstance(path_or_bytes, bytes) else path_or_bytes
    result = {}
    with ZipFile(source) as archive:
        for name in archive.namelist():
            if not (name.startswith("xl/worksheets/") and name.endswith(".xml")):
                continue
            for cell in CELL_RE.finditer(archive.read(name)):
                formula = FORMULA_RE.search(cell.group(2) or b"")
                if formula is not None:
                    result[(name, cell.group(1))] = formula.group(0)
    return result


def sheet_cells(payload, name):
    with ZipFile(BytesIO(payload)) as archive:
        workbook = ET.fromstring(archive.read("xl/workbook.xml"))
        links = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        targets = {node.get("Id"): node.get("Target") for node in links}
        sheet = next(s for s in workbook.find("m:sheets", NS) if s.get("name") == name)
        target = targets[sheet.get("{" + DOC_REL + "}id")]
        target = target.lstrip("/") if target.startswith("/") else "xl/" + target
        root = ET.fromstring(archive.read(target))
        return {node.get("r"): node for node in root.findall(".//m:sheetData/m:row/m:c", NS)}


class BudgetAPIAcceptanceTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        if not SOURCE.is_file():
            raise unittest.SkipTest("The real uploaded 2027 budget workbook is not available.")
        cls.temp = tempfile.TemporaryDirectory(prefix="budget-api-acceptance-", dir="/tmp")
        cls.previous_data = os.environ.get("BUS_BUDGET_DATA_DIR")
        os.environ["BUS_BUDGET_DATA_DIR"] = cls.temp.name
        cls.source_hash = digest(SOURCE)
        cls.original_formulas = formulas(SOURCE)
        sys.modules.pop("app", None)
        cls.module = importlib.import_module("app")
        cls.client = TestClient(cls.module.app)
        cls.client.__enter__()
        response = cls.client.post("/api/import", files={
            "file": (SOURCE.name, SOURCE.read_bytes(),
                     "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        })
        if response.status_code != 200:
            cls.client.__exit__(None, None, None)
            cls.temp.cleanup()
            raise AssertionError(f"Real workbook import failed: {response.status_code} {response.text[:2000]}")
        cls.baseline = response.json()
        cls.v0_id = cls.baseline["version"]["id"]
        active = json.loads((Path(cls.temp.name) / "active.json").read_text())
        cls.dataset = Path(cls.temp.name) / "datasets" / active["id"]

    @classmethod
    def tearDownClass(cls):
        cls.client.__exit__(None, None, None)
        cls.temp.cleanup()
        if cls.previous_data is None:
            os.environ.pop("BUS_BUDGET_DATA_DIR", None)
        else:
            os.environ["BUS_BUDGET_DATA_DIR"] = cls.previous_data

    def post(self, url, payload=None, expected=200):
        response = self.client.post(url, json=payload) if payload is not None else self.client.post(url)
        self.assertEqual(response.status_code, expected, response.text[:2000])
        return response.json()

    def row(self, view, key="1月:J7"):
        return next(r for r in view["rows"] if r["key"] == key)

    def assert_route(self, view, vehicles, trips):
        row = self.row(view)
        self.assertEqual((row["month"], row["route"], row["source_row"]), ("1月", "136路", 7))
        self.assertEqual(row["vehicles"], vehicles)
        self.assertEqual(row["monthly_trips"], trips)
        self.assertEqual(row["original_vehicles"], 15)
        self.assertEqual(row["original_monthly_trips"], 1536)
        self.assertEqual(row["a_days"], 18)
        return row

    def assert_original_formulas(self, payload):
        actual = formulas(payload)
        for location, original in self.original_formulas.items():
            self.assertEqual(actual.get(location), original, repr(location))

    def test_01_import_no_change_and_original_anomalies(self):
        meta, v0 = self.baseline["dataset"], self.baseline["version"]
        self.assertTrue(meta["baseline_verified"], meta["baseline_differences"][:10])
        self.assertEqual(meta["baseline_differences"], [])
        self.assertIn("LibreOffice", meta["baseline_engine"])
        self.assertEqual(meta["sha256"], self.source_hash)
        self.assertEqual(meta["months"], [f"{month}月" for month in range(1, 13)])
        self.assertEqual(len(v0["rows"]), 304 * 12)
        self.assertEqual(v0["label"], "V0")
        self.assertIsNone(v0["parent_id"])
        self.assertEqual(v0["operations"], [])
        self.assertEqual(v0["totals"]["trip_delta"], 0)
        row = self.assert_route(v0, 15, 1536)
        self.assertEqual((row["day_trips"], row["max_vehicles"], row["reserve"]), (60, 15, 0))
        self.assertGreater(len(self.original_formulas), 200000)
        self.assertEqual(digest(self.dataset / "original.xlsx"), self.source_hash)
        self.assertEqual(digest(self.dataset / "versions" / f"{self.v0_id}.xlsx"), self.source_hash)
        self.assertEqual(formulas(self.dataset / "baseline-calculated.xlsx"), self.original_formulas)
        codes = {w["code"] for w in v0["warnings"]}
        self.assertIn("negative_reserve", codes)
        self.assertIn("out_of_month_dates", codes)
        self.assertTrue(all(w["origin"] == "existing" for w in v0["warnings"]))
        self.assertEqual(digest(SOURCE), self.source_hash)

    def test_02_adjustment_undo_versions_restore_and_exports(self):
        state = self.post("/api/draft", {"version_id": self.v0_id})
        self.assert_route(state["draft"], 15, 1536)
        draft_file = self.dataset / "draft.json"
        draft_before = draft_file.read_bytes()
        # A locked member rejects the entire batch. This source has no original
        # zero A-plan inputs, so create zero through one supported reduction.
        self.post("/api/draft/adjust", {"keys": ["1月:J7"], "locked_routes": ["136路"]}, 422)
        self.assertEqual(draft_file.read_bytes(), draft_before)
        one = next(r for r in state["draft"]["rows"] if r["month"] == "1月" and r["vehicles"] == 1)
        zero_state = self.post("/api/draft/adjust", {"keys": [one["key"]], "reason": "准备零投放边界验收"})
        zero = self.row(zero_state["draft"], one["key"])
        self.assertEqual(zero["vehicles"], 0)
        zero_before = draft_file.read_bytes()
        self.post("/api/draft/adjust", {"keys": ["1月:J7", zero["key"]]}, 422)
        self.assertEqual(draft_file.read_bytes(), zero_before)
        self.post("/api/draft/adjust", {"keys": [zero["key"]]}, 422)
        self.assertEqual(draft_file.read_bytes(), zero_before)
        reset = self.post("/api/draft/undo")["draft"]
        self.assertEqual(self.row(reset, one["key"])["vehicles"], 1)
        self.assertEqual(reset["changes"], [])
        draft_before = draft_file.read_bytes()
        self.post("/api/draft/adjust", {"keys": ["1月:J7"], "delta": -2}, 422)
        self.post("/api/draft/adjust", {"keys": ["1月:J7"], "delta": -1.5}, 422)
        self.post("/api/draft/adjust", {"keys": ["1月:J7", "1月:J99999"]}, 422)
        self.post("/api/draft/save", {"reason": "尚未调整"}, 422)
        self.assertEqual(draft_file.read_bytes(), draft_before)

        state = self.post("/api/draft/adjust", {"keys": ["1月:J7"], "reason": "验收：A计划营运1减1"})
        adjusted = state["draft"]
        row = self.assert_route(adjusted, 14, 1464)
        self.assertEqual((row["day_trips"], row["max_vehicles"], row["reserve"]), (56, 14, 1))
        self.assertEqual(adjusted["totals"]["trip_delta"], -72)
        self.assertEqual(adjusted["totals"]["affected_rows"], 1)
        before_rows = {r["key"]: r for r in self.baseline["version"]["rows"]}
        for current in adjusted["rows"]:
            if current["key"] != "1月:J7":
                self.assertEqual(current, before_rows[current["key"]], current["key"])
        original_daily = self.row(self.baseline["version"])["daily"]
        for before, after in zip(original_daily, row["daily"]):
            self.assertEqual(after["plan"], before["plan"])
            self.assertEqual(after["trips"], before["trips"] - (4 if before["plan"] == "A" else 0))
        adjusted_bytes = Path(adjusted["workbook"]).read_bytes()
        self.assertEqual(formulas(adjusted_bytes), self.original_formulas)
        month_cells = sheet_cells(adjusted_bytes, "1月")
        summary_cells = sheet_cells(adjusted_bytes, "汇总")
        for cell, value in {"J7": "14", "N7": "1", "H7": "14", "H55": "714", "H315": "3573"}.items():
            self.assertEqual(month_cells[cell].findtext("m:v", namespaces=NS), value)
        for cell, value in {"F4": "56", "G4": "48", "H4": "24", "T4": "56", "M4": "24"}.items():
            self.assertEqual(summary_cells[cell].findtext("m:v", namespaces=NS), value)
        self.post("/api/draft/save", {"reason": "   "}, 422)
        self.post("/api/draft/save", {"reason": ""}, 422)
        self.assertEqual(self.post("/api/draft", {"version_id": self.v0_id})["draft"]["id"], adjusted["id"])

        undone = self.post("/api/draft/undo")["draft"]
        self.assert_route(undone, 15, 1536)
        self.assertEqual(undone["changes"], [])
        self.assertEqual(undone["operations"], [])
        self.assertEqual(undone["audit_history"][-1]["type"], "reduce")
        self.assertTrue(undone["audit_history"][-1]["undone_at"])
        self.assertEqual(digest(Path(undone["workbook"])), self.source_hash)

        self.post("/api/draft/adjust", {"keys": ["1月:J7"], "reason": "正式验收减车"})
        saved = self.post("/api/draft/save", {"reason": "V1：1月136路A计划营运1减一台"})
        v1 = saved["version"]
        self.assertEqual(v1["label"], "V1")
        self.assertEqual(v1["parent_id"], self.v0_id)
        self.assertIsNone(saved["draft"])
        self.assert_route(v1, 14, 1464)
        self.assertEqual(v1["operations"][0]["scope"][0]["before"], 15)
        self.assertEqual(v1["operations"][0]["scope"][0]["after"], 14)
        self.assertTrue(v1["audit_history"])

        # A V1 child restores to V1's 14, rather than blindly returning to V0's 15.
        self.post("/api/draft", {"version_id": v1["id"]})
        reduced_again = self.post("/api/draft/adjust", {"keys": ["1月:J7"], "reason": "V2草稿再减一台"})["draft"]
        self.assert_route(reduced_again, 13, 1392)
        self.post("/api/draft", {"version_id": self.v0_id}, 409)
        restored = self.post("/api/draft/restore", {"keys": ["1月:J7"], "reason": "恢复本草稿父版本"})["draft"]
        self.assert_route(restored, 14, 1464)
        self.assertEqual(restored["changes"], [])
        self.assertEqual(restored["operations"][-1]["type"], "restore")
        self.assertEqual(restored["operations"][-1]["scope"][0]["after"], 14)
        v2 = self.post("/api/draft/save", {"reason": "V2：局部恢复到父版本并保留记录"})["version"]
        self.assertEqual(v2["label"], "V2")
        self.assertEqual(v2["parent_id"], v1["id"])
        self.assert_route(v2, 14, 1464)

        workbook = self.client.get("/api/export", params={"version_id": v1["id"], "kind": "workbook"})
        self.assertEqual(workbook.status_code, 200, workbook.text[:200] if workbook.status_code != 200 else "")
        self.assertIn("V1", workbook.headers["content-disposition"])
        self.assertEqual(formulas(workbook.content), self.original_formulas)
        self.assertEqual(sheet_cells(workbook.content, "1月")["J7"].findtext("m:v", namespaces=NS), "14")
        monthly = self.client.get("/api/export", params={"version_id": v1["id"], "kind": "monthly"})
        self.assertEqual(monthly.status_code, 200, monthly.text[:200] if monthly.status_code != 200 else "")
        self.assertIn("V1", monthly.headers["content-disposition"])
        self.assert_original_formulas(monthly.content)
        report = sheet_cells(monthly.content, "月趟次试算")
        report_text = " ".join("".join(cell.itertext()) for cell in report.values())
        for expected in ("V1", v1["id"], v1["reason"], self.source_hash, "非已确认报送模板"):
            self.assertIn(expected, report_text)
        route_row = next(int(ref[1:]) for ref, cell in report.items()
                         if ref.startswith("C") and "".join(cell.itertext()) == "136路")
        self.assertEqual(report[f"D{route_row}"].findtext("m:v", namespaces=NS), "1464")
        self.assertEqual(report[f"D{route_row}"].findtext("m:f", namespaces=NS), "SUM('汇总'!M4:AQ4)")
        with ZipFile(BytesIO(workbook.content)) as before, ZipFile(BytesIO(monthly.content)) as after:
            for member in before.namelist():
                if member.startswith("xl/worksheets/"):
                    self.assertEqual(before.read(member), after.read(member), member)

        original_export = self.client.get("/api/export", params={"version_id": self.v0_id, "kind": "workbook"})
        self.assertEqual(original_export.status_code, 200)
        self.assertEqual(hashlib.sha256(original_export.content).hexdigest(), self.source_hash)
        latest = self.client.get("/api/state").json()
        self.assertEqual([v["label"] for v in latest["dataset"]["versions"]], ["V0", "V1", "V2"])
        self.assertEqual(latest["dataset"]["versions"][0]["id"], self.v0_id)
        self.assertEqual(digest(SOURCE), self.source_hash)
        self.assertEqual(digest(self.dataset / "original.xlsx"), self.source_hash)
        self.assertEqual(digest(self.dataset / "versions" / f"{self.v0_id}.xlsx"), self.source_hash)

    def test_03_health_static_and_invalid_export(self):
        self.assertEqual(self.client.get("/api/health").status_code, 200)
        page = self.client.get("/")
        self.assertEqual(page.status_code, 200)
        self.assertIn("营运预算", page.text)
        for asset in re.findall(r'(?:src|href)="(/[^"?#]+\.(?:js|css))"', page.text):
            self.assertEqual(self.client.get(asset).status_code, 200, asset)
        self.assertEqual(self.client.get("/api/export", params={"version_id": "missing"}).status_code, 404)
        self.assertEqual(self.client.get("/api/export", params={"version_id": self.v0_id, "kind": "invalid"}).status_code, 422)
        self.post("/api/import", {}, 422)


if __name__ == "__main__":
    unittest.main()
