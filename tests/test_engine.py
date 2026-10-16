"""Acceptance checks against the real template and spreadsheet engine.

Set BUDGET_TEST_WORKBOOK to run on another machine. The original is never saved.
"""
import hashlib
import os
from pathlib import Path
import tempfile
import unittest
from zipfile import ZipFile

from budget_engine import (
    BudgetError,
    _formula_signature,
    compare_formula_caches,
    engine_status,
    parse_workbook,
    recalculate_preserving_formulas,
)


def template():
    configured = os.environ.get("BUDGET_TEST_WORKBOOK")
    candidates = [Path(configured)] if configured else list(Path("/workspace/attachments").rglob("2027年预算基础表.xlsx"))
    if not candidates or not candidates[0].is_file():
        raise unittest.SkipTest("Set BUDGET_TEST_WORKBOOK to the original budget template")
    return candidates[0]


class EngineAcceptanceTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.original = template()
        if not engine_status()["available"]:
            raise unittest.SkipTest("LibreOffice is required for real recalculation")
        cls.temporary = tempfile.TemporaryDirectory(prefix="budget-test-", dir="/tmp")
        cls.addClassCleanup(cls.temporary.cleanup)
        directory = Path(cls.temporary.name)
        cls.before_hash = hashlib.sha256(cls.original.read_bytes()).hexdigest()
        cls.baseline = directory / "baseline.xlsx"
        cls.adjusted = directory / "adjusted.xlsx"
        recalculate_preserving_formulas(cls.original, [], cls.baseline)
        recalculate_preserving_formulas(cls.original, [{"sheet": "1月", "cell": "J7", "value": 14}], cls.adjusted)

    def test_no_change_matches_every_formula_cache(self):
        comparison = compare_formula_caches(self.original, self.baseline)
        self.assertTrue(comparison["identical"], comparison["differences"][:20])
        self.assertEqual(comparison["before_formula_count"], 200736)
        self.assertEqual(hashlib.sha256(self.original.read_bytes()).hexdigest(), self.before_hash)

    def test_real_recalculation_and_formula_preservation(self):
        view = parse_workbook(self.adjusted)
        january = view["months"][0]
        route = next(r for r in january["routes"] if r["line"] == "136路")
        self.assertEqual(route["vehicles_a"], 14)
        self.assertEqual(route["reserve_a"], 1)
        self.assertEqual(route["max_regular"], 14)
        self.assertEqual(route["plan_trips"]["A"], 56)
        self.assertEqual(route["plan_trips"]["B"], 48)
        self.assertEqual(route["plan_trips"]["C"], 24)
        self.assertEqual(route["monthly_trips"], 1464)
        self.assertEqual(january["totals"]["monthly_trips"], 374910.5)
        self.assertEqual(january["totals"]["max_regular"], 3573)
        with ZipFile(self.original) as before, ZipFile(self.adjusted) as after:
            self.assertEqual(before.namelist(), after.namelist())
            for member in before.namelist():
                first, last = before.read(member), after.read(member)
                if member.startswith("xl/worksheets/") and member.endswith(".xml"):
                    self.assertEqual(_formula_signature(first), _formula_signature(last), member)
                else:
                    self.assertEqual(first, last, member)
        comparison = compare_formula_caches(self.original, self.adjusted)
        self.assertGreater(comparison["difference_count"], 0)
        for changed in comparison["differences"]:
            self.assertIn(changed["sheet"], ("1月", "汇总"))
            if changed["sheet"] == "汇总":
                self.assertLessEqual(int("".join(c for c in changed["cell"] if c.isdigit())), 312)


class TemplateProtectionTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.original = template()

    def test_ambiguous_original_arrangements_are_retained(self):
        view = parse_workbook(self.original)
        self.assertEqual(len(view["months"]), 12)
        self.assertTrue(all(len(month["routes"]) == 304 for month in view["months"]))
        february = view["months"][1]
        self.assertEqual(february["days"], 28)
        self.assertEqual(len(february["dates"]), 31)
        self.assertEqual(february["dates"][-1]["date"], "2027-03-03")
        self.assertIs(february["dates"][-1]["valid"], False)
        self.assertEqual(next(w for w in view["warnings"] if w["code"] == "cross_month_references")["count"], 608)
        warnings = [w for route in view["months"][0]["routes"] for w in route["warnings"]]
        self.assertEqual({w["cell"] for w in warnings if w["code"] == "negative_reserve"}, {"N259", "N260", "Y260"})
        self.assertTrue(any(w["code"] == "missing_trip_input" and w["cell"] == "AK201" for w in warnings))

    def test_formula_cells_and_original_are_protected(self):
        with tempfile.TemporaryDirectory(prefix="budget-protection-", dir="/tmp") as temporary:
            forbidden = Path(temporary) / "forbidden.xlsx"
            with self.assertRaises(BudgetError):
                recalculate_preserving_formulas(self.original, [{"sheet": "1月", "cell": "H7", "value": 14}], forbidden)
            with self.assertRaises(BudgetError):
                recalculate_preserving_formulas(self.original, [{"sheet": "1月", "cell": "J7", "value": 14}], self.original)
            self.assertFalse(forbidden.exists())


if __name__ == "__main__":
    unittest.main()
