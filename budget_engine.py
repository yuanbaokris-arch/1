"""Read the original budget workbook and recalculate with a real spreadsheet engine.

The business formulas are never reimplemented here.  The working copy is recalculated
by LibreOffice; only its cached formula results are copied into the original XML.
All formulas, workbook extensions, formatting, and unrelated ZIP entries are retained.
"""
from __future__ import annotations

import calendar
from datetime import datetime, timedelta
import hashlib
import math
import os
from pathlib import Path
import re
import shutil
import subprocess
import tempfile
import time
from zipfile import BadZipFile, ZipFile

from lxml import etree

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
TAG = lambda name: "{" + NS + "}" + name
CELL_RE = re.compile(rb"<c\b(?P<attrs>[^>]*?)(?:/\s*>|>(?P<body>.*?)</c>)", re.S)
ATTR_R = re.compile(rb'\br="([^"]+)"')
ATTR_T = re.compile(rb'\s+t="[^"]*"')
VALUE_RE = re.compile(rb"<(?:v|is)\b[^>]*(?:/>|>.*?</(?:v|is)>)", re.S)
FORMULA_RE = re.compile(rb"<f\b[^>]*(?:/>|>.*?</f>)", re.S)
MONTH_RE = re.compile(r"^(?:[1-9]|1[0-2])月$")
MAX_UNCOMPRESSED_BYTES = 160 * 1024 * 1024


class BudgetError(ValueError):
    """A workbook cannot safely be parsed or adjusted."""


class CalculationError(RuntimeError):
    """The spreadsheet engine failed; no substitute calculation was performed."""


def _xml(data: bytes):
    return etree.fromstring(data, parser=etree.XMLParser(resolve_entities=False, no_network=True))


def _column(index: int) -> str:
    result = ""
    while index:
        index, digit = divmod(index - 1, 26)
        result = chr(65 + digit) + result
    return result


def _number(value):
    return isinstance(value, (int, float)) and not isinstance(value, bool) and math.isfinite(value)


def _sum(values):
    return sum(value for value in values if _number(value))


def _sheet_names(archive: ZipFile) -> dict[str, str]:
    workbook = _xml(archive.read("xl/workbook.xml"))
    rels = _xml(archive.read("xl/_rels/workbook.xml.rels"))
    targets = {r.get("Id"): r.get("Target") for r in rels}
    result = {}
    for sheet in workbook.findall("./" + TAG("sheets") + "/" + TAG("sheet")):
        target = targets[sheet.get("{" + REL_NS + "}id")]
        if target.startswith("/"):
            target = target.lstrip("/")
        else:
            target = "xl/" + target
        if ".." in Path(target).parts:
            raise BudgetError("工作簿包含不支持的工作表路径。")
        result[sheet.get("name")] = target
    return result


def _shared_strings(archive: ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in archive.namelist():
        return []
    root = _xml(archive.read("xl/sharedStrings.xml"))
    return ["".join(node.itertext()) for node in root.findall(TAG("si"))]


def _value(cell, strings):
    kind = cell.get("t")
    if kind == "inlineStr":
        return "".join(cell.find(TAG("is")).itertext()) if cell.find(TAG("is")) is not None else ""
    node = cell.find(TAG("v"))
    if node is None:
        return None
    if node.text is None:
        return "" if kind == "str" else None
    text = node.text
    if kind == "s":
        return strings[int(text)]
    if kind in ("str", "e", "d"):
        return text
    if kind == "b":
        return text == "1"
    try:
        value = float(text)
        return int(value) if value.is_integer() else value
    except (ValueError, OverflowError):
        return text


def _read(path):
    try:
        with ZipFile(path) as archive:
            if sum(info.file_size for info in archive.infolist()) > MAX_UNCOMPRESSED_BYTES:
                raise BudgetError("工作簿解压后过大，请使用原始预算基础表模板。")
            sheets = _sheet_names(archive)
            strings = _shared_strings(archive)
            data = {}
            formula_count = 0
            for name, member in sheets.items():
                root = _xml(archive.read(member))
                cells = {}
                for cell in root.findall("./" + TAG("sheetData") + "/" + TAG("row") + "/" + TAG("c")):
                    formula = cell.find(TAG("f"))
                    if formula is not None:
                        formula_count += 1
                    cells[cell.get("r")] = {
                        "value": _value(cell, strings),
                        "formula": formula.text if formula is not None else None,
                        "has_formula": formula is not None,
                        "formula_shared": dict(formula.attrib) if formula is not None else None,
                        "type": cell.get("t"),
                    }
                data[name] = cells
            return sheets, data, formula_count
    except (BadZipFile, KeyError, etree.XMLSyntaxError, IndexError, OSError) as exc:
        raise BudgetError("无法读取 Excel 工作簿，请上传原始 .xlsx 预算基础表。") from exc


def spreadsheet_executable() -> str | None:
    """Find Calc on a standard macOS application install as well as PATH."""
    configured = os.environ.get("BUS_BUDGET_SOFFICE")
    if configured:
        candidate = Path(configured).expanduser()
        return str(candidate) if candidate.is_file() and os.access(candidate, os.X_OK) else None
    executable = shutil.which("soffice") or shutil.which("libreoffice")
    if executable:
        return executable
    for candidate in (
        Path("/Applications/LibreOffice.app/Contents/MacOS/soffice"),
        Path.home() / "Applications/LibreOffice.app/Contents/MacOS/soffice",
    ):
        if candidate.is_file() and os.access(candidate, os.X_OK):
            return str(candidate)
    return None


def engine_status() -> dict:
    executable = spreadsheet_executable()
    result = {"available": bool(executable), "name": "LibreOffice", "executable": executable}
    if executable:
        try:
            process = subprocess.run([executable, "--version"], capture_output=True, text=True, timeout=10)
            result["version"] = process.stdout.strip()
            result["available"] = process.returncode == 0
        except (OSError, subprocess.TimeoutExpired):
            result["available"] = False
    return result


def _warning(code, message, **context):
    return {"code": code, "message": message, **context}


def parse_workbook(path) -> dict:
    """Return JSON-compatible inputs, actual formula caches, and retained anomalies.

    monthly_trips is a display-only sum of all 31 existing daily result cells. It
    is not a new business formula and includes the original out-of-month columns.
    """
    path = Path(path)
    sheets, data, formula_count = _read(path)
    month_names = sorted((s for s in sheets if MONTH_RE.fullmatch(s)), key=lambda s: int(s[:-1]))
    if len(month_names) != 12 or "汇总" not in data:
        raise BudgetError("模板缺少 1—12 月或汇总工作表，请使用预算基础表。")
    summary = data["汇总"]
    def val(cells, address):
        return cells.get(address, {}).get("value")
    summary_index = {}
    summary_totals = {}
    summary_departments = {}
    for address in summary:
        if not re.fullmatch(r"C[0-9]+", address):
            continue
        row = int(address[1:])
        month, line = val(summary, f"D{row}"), val(summary, address)
        department, team = val(summary, f"A{row}"), val(summary, f"B{row}")
        if line == "合计":
            if department == "全司":
                summary_totals[month] = row
            else:
                summary_departments.setdefault(month, []).append((department, row))
        elif month in month_names and line:
            key = (month, department, team, line)
            if key in summary_index:
                raise BudgetError(f"汇总中线路标识重复，无法可靠对应：{key}。")
            summary_index[key] = row
    warnings = [
        _warning("mileage_unconfirmed", "公里数换算口径尚未确认；本版仅展示原公式计算的趟次，不推算公里数。", severity="info"),
        _warning("minimum_unconfirmed", "最低投放、最低趟次及日期类型关系尚未确认，暂不能判断是否达到业务底线。", severity="info"),
        _warning("decimal_rule_unconfirmed", "原表存在 2.5、3.5、4.5 等单车日趟次，保留原值；允许的小数步长待确认。", severity="info"),
    ]
    months = []
    patterns = []
    for name in month_names:
        cells = data[name]
        year, month = val(cells, "G1"), val(cells, "G2")
        if not _number(year) or not _number(month) or not 1 <= month <= 12:
            raise BudgetError(f"{name}的年份或月份无效。")
        year, month = int(year), int(month)
        days = calendar.monthrange(year, month)[1]
        dates = []
        month_warnings = []
        for day in range(1, 32):
            coordinate = f"{_column(85 + day)}4"
            cached = val(cells, coordinate)
            date = None
            if _number(cached):
                date = (datetime(1899, 12, 30) + timedelta(days=cached)).date().isoformat()
            dates.append({"cell": coordinate, "day": day, "date": date, "valid": day <= days})
        if days < 31:
            month_warnings.append(_warning("out_of_month_dates", f"{name}保留原表 31 列逐日安排，其中 {31-days} 列超出本月；趟次合计仍按原表全部 31 列展示。", sheet=name, days=list(range(days + 1, 32)), severity="warning"))
        routes = []
        for address in cells:
            if not re.fullmatch(r"C[0-9]+", address):
                continue
            row = int(address[1:]); line = val(cells, address)
            if row < 6 or not line or line == "合计":
                continue
            department, team = val(cells, f"A{row}"), val(cells, f"B{row}")
            summary_row = summary_index.get((name, department, team, line))
            if summary_row is None:
                raise BudgetError(f"{name}第 {row} 行线路 {line} 无法唯一对应汇总记录。")
            local_warnings = []
            for plan_index, plan in enumerate("ABCDEFG"):
                start = 9 + 11 * plan_index
                reserve_cell = f"{_column(start + 5)}{row}"
                reserve = val(cells, reserve_cell)
                if _number(reserve) and reserve < 0:
                    local_warnings.append(_warning("negative_reserve", f"{reserve_cell}：{plan}计划机动车数量为 {reserve}，原表已有或本版仍存在负数。", sheet=name, cell=reserve_cell, line=line, plan=plan, severity="warning"))
                for offset in range(4):
                    vehicle_cell = f"{_column(start + 1 + offset)}{row}"
                    trip_cell = f"{_column(start + 6 + offset)}{row}"
                    vehicles, trips = val(cells, vehicle_cell), val(cells, trip_cell)
                    if _number(vehicles) and (vehicles < 0 or not float(vehicles).is_integer()):
                        local_warnings.append(_warning("invalid_vehicle_count", f"{vehicle_cell} 投放数不是非负整数，保留并提示。", sheet=name, cell=vehicle_cell, line=line, severity="warning"))
                    if _number(vehicles) and vehicles > 0 and trips is None:
                        local_warnings.append(_warning("missing_trip_input", f"{vehicle_cell} 有投放 {vehicles} 辆，但单车日趟次 {trip_cell} 为空，需核实。", sheet=name, cell=trip_cell, line=line, plan=plan, severity="warning"))
                    if _number(trips) and trips < 0:
                        local_warnings.append(_warning("negative_trip_input", f"{trip_cell} 单车日趟次为负数，保留并提示。", sheet=name, cell=trip_cell, line=line, severity="warning"))
            daily = []
            for day in range(1, 32):
                plan_cell = f"{_column(85 + day)}{row}"
                trip_cell = f"{_column(12 + day)}{summary_row}"
                plan, trips = val(cells, plan_cell), val(summary, trip_cell)
                daily.append({"day": day, "plan": plan, "trips": trips, "valid": day <= days, "plan_cell": plan_cell, "result_cell": trip_cell})
                if plan not in (None, "", *"ABCDEFG"):
                    local_warnings.append(_warning("unknown_plan", f"{plan_cell} 的计划代码 {plan} 不在 A—G 中。", sheet=name, cell=plan_cell, line=line, severity="warning"))
                if _number(trips) and trips < 0:
                    local_warnings.append(_warning("negative_daily_trips", f"汇总!{trip_cell} 趟次为负数。", sheet="汇总", cell=trip_cell, line=line, severity="warning"))
            def detail(address):
                record = cells.get(address, {})
                return {"cell": address, "value": record.get("value"), "formula": record.get("formula"), "has_formula": record.get("has_formula", False), "shared_formula": record.get("formula_shared")}
            route = {
                "id": f"{name}:J{row}", "sheet": name, "row": row, "department": department,
                "team": team, "line": line, "start": val(cells, f"D{row}"), "end": val(cells, f"E{row}"),
                "line_km": val(cells, f"F{row}"), "fleet": val(cells, f"G{row}"),
                "vehicles_a": val(cells, f"J{row}"), "a_days": val(cells, f"I{row}"),
                "trips_per_vehicle_a": val(cells, f"O{row}"), "max_regular": val(cells, f"H{row}"),
                "reserve_a": val(cells, f"N{row}"), "input_cell": f"J{row}",
                "plan_trips": {plan: val(summary, f"{_column(6+i)}{summary_row}") for i, plan in enumerate("ABCDEFG")},
                "daily": daily, "monthly_trips": _sum(d["trips"] for d in daily),
                "calendar_month_trips": _sum(d["trips"] for d in daily if d["valid"]),
                "summary_row": summary_row, "warnings": local_warnings,
                "cells": {key: detail(f"{column}{row}") for key, column in [("input", "J"), ("a_days", "I"), ("max_regular", "H"), ("reserve_a", "N"), ("trips_per_vehicle_a", "O")]},
            }
            routes.append(route)
        total_row = summary_totals.get(name)
        daily_totals = [val(summary, f"{_column(12+day)}{total_row}") for day in range(1, 32)] if total_row else []
        departments = []
        for department, summary_row in summary_departments.get(name, []):
            departments.append({"department": department, "summary_row": summary_row, "max_regular": val(summary, f"E{summary_row}"), "monthly_trips": _sum(val(summary, f"{_column(12+day)}{summary_row}") for day in range(1,32))})
        months.append({"name": name, "month": month, "year": year, "days": days, "dates": dates, "routes": routes, "warnings": month_warnings, "departments": departments,
            "totals": {"monthly_trips": _sum(daily_totals), "calendar_month_trips": _sum(daily_totals[:days]), "max_regular": val(summary, f"E{total_row}") if total_row else None, "summary_row": total_row, "daily": daily_totals, "basis": "原表全部31列逐日公式缓存的合计", "route_count": len(routes)}})
        patterns.append([(r["line"], r["vehicles_a"], r["trips_per_vehicle_a"], [(d["plan"]) for d in r["daily"]]) for r in routes])
    if patterns and all(p == patterns[0] for p in patterns[1:]):
        warnings.append(_warning("same_monthly_arrangements", "12 个月的 A 计划投放、单车日趟次及逐日计划代码相同；原始安排照常保留，请核实月度编制状态。", severity="warning"))
    cross_count = 0
    examples = []
    for address, record in summary.items():
        formula = record["formula"]
        if not formula:
            continue
        expected_month = val(summary, "D" + re.sub(r"[^0-9]", "", address))
        references = re.findall(r"'([0-9]+月)'!", formula)
        if expected_month and any(month != expected_month for month in references):
            cross_count += 1
            if len(examples) < 6:
                examples.append({"cell": address, "month": expected_month, "formula": formula})
    if cross_count:
        warnings.append(_warning("cross_month_references", f"汇总表至少 {cross_count} 个显式公式跨月引用（共享公式另计），保留原公式，请核实口径。", severity="warning", count=cross_count, examples=examples))
    error_cells = [{"sheet": name, "cell": address, "value": record["value"]} for name, cells in data.items() for address, record in cells.items() if record["type"] == "e"]
    if error_cells:
        warnings.append(_warning("formula_errors", f"工作簿含 {len(error_cells)} 个 Excel 错误结果，已保留。", severity="warning", cells=error_cells[:50], count=len(error_cells)))
    return {"filename": path.name, "sha256": hashlib.sha256(path.read_bytes()).hexdigest(), "sheets": list(sheets), "months": months, "warnings": warnings, "error_cells": error_cells, "formula_count": formula_count, "engine": engine_status(), "scope": "A计划营运1投放；原公式重算；公里口径待确认", "total_routes": sum(len(month["routes"]) for month in months)}


def _patch_cells(raw: bytes, changes: dict[str, int]) -> bytes:
    found = set()
    def patch(match):
        attrs, body = match.group("attrs"), match.group("body") or b""
        address_match = ATTR_R.search(attrs)
        address = address_match.group(1).decode() if address_match else None
        if address not in changes:
            return match.group(0)
        if FORMULA_RE.search(body):
            raise BudgetError(f"{address} 含公式，禁止覆盖。")
        found.add(address)
        body = VALUE_RE.sub(b"", body)
        attrs = ATTR_T.sub(b"", attrs)
        return b"<c" + attrs + b"><v>" + str(changes[address]).encode() + b"</v>" + body + b"</c>"
    output = CELL_RE.sub(patch, raw)
    if found != set(changes):
        raise BudgetError("找不到需要修改的输入格：" + ", ".join(set(changes)-found))
    return output


def _without_formula_caches(raw: bytes) -> bytes:
    """Remove old result caches from the temporary engine input only.

    LibreOffice can retain inter-sheet results on ordinary XLSX conversion even
    with fullCalcOnLoad enabled. Empty caches require it to calculate every
    original formula. The final output still starts from the original archive.
    """
    def strip(match):
        body = match.group("body") or b""
        if not FORMULA_RE.search(body):
            return match.group(0)
        return b"<c" + match.group("attrs") + b">" + VALUE_RE.sub(b"", body) + b"</c>"
    return CELL_RE.sub(strip, raw)


def _formula_cache(archive, member, strings):
    root = _xml(archive.read(member))
    result = {}
    for cell in root.findall("./" + TAG("sheetData") + "/" + TAG("row") + "/" + TAG("c")):
        if cell.find(TAG("f")) is None:
            continue
        result[cell.get("r")] = (_value(cell, strings), cell.get("t"))
    return result


def _merge_caches(raw, results):
    count = 0
    def merge(match):
        nonlocal count
        attrs, body = match.group("attrs"), match.group("body") or b""
        if not FORMULA_RE.search(body):
            return match.group(0)
        address = ATTR_R.search(attrs).group(1).decode()
        if address not in results:
            raise CalculationError(f"计算引擎遗漏公式结果 {address}，未生成新版本。")
        value, kind = results[address]
        if value is None:
            raise CalculationError(f"公式 {address} 没有可验证的缓存结果，未生成新版本。")
        count += 1
        attrs = ATTR_T.sub(b"", attrs)
        body = VALUE_RE.sub(b"", body)
        if kind == "e":
            if not re.search(rb'\bt="e"', match.group("attrs")):
                raise CalculationError(f"公式 {address} 产生新的 Excel 错误 {value}，未生成新版本。")
            attrs += b' t="e"'
            text = str(value)
        elif isinstance(value, bool):
            attrs += b' t="b"'
            text = "1" if value else "0"
        elif isinstance(value, str):
            attrs += b' t="str"'
            text = value
        else:
            text = str(value)
        escaped = text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").encode("utf-8")
        return b"<c" + attrs + b">" + body + b"<v>" + escaped + b"</v></c>"
    return CELL_RE.sub(merge, raw), count


def _formula_signature(raw):
    return [(ATTR_R.search(m.group("attrs")).group(1), FORMULA_RE.search(m.group("body") or b"").group(0)) for m in CELL_RE.finditer(raw) if FORMULA_RE.search(m.group("body") or b"")]


def recalculate_preserving_formulas(source_path, changes, output_path, timeout=120) -> dict:
    """Adjust only monthly A-plan 营运1 inputs, recalc, and create an immutable copy.

    changes: [{"sheet": "1月", "cell": "J7", "value": 14}]
    Raises on existing output paths, invalid inputs, timeout, or missing caches.
    """
    source, output = Path(source_path).resolve(), Path(output_path).resolve()
    if source == output or output.exists():
        raise BudgetError("原版本及已存在版本不可覆盖，请使用全新的输出文件名。")
    executable = spreadsheet_executable()
    if not executable:
        raise CalculationError("未安装 LibreOffice，无法执行保留原公式的真实重算。")
    if not isinstance(changes, list):
        raise BudgetError("调整参数必须是列表。")
    sheet_members, data, expected_count = _read(source)
    grouped = {}
    normalized = []
    seen = set()
    for change in changes:
        sheet, address, value = change.get("sheet"), change.get("cell"), change.get("value")
        if sheet not in data or not MONTH_RE.fullmatch(sheet) or not isinstance(address, str) or not re.fullmatch(r"J[0-9]+", address):
            raise BudgetError("本版仅允许修改指定月份 J 列的 A 计划营运1投放。")
        row = int(address[1:]); cells = data[sheet]
        line = cells.get(f"C{row}", {}).get("value")
        if row < 6 or not line or line == "合计":
            raise BudgetError("禁止修改标题、合计或非线路记录。")
        if isinstance(value, bool) or not isinstance(value, int) or value < 0:
            raise BudgetError("投放数必须为非负整数。")
        record = cells.get(address, {})
        if record.get("has_formula"):
            raise BudgetError("禁止修改或覆盖公式格。")
        if (sheet, address) in seen:
            raise BudgetError("同一输入格重复出现。")
        seen.add((sheet, address))
        grouped.setdefault(sheet, {})[address] = value
        normalized.append({"sheet": sheet, "cell": address, "before": record.get("value"), "value": value, "line": line})
    del data
    output.parent.mkdir(parents=True, exist_ok=True)
    started = time.monotonic()
    with tempfile.TemporaryDirectory(prefix="bus-budget-", dir="/tmp") as temporary:
        root = Path(temporary)
        engine_input = root / "input.xlsx"
        engine_output = root / "calculated"; engine_output.mkdir()
        profile = root / "profile"
        with ZipFile(source) as original, ZipFile(engine_input, "w") as working:
            for info in original.infolist():
                raw = original.read(info.filename)
                sheet = next((key for key, member in sheet_members.items() if member == info.filename), None)
                if sheet in grouped:
                    raw = _patch_cells(raw, grouped[sheet])
                if sheet:
                    raw = _without_formula_caches(raw)
                if info.filename == "xl/workbook.xml":
                    tree = _xml(raw)
                    calc = tree.find(TAG("calcPr"))
                    if calc is None:
                        calc = etree.SubElement(tree, TAG("calcPr"))
                    calc.set("fullCalcOnLoad", "1"); calc.set("forceFullCalc", "1"); calc.set("calcMode", "auto")
                    raw = etree.tostring(tree, xml_declaration=True, encoding="UTF-8")
                working.writestr(info, raw)
        env = os.environ.copy()
        env["XDG_CACHE_HOME"] = str(root / "cache")
        env["TMPDIR"] = str(root)
        command = [executable, "-env:UserInstallation=" + profile.as_uri(), "--headless", "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", str(engine_output), str(engine_input)]
        try:
            process = subprocess.run(command, capture_output=True, text=True, timeout=timeout, env=env)
        except subprocess.TimeoutExpired as exc:
            raise CalculationError(f"Excel 公式重算超过 {timeout} 秒，未生成新版本。") from exc
        except OSError as exc:
            raise CalculationError("无法启动 LibreOffice 计算引擎。") from exc
        recalculated = engine_output / "input.xlsx"
        if process.returncode != 0 or not recalculated.is_file():
            detail = (process.stderr + process.stdout)[-1000:]
            raise CalculationError("LibreOffice 未能生成重算文件：" + detail)
        merged = root / "merged.xlsx"
        merged_count = 0
        with ZipFile(source) as original, ZipFile(recalculated) as calculated, ZipFile(merged, "w") as final:
            calculated_members = _sheet_names(calculated)
            strings = _shared_strings(calculated)
            reverse = {member: name for name, member in sheet_members.items()}
            for info in original.infolist():
                raw = original.read(info.filename)
                sheet = reverse.get(info.filename)
                if sheet:
                    before_signature = _formula_signature(raw)
                    if sheet in grouped:
                        raw = _patch_cells(raw, grouped[sheet])
                    if before_signature:
                        if sheet not in calculated_members:
                            raise CalculationError(f"重算后缺少工作表 {sheet}。")
                        results = _formula_cache(calculated, calculated_members[sheet], strings)
                        raw, count = _merge_caches(raw, results)
                        merged_count += count
                    if _formula_signature(raw) != before_signature:
                        raise CalculationError("公式保护检查失败，未生成新版本。")
                final.writestr(info, raw)
        if merged_count != expected_count:
            raise CalculationError(f"公式数量不一致：{expected_count} → {merged_count}，未生成新版本。")
        # Exclusive creation prevents accidental overwrite, including racing requests.
        try:
            with output.open("xb") as target, merged.open("rb") as result:
                shutil.copyfileobj(result, target)
        except FileExistsError as exc:
            raise BudgetError("输出版本已存在，不可覆盖。") from exc
        except OSError:
            if output.exists():
                output.unlink()
            raise
    return {"output_path": str(output), "engine": "LibreOffice", "formula_count": merged_count, "formulas_preserved": True, "changes": normalized, "elapsed_seconds": round(time.monotonic() - started, 2)}


def compare_formula_caches(before_path, after_path) -> dict:
    """Compare cached results without deriving or evaluating any formula."""
    before_sheets, before, before_count = _read(before_path)
    after_sheets, after, after_count = _read(after_path)
    differences = []
    for sheet, cells in before.items():
        for address, record in cells.items():
            if not record["has_formula"]:
                continue
            new = after.get(sheet, {}).get(address)
            if new is None or new["value"] != record["value"]:
                differences.append({"sheet": sheet, "cell": address, "before": record["value"], "after": new["value"] if new else None})
    return {"identical": before_count == after_count and not differences, "before_formula_count": before_count, "after_formula_count": after_count, "difference_count": len(differences), "differences": differences}
