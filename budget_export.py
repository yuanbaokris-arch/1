"""Add a traceable, formula-based monthly report without resaving source sheets.

The existing worksheets are copied byte for byte. Only package registrations and
additional report styles are updated; original formulas and cached values stay
under the calculation engine's control.
"""
from __future__ import annotations

from collections import OrderedDict
from datetime import datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
import os
import posixpath
import re
import tempfile
import xml.etree.ElementTree as ET
from xml.sax.saxutils import escape, quoteattr
from zipfile import ZipFile
from zoneinfo import ZoneInfo

MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
DOC_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
NS = {"m": MAIN}
REPORT_NAME = "月趟次试算"


def _column(number: int) -> str:
    result = ""
    while number:
        number, digit = divmod(number - 1, 26)
        result = chr(65 + digit) + result
    return result


def _cell_text(cell: ET.Element | None, strings: list[str]) -> str:
    if cell is None:
        return ""
    if cell.get("t") == "inlineStr":
        return "".join(t.text or "" for t in cell.findall("m:is//m:t", NS))
    value = cell.findtext("m:v", "", NS)
    if cell.get("t") == "s":
        return strings[int(value)] if value else ""
    return value


def _package_target(target: str) -> str:
    return posixpath.normpath(target.lstrip("/") if target.startswith("/") else "xl/" + target)


def _read_summary(package: ZipFile) -> tuple[list[tuple[str, str, str]], dict]:
    workbook = ET.fromstring(package.read("xl/workbook.xml"))
    relations = ET.fromstring(package.read("xl/_rels/workbook.xml.rels"))
    targets = {r.get("Id"): r.get("Target") for r in relations}
    sheets = workbook.find("m:sheets", NS)
    summary = next((s for s in sheets if s.get("name") == "汇总"), None)
    if summary is None:
        raise ValueError("未找到原表的‘汇总’工作表，不能生成月趟次试算。")
    target = targets.get(summary.get("{" + DOC_REL + "}id"))
    if not target:
        raise ValueError("‘汇总’工作表的关联信息缺失。")
    strings = []
    if "xl/sharedStrings.xml" in package.namelist():
        strings = ["".join(t.text or "" for t in item.findall(".//m:t", NS))
                   for item in ET.fromstring(package.read("xl/sharedStrings.xml"))]
    root = ET.fromstring(package.read(_package_target(target)))
    months = {month: OrderedDict() for month in range(1, 13)}
    for row in root.findall("m:sheetData/m:row", NS):
        cells = {re.sub(r"\d+$", "", c.get("r", "")): c for c in row}
        department, fleet, route, month_label = (_cell_text(cells.get(c), strings) for c in "ABCD")
        month_match = re.fullmatch(r"(1[0-2]|[1-9])月", month_label.strip())
        if not month_match or not route or route.strip() in {"合计", "小计", "总计"}:
            continue
        # All four source metadata fields identify a record; missing identifiers
        # are rejected rather than inferring them from adjacent rows.
        if not department or not fleet:
            raise ValueError(f"汇总第 {row.get('r')} 行的业务部或车队为空，无法确定线路。")
        key = (department, fleet, route)
        month = int(month_match.group(1))
        if key in months[month]:
            raise ValueError(f"汇总中 {month} 月线路 {key} 重复，无法唯一定位。")
        months[month][key] = (int(row.get("r")), cells)
    keys = list(months[1])
    if not keys:
        raise ValueError("汇总中没有可识别的 1 月线路明细。")
    for month in range(1, 13):
        if set(months[month]) != set(keys):
            raise ValueError(f"汇总中 {month} 月与 1 月线路清单不一致，不能猜测对应关系。")
    return keys, months


def _sum_cached(cells: dict[str, ET.Element]) -> tuple[str | None, bool]:
    """Apply Excel SUM's cache rules only to existing M:AQ values.

    An uncached formula remains unknown. Excel error values propagate. Plain
    blanks and text are ignored, as they are by a SUM(range) formula.
    """
    total = Decimal(0)
    for col in range(13, 44):
        cell = cells.get(_column(col))
        if cell is None:
            continue
        value = cell.findtext("m:v", "", NS)
        if cell.get("t") == "e":
            return value or "#VALUE!", True
        is_formula = cell.find("m:f", NS) is not None
        # t="str" plus an explicitly present empty <v/> is a cached empty
        # string (e.g. IF(...,"")), which SUM ignores. It is not an absent
        # numeric cache. Missing formula caches remain unknown.
        known_empty_string = cell.get("t") == "str" and cell.find("m:v", NS) is not None
        if is_formula and not value and not known_empty_string:
            return None, False
        if not value or cell.get("t") in {"s", "str", "inlineStr", "b"}:
            continue
        try:
            total += Decimal(value)
        except InvalidOperation:
            return None, False
    return _number(total), False


def _number(value: Decimal) -> str:
    return format(value, "f")


def _inline(ref: str, text: str, style: int) -> str:
    # inlineStr prevents a user supplied reason or version from becoming a formula.
    return f'<c r="{ref}" s="{style}" t="inlineStr"><is><t xml:space="preserve">{escape(str(text))}</t></is></c>'


def _formula(ref: str, formula: str, cached: str | None, is_error: bool, style: int) -> str:
    error_type = ' t="e"' if is_error else ""
    value = f"<v>{escape(cached)}</v>" if cached is not None else ""
    return f'<c r="{ref}" s="{style}"{error_type}><f>{escape(formula)}</f>{value}</c>'


def _append_xml_collection(raw: bytes, tag: str, addition: str, old_count: int, extra: int) -> bytes:
    """Append styles while leaving each existing style's XML untouched."""
    pattern = rb"(<" + tag.encode() + rb"\b[^>]*>)(.*?)(</" + tag.encode() + rb">)"
    def replace(match):
        opening, count = re.subn(rb'\bcount="\d+"', f'count="{old_count + extra}"'.encode(), match[1], count=1)
        if count != 1:
            raise ValueError(f"样式集合 {tag} 缺少可识别的数量属性。")
        return opening + match[2] + addition.encode("utf-8") + match[3]
    updated, count = re.subn(pattern, replace, raw, flags=re.DOTALL)
    if count != 1:
        raise ValueError(f"不能安全追加 {tag} 样式；原样式未修改。")
    return updated


def _report_styles(raw: bytes) -> tuple[bytes, dict[str, int]]:
    root = ET.fromstring(raw)
    counts = {name: len(root.find("m:" + name, NS)) for name in ("fonts", "fills", "cellXfs")}
    font = counts["fonts"]
    fill = counts["fills"]
    first_style = counts["cellXfs"]
    fonts = ('<font><sz val="11"/><color rgb="FF17243A"/><name val="微软雅黑"/><charset val="134"/></font>'
             '<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="微软雅黑"/><charset val="134"/></font>'
             '<font><b/><sz val="17"/><color rgb="FF17243A"/><name val="微软雅黑"/><charset val="134"/></font>'
             '<font><sz val="10"/><color rgb="FF9A5700"/><name val="微软雅黑"/><charset val="134"/></font>')
    fills = ('<fill><patternFill patternType="solid"><fgColor rgb="FF233C5B"/><bgColor indexed="64"/></patternFill></fill>'
             '<fill><patternFill patternType="solid"><fgColor rgb="FFEAF0F6"/><bgColor indexed="64"/></patternFill></fill>')
    def xf(font_id, fill_id=0, numfmt=0, alignment="left", wrap=False):
        return (f'<xf numFmtId="{numfmt}" fontId="{font_id}" fillId="{fill_id}" borderId="0" xfId="0" '
                'applyFont="1" applyFill="1" applyAlignment="1" applyNumberFormat="1">'
                f'<alignment horizontal="{alignment}" vertical="center"' + (' wrapText="1"' if wrap else '') + '/></xf>')
    specs = OrderedDict([
        ("body", xf(font)),
        ("wrapped_body", xf(font, wrap=True)),
        ("number", xf(font, numfmt=2, alignment="right")),
        ("header", xf(font + 1, fill, alignment="center", wrap=True)),
        ("title", xf(font + 2)),
        ("note", xf(font + 3, wrap=True)),
        ("total_label", xf(font, fill + 1)),
        ("total_number", xf(font, fill + 1, numfmt=2, alignment="right")),
    ])
    raw = _append_xml_collection(raw, "fonts", fonts, counts["fonts"], 4)
    raw = _append_xml_collection(raw, "fills", fills, counts["fills"], 2)
    raw = _append_xml_collection(raw, "cellXfs", "".join(specs.values()), first_style, len(specs))
    return raw, {name: first_style + i for i, name in enumerate(specs)}


def _report_xml(keys, months, styles, version_label, version_id, reason, source_sha256) -> bytes:
    rows = []
    generated = datetime.now(ZoneInfo("Asia/Shanghai")).strftime("%Y-%m-%d %H:%M:%S +08:00")
    notes = [
        (1, f"公交营运月趟次试算 · {version_label}", "title", 32),
        (2, f"方案版本：{version_label}    版本标识：{version_id}    导出时间：{generated}", "body", 28),
        (3, "口径：按原表逐日 31 列合计；2 月等跨月日期尚待确认。本表为试算清单，非已确认报送模板。", "note", 32),
        (4, f"调整原因：{reason or '未填写'}", "wrapped_body", max(30, min(150, 16 * (1 + len(reason) // 110)))),
        (5, f"初始来源 SHA-256：{source_sha256}", "body", 28),
        (6, "公式说明：各月直接 SUM 原“汇总”表 M:AQ；全年为 12 个月合计。原业务公式、原有日期与工作表保持不变；缓存来自导出版本，请用 Excel 完整重算核对。", "note", 36),
    ]
    for row, text, style, height in notes:
        rows.append(f'<row r="{row}" ht="{height}" customHeight="1">{_inline(f"A{row}", text, styles[style])}</row>')
    headers = ["所属业务部", "车队", "线路"] + [f"{m}月趟次" for m in range(1, 13)] + ["全年合计"]
    rows.append('<row r="8" ht="30" customHeight="1">' + ''.join(_inline(f'{_column(i)}8', h, styles['header']) for i, h in enumerate(headers, 1)) + '</row>')
    column_caches = [[] for _ in range(13)]
    for index, key in enumerate(keys, 9):
        cells = [_inline(f"{_column(i)}{index}", value, styles["body"]) for i, value in enumerate(key, 1)]
        cached_months = []
        for month in range(1, 13):
            source_row, source_cells = months[month][key]
            cached, error = _sum_cached(source_cells)
            cached_months.append((cached, error))
            cells.append(_formula(f"{_column(month + 3)}{index}", f"SUM('汇总'!M{source_row}:AQ{source_row})", cached, error, styles["number"]))
        annual, annual_error = _sum_values(cached_months)
        cells.append(_formula(f"P{index}", f"SUM(D{index}:O{index})", annual, annual_error, styles["number"]))
        for col_index, value in enumerate(cached_months + [(annual, annual_error)]):
            column_caches[col_index].append(value)
        rows.append(f'<row r="{index}" ht="23" customHeight="1">' + ''.join(cells) + '</row>')
    last_data = len(keys) + 8
    total_row = last_data + 1
    totals = [_inline(f"A{total_row}", "全司合计", styles["total_label"])]
    for col, values in enumerate(column_caches, 4):
        cached, error = _sum_values(values)
        label = _column(col)
        totals.append(_formula(f"{label}{total_row}", f"SUM({label}9:{label}{last_data})", cached, error, styles["total_number"]))
    rows.append(f'<row r="{total_row}" ht="27" customHeight="1">' + ''.join(totals) + '</row>')
    merges = ''.join(f'<mergeCell ref="A{r}:P{r}"/>' for r in range(1, 7)) + f'<mergeCell ref="A{total_row}:C{total_row}"/>'
    xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           f'<worksheet xmlns="{MAIN}"><dimension ref="A1:P{total_row}"/>'
           '<sheetViews><sheetView workbookViewId="0" showGridLines="0"><pane xSplit="3" ySplit="8" topLeftCell="D9" activePane="bottomRight" state="frozen"/></sheetView></sheetViews>'
           '<sheetFormatPr defaultRowHeight="23"/><cols><col min="1" max="1" width="17" customWidth="1"/>'
           '<col min="2" max="2" width="14" customWidth="1"/><col min="3" max="3" width="24" customWidth="1"/>'
           '<col min="4" max="16" width="15" customWidth="1"/></cols><sheetData>' + ''.join(rows) + '</sheetData>'
           f'<autoFilter ref="A8:P{last_data}"/><mergeCells count="7">{merges}</mergeCells>'
           '<printOptions horizontalCentered="1"/><pageMargins left="0.25" right="0.25" top="0.5" bottom="0.5" header="0.2" footer="0.2"/>'
           '<pageSetup paperSize="9" orientation="landscape" fitToWidth="1" fitToHeight="0"/></worksheet>')
    return xml.encode("utf-8")


def _sum_values(values: list[tuple[str | None, bool]]) -> tuple[str | None, bool]:
    # Do not turn unknown source caches into apparent zero results.
    error = next((value for value, is_error in values if is_error), None)
    if error is not None:
        return error, True
    if any(value is None for value, _ in values):
        return None, False
    return _number(sum((Decimal(value) for value, _ in values), Decimal(0))), False


def _insert_before(raw: bytes, closing: bytes, content: str) -> bytes:
    if raw.count(closing) != 1:
        raise ValueError("工作簿结构不符合安全追加条件，未修改原文件。")
    return raw.replace(closing, content.encode("utf-8") + closing, 1)


def add_monthly_report(source_path, output_path, version_label, version_id, reason, source_sha256):
    """Return output_path after appending a report to a *new* XLSX package.

    source_path must already contain the selected version's inputs and formula
    caches. No original worksheet XML is parsed and rewritten. Existing reports
    are refused, so this operation cannot replace an earlier report silently.
    """
    source = Path(source_path)
    output = Path(output_path)
    if source.resolve() == output.resolve():
        raise ValueError("导出必须使用新文件，不能覆盖来源工作簿。")
    if output.exists():
        raise ValueError("导出目标已存在，请使用新文件名，避免覆盖已有版本。")
    if not re.fullmatch(r"[a-fA-F0-9]{64}", str(source_sha256)):
        raise ValueError("来源 SHA-256 必须是完整的 64 位摘要。")
    with ZipFile(source) as package:
        if package.testzip() is not None:
            raise ValueError("来源工作簿 ZIP 校验失败，不能导出。")
        original_names = package.namelist()
        if len(set(original_names)) != len(original_names):
            raise ValueError("来源工作簿存在重复 ZIP 项，不能安全导出。")
        workbook_bytes = package.read("xl/workbook.xml")
        workbook = ET.fromstring(workbook_bytes)
        sheets = workbook.find("m:sheets", NS)
        if any(s.get("name") == REPORT_NAME for s in sheets):
            raise ValueError("来源中已存在‘月趟次试算’，请从保存的方案版本重新导出。")
        keys, months = _read_summary(package)
        rel_bytes = package.read("xl/_rels/workbook.xml.rels")
        relations = ET.fromstring(rel_bytes)
        existing_ids = {r.get("Id") for r in relations}
        next_id = 1
        while f"rId{next_id}" in existing_ids:
            next_id += 1
        relationship_id = f"rId{next_id}"
        sheet_id = max(int(s.get("sheetId")) for s in sheets) + 1
        path_number = 1
        while f"xl/worksheets/sheet{path_number}.xml" in original_names:
            path_number += 1
        sheet_path = f"xl/worksheets/sheet{path_number}.xml"
        styles, style_ids = _report_styles(package.read("xl/styles.xml"))
        report = _report_xml(keys, months, style_ids, str(version_label), str(version_id), str(reason or ""), str(source_sha256))
        changed = {
            "xl/styles.xml": styles,
            "xl/workbook.xml": _insert_before(workbook_bytes, b"</sheets>", f'<sheet name={quoteattr(REPORT_NAME)} sheetId="{sheet_id}" r:id="{relationship_id}"/>'),
            "xl/_rels/workbook.xml.rels": _insert_before(rel_bytes, b"</Relationships>", f'<Relationship Id="{relationship_id}" Type="{DOC_REL}/worksheet" Target="worksheets/sheet{path_number}.xml"/>'),
            "[Content_Types].xml": _insert_before(package.read("[Content_Types].xml"), b"</Types>", f'<Override PartName="/{sheet_path}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'),
        }
        output.parent.mkdir(parents=True, exist_ok=True)
        temporary_path = None
        try:
            with tempfile.NamedTemporaryFile(prefix=".budget-report-", suffix=".xlsx", dir=output.parent, delete=False) as temporary:
                temporary_path = Path(temporary.name)
            with ZipFile(temporary_path, "w") as result:
                for info in package.infolist():
                    result.writestr(info, changed.get(info.filename, package.read(info.filename)))
                result.writestr(sheet_path, report)
            with ZipFile(temporary_path) as result:
                if result.testzip() is not None:
                    raise ValueError("生成文件 ZIP 校验失败。")
            # Exclusive creation avoids replacing a file created since the check.
            os.link(temporary_path, output)
        finally:
            if temporary_path is not None:
                temporary_path.unlink(missing_ok=True)
    return str(output)
