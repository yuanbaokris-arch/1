"""Report package integrity and formula preservation acceptance checks."""
from pathlib import Path
import tempfile
import unittest
import xml.etree.ElementTree as ET
from zipfile import ZIP_DEFLATED, ZipFile

from budget_export import MAIN, DOC_REL, NS, add_monthly_report


def make_workbook(path, *, missing_month=False, missing_cache=False):
    rows = []
    for month in range(1, 13):
        if missing_month and month == 7:
            continue
        for fleet, offset in [('1车队', 7), ('2车队', 43)]:
            row = month * 100 + offset
            text_cells = ''.join(f'<c r="{c}{row}" t="inlineStr"><is><t>{v}</t></is></c>'
                                 for c, v in zip('ABCD', ['一业务部', fleet, '136路', f'{month}月']))
            daily_cells = ''
            for column in range(13, 44):
                from budget_export import _column
                cache = '' if missing_cache and month == 1 and fleet == '1车队' and column == 13 else f'<v>{month}</v>'
                daily_cells += f'<c r="{_column(column)}{row}"><f>{month}+0</f>{cache}</c>'
            rows.append(f'<row r="{row}">{text_cells}{daily_cells}</row>')
        row = month * 100 + 90
        rows.append(f'<row r="{row}"><c r="C{row}" t="inlineStr"><is><t>合计</t></is></c>'
                    f'<c r="D{row}" t="inlineStr"><is><t>{month}月</t></is></c></row>')
    parts = {
        '[Content_Types].xml': '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"></Types>',
        'xl/workbook.xml': f'<workbook xmlns="{MAIN}" xmlns:r="{DOC_REL}"><sheets><sheet name="汇总" sheetId="14" r:id="rId5"/></sheets></workbook>',
        'xl/_rels/workbook.xml.rels': f'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId5" Type="{DOC_REL}/worksheet" Target="worksheets/sheet7.xml"/></Relationships>',
        'xl/styles.xml': f'<styleSheet xmlns="{MAIN}"><fonts count="1"><font><sz val="11"/></font></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf/></cellStyleXfs><cellXfs count="1"><xf/></cellXfs></styleSheet>',
        'xl/worksheets/sheet7.xml': f'<worksheet xmlns="{MAIN}"><sheetData>{"".join(rows)}</sheetData></worksheet>',
        'customXml/untouched.xml': '<custom>keep all source package parts</custom>',
    }
    with ZipFile(path, 'w', ZIP_DEFLATED) as z:
        for name, data in parts.items():
            z.writestr(name, data)


class MonthlyReportTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.source = Path(self.temp.name) / 'source.xlsx'
        self.output = Path(self.temp.name) / 'report.xlsx'
        make_workbook(self.source)

    def export(self, **kwargs):
        return add_monthly_report(self.source, self.output, 'V1', 'immutable-001',
                                  '=HYPERLINK("https://example.invalid") & 核验', 'a' * 64, **kwargs)

    def test_report_preserves_original_worksheets_and_formula_xml(self):
        self.assertEqual(self.export(), str(self.output))
        with ZipFile(self.source) as source, ZipFile(self.output) as output:
            self.assertIsNone(output.testzip())
            for name in source.namelist():
                if name not in {'xl/styles.xml', 'xl/workbook.xml', 'xl/_rels/workbook.xml.rels', '[Content_Types].xml'}:
                    self.assertEqual(source.read(name), output.read(name), name)
            report = ET.fromstring(output.read('xl/worksheets/sheet1.xml'))
            original_formulas = ET.fromstring(source.read('xl/worksheets/sheet7.xml')).findall('.//m:f', NS)
            all_formulas = sum(len(ET.fromstring(output.read(name)).findall('.//m:f', NS))
                               for name in output.namelist() if name.startswith('xl/worksheets/') and name.endswith('.xml'))
            self.assertEqual(all_formulas - len(original_formulas), 2 * 13 + 13)
            cells = {c.get('r'): c for c in report.findall('.//m:c', NS)}
            self.assertEqual(cells['D9'].findtext('m:f', namespaces=NS), "SUM('汇总'!M107:AQ107)")
            self.assertEqual(cells['E9'].findtext('m:f', namespaces=NS), "SUM('汇总'!M207:AQ207)")
            self.assertEqual(cells['D9'].findtext('m:v', namespaces=NS), '31')
            self.assertEqual(cells['P9'].findtext('m:v', namespaces=NS), '2418')
            self.assertEqual(cells['A4'].get('t'), 'inlineStr')
            self.assertIsNone(cells['A4'].find('m:f', NS))
            text = ''.join(report.itertext())
            for value in ['V1', 'immutable-001', 'a' * 64, '跨月日期尚待确认', '非已确认报送模板']:
                self.assertIn(value, text)
            # Existing style definitions remain intact even though report styles
            # are appended to the same collections.
            before = ET.fromstring(source.read('xl/styles.xml'))
            after = ET.fromstring(output.read('xl/styles.xml'))
            for tag in ('fonts', 'fills', 'cellXfs'):
                for i, child in enumerate(before.find('m:' + tag, NS)):
                    self.assertEqual(ET.tostring(child), ET.tostring(after.find('m:' + tag, NS)[i]))

    def test_missing_caches_stay_unknown(self):
        make_workbook(self.source, missing_cache=True)
        self.export()
        with ZipFile(self.output) as z:
            cells = {c.get('r'): c for c in ET.fromstring(z.read('xl/worksheets/sheet1.xml')).findall('.//m:c', NS)}
            self.assertIsNone(cells['D9'].find('m:v', NS))
            self.assertIsNone(cells['P9'].find('m:v', NS))
            self.assertIsNone(cells['D11'].find('m:v', NS))

    def test_cached_empty_string_is_ignored_by_sum(self):
        from budget_export import _sum_cached
        formula_empty = ET.fromstring(f'<c xmlns="{MAIN}" t="str"><f>IF(1=0,1,"")</f><v/></c>')
        numeric = ET.fromstring(f'<c xmlns="{MAIN}"><v>4.5</v></c>')
        self.assertEqual(_sum_cached({"M": formula_empty, "N": numeric}), ('4.5', False))

    def test_missing_month_is_rejected_without_guessing(self):
        make_workbook(self.source, missing_month=True)
        with self.assertRaisesRegex(ValueError, '7 月'):
            self.export()
        self.assertFalse(self.output.exists())

    def test_refuses_overwrite_and_repeat_report(self):
        with self.assertRaisesRegex(ValueError, '不能覆盖'):
            add_monthly_report(self.source, self.source, 'V1', 'id', '', 'a' * 64)
        self.export()
        original_output = self.output.read_bytes()
        with self.assertRaisesRegex(ValueError, '已存在'):
            self.export()
        self.assertEqual(self.output.read_bytes(), original_output)
        with self.assertRaisesRegex(ValueError, '已存在'):
            add_monthly_report(self.output, self.output.parent / 'repeat.xlsx', 'V2', 'id', '', 'a' * 64)


if __name__ == '__main__':
    unittest.main()
