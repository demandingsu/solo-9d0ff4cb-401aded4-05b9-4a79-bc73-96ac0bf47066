"""端到端自测。

重点场景：园区总计中本楼栋同一读数先被计入、又被扣除（跨表汇总），
验证导出的试算表在修改读数后与「原表做同样修改后重算」结果一致。

测试用最小的“整数求和模型”独立模拟原表重算，不依赖被测代码的求值逻辑。
"""
from __future__ import annotations

import io
import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from openpyxl import Workbook, load_workbook  # noqa: E402

from app.analyzer import FormulaError, WorkbookAnalyzer, parse_total_target  # noqa: E402
from app.builder import SUMMARY_SHEET, build_workbook, save_workbook  # noqa: E402
from app.formula import parse_formula  # noqa: E402


def _roundtrip(wb) -> Workbook:
    return load_workbook(io.BytesIO(save_workbook(wb)), data_only=False)


def _tokenize_simple(body: str):
    """独立词法：NUMBER / REF / SUM( / 运算符 / 括号 / 逗号。"""
    tokens = []
    i, n = 0, len(body)
    while i < n:
        ch = body[i]
        if ch.isspace():
            i += 1
        elif ch in "+-(),":
            tokens.append((ch, ch))
            i += 1
        elif ch.isdigit():
            j = i + 1
            while j < n and body[j].isdigit():
                j += 1
            tokens.append(("NUMBER", body[i:j]))
            i = j
        elif body.startswith("SUM(", i):
            tokens.append(("SUM", "SUM"))
            tokens.append(("(", "("))
            i += 4
        else:
            # 引用：可选 '工作表'! 前缀 + 单元格/区域
            start = i
            if ch == "'":
                j = i + 1
                while j < n and body[j] != "'":
                    j += 1
                j = body.index("!", j) + 1
            else:
                j = i
                while j < n and body[j] not in "+-(), ":
                    j += 1
            # 继续吃到运算符/括号/空白为止（覆盖无引号工作表名与区域）
            while j < n and body[j] not in "+-(), ":
                j += 1
            tokens.append(("REF", body[start:j]))
            i = j
    return tokens


def _eval_simple(
    formula: str,
    values: dict[tuple[str, str], int],
    constants: dict[tuple[str, str], int] | None = None,
) -> int:
    """独立的极简求和求值器：只支持一元 +/-、SUM(...)、引用与整数。

    与生产代码分开实现（含独立词法），用于交叉核对预期结果。
    constants 提供汇总表上已经折算好的常数单元格。
    """
    constants = constants or {}
    tokens = _tokenize_simple(formula[1:])
    pos = 0

    def parse_expr() -> int:
        nonlocal pos
        total = parse_term()
        while pos < len(tokens) and tokens[pos][0] in ("+", "-"):
            op = tokens[pos][0]
            pos += 1
            rhs = parse_term()
            total = total + rhs if op == "+" else total - rhs
        return total

    def parse_term() -> int:
        nonlocal pos
        kind, text = tokens[pos]
        if kind in ("+", "-"):
            pos += 1
            value = parse_term()
            return value if kind == "+" else -value
        pos += 1
        if kind == "NUMBER":
            return int(text)
        if kind == "REF":
            return _sum_ref(text, values, constants)
        if kind == "SUM":
            assert tokens[pos][0] == "("
            pos += 1
            total = parse_expr()
            while tokens[pos][0] == ",":
                pos += 1
                total += parse_expr()
            assert tokens[pos][0] == ")"
            pos += 1
            return total
        if kind == "(":
            value = parse_expr()
            assert tokens[pos][0] == ")"
            pos += 1
            return value
        raise AssertionError(f"测试求值器无法处理记号: {(kind, text)}")

    return parse_expr()


def _sum_ref(ref: str, values: dict[tuple[str, str], int],
             constants: dict[tuple[str, str], int]) -> int:
    if "!" in ref:
        sheet, coord = ref.split("!", 1)
        sheet = sheet.strip("'")
    else:
        raise AssertionError("测试中引用都应带工作表名")
    coord = coord.replace("$", "")
    if ":" not in coord:
        key = (sheet, coord)
        if key in constants:
            return constants[key]
        if sheet == SUMMARY_SHEET and coord == "B1":
            # 试算汇总!B1 是总计本身，测试中不允许再嵌套引用
            raise AssertionError("总计公式不应自引用")
        return values[key]
    from openpyxl.utils import range_boundaries, get_column_letter

    c1, r1, c2, r2 = range_boundaries(coord)
    total = 0
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            total += values[(sheet, f"{get_column_letter(c)}{r}")]
    return total


class CrossSheetNettingTest(unittest.TestCase):
    """跨表汇总：同一读数先计入再扣除。"""

    def setUp(self):
        wb = Workbook()
        # 1 号楼（本楼栋）：读数 + 表内小计公式
        s1 = wb.active
        s1.title = "1号楼"
        s1["A1"] = "电表"
        s1["B1"] = "读数"
        s1["A2"] = "电梯"
        s1["B2"] = 100
        s1["A3"] = "水泵"
        s1["B3"] = 30
        s1["A4"] = "楼内小计"
        s1["B4"] = "=SUM(B2:B3)"

        # 2 号楼（外部底数）
        s2 = wb.create_sheet("2号楼")
        s2["A1"] = "电表"
        s2["B1"] = "读数"
        s2["A2"] = "电梯"
        s2["B2"] = 200
        s2["A3"] = "水泵"
        s2["B3"] = 80
        s2["A4"] = "楼内小计"
        s2["B4"] = "=B2+B3"

        # 园区汇总：
        #   先把 1 号楼小计计入，再扣除 1 号楼水泵读数（同一读数先计入再扣除的净额场景）
        #   再加上 2 号楼小计
        sm = wb.create_sheet("园区汇总")
        sm["A1"] = "项目"
        sm["B1"] = "数值"
        sm["A2"] = "含重复计入的园区电量"
        sm["B2"] = "='1号楼'!B4+'2号楼'!B4"
        sm["A3"] = "应扣除的1号楼水泵"
        sm["B3"] = "='1号楼'!B3"
        sm["A4"] = "园区总计"
        sm["B4"] = "=B2-B3"

        self.wb = wb
        # 原表“真值”模型（独立计算）
        self.values = {
            ("1号楼", "B2"): 100,
            ("1号楼", "B3"): 30,
            ("2号楼", "B2"): 200,
            ("2号楼", "B3"): 80,
        }

    def _original_total(self, overrides: dict[tuple[str, str], int]) -> int:
        values = dict(self.values)
        values.update(overrides)
        s1_sub = values[("1号楼", "B2")] + values[("1号楼", "B3")]
        s2_sub = values[("2号楼", "B2")] + values[("2号楼", "B3")]
        gross = s1_sub + s2_sub
        deduct = values[("1号楼", "B3")]
        return gross - deduct

    def test_analyze_keeps_netted_reading_with_opposite_signs(self):
        analyzer = WorkbookAnalyzer(self.wb, "1号楼", "园区汇总", "B4")
        result = analyzer.analyze()

        # 外部常数 = 2 号楼小计 280
        self.assertEqual(result.external_constant, 280)
        refs = [(t.sign, t.ref.sheet, t.ref.coord, t.ref.is_range)
                for t in result.building_terms]
        # B4=SUM(B2:B3) 以 +区域 保留，总计再以单点 -B3 扣除：
        # B3 被“先计入（区域内）再扣除（单点）”，两个 B3 净额为零
        self.assertIn((1, "1号楼", "B2:B3", True), refs)
        self.assertIn((-1, "1号楼", "B3", False), refs)
        net_b3 = sum(
            t.sign for t in result.building_terms
            if not t.ref.is_range and t.ref.coord == "B3"
        )
        self.assertEqual(net_b3, -1)
        # 区域内的 B3（+1）与单点 -B3 相抵，B3 对总计的净系数为 0
        self.assertEqual(1 + net_b3, 0)

    def test_exported_total_matches_original_recalc_after_edit(self):
        analyzer = WorkbookAnalyzer(self.wb, "1号楼", "园区汇总", "B4")
        result = analyzer.analyze()
        out = _roundtrip(build_workbook(self.wb, result))

        self.assertEqual(out.sheetnames, ["1号楼", SUMMARY_SHEET])
        # 本楼栋读数与公式保留
        s1 = out["1号楼"]
        self.assertEqual(s1["B2"].value, 100)
        self.assertEqual(s1["B3"].value, 30)
        self.assertEqual(s1["B4"].value, "=SUM('1号楼'!B2:B3)")

        summary = out[SUMMARY_SHEET]
        self.assertEqual(summary["A1"].value, 280)
        total_formula = summary["B1"].value
        self.assertTrue(total_formula.startswith("="))
        # 公式中只允许 +、-、SUM、同簿引用
        body = total_formula[1:]
        for token in ("*", "/", "&", "VLOOKUP", "IF(", "["):
            self.assertNotIn(token, body)

        # 修改本楼栋读数：B2 100 -> 150，B3 30 -> 55
        s1["B2"] = 150
        s1["B3"] = 55
        out_values = {
            ("1号楼", "B2"): 150,
            ("1号楼", "B3"): 55,
            ("2号楼", "B2"): 200,
            ("2号楼", "B3"): 80,
        }
        trial_total = _eval_simple(
            total_formula, out_values,
            constants={(SUMMARY_SHEET, "A1"): summary["A1"].value},
        )
        expected = self._original_total({("1号楼", "B2"): 150, ("1号楼", "B3"): 55})
        self.assertEqual(trial_total, expected)
        # 手工核对：(150+55)+(200+80)-55 = 430
        self.assertEqual(trial_total, 430)

        # 再改一次，验证“先计入再扣除”的 B3 净额为零、不影响总计
        s1["B3"] = 999
        out_values[("1号楼", "B3")] = 999
        trial_total_2 = _eval_simple(
            total_formula, out_values,
            constants={(SUMMARY_SHEET, "A1"): summary["A1"].value},
        )
        # 同前提对照：原表在 B2=150、B3=999 时的重算结果
        expected_2 = self._original_total({("1号楼", "B2"): 150, ("1号楼", "B3"): 999})
        self.assertEqual(trial_total_2, expected_2)
        # B3 被净额抵消：280（外部常数）+ B2=150 = 430，与上一轮相同
        self.assertEqual(trial_total_2, 430)
        self.assertEqual(trial_total_2, 280 + 150)

    def test_baseline_matches_original_total(self):
        analyzer = WorkbookAnalyzer(self.wb, "1号楼", "园区汇总", "B4")
        result = analyzer.analyze()
        out = _roundtrip(build_workbook(self.wb, result))
        base_values = dict(self.values)
        baseline = _eval_simple(
            out[SUMMARY_SHEET]["B1"].value, base_values,
            constants={(SUMMARY_SHEET, "A1"): out[SUMMARY_SHEET]["A1"].value},
        )
        self.assertEqual(baseline, self._original_total({}))
        self.assertEqual(baseline, 380)  # (100+30)+(200+80)-30


class RangeAndUnsupportedTest(unittest.TestCase):
    def _wb_with_external_ranges(self):
        wb = Workbook()
        s1 = wb.active
        s1.title = "1号楼"
        s1["B2"] = 10
        s1["B3"] = 20
        other = wb.create_sheet("2号楼")
        for i, v in enumerate([7, 11, 13], start=2):
            other[f"B{i}"] = v
        other["C2"] = "=B2*2"  # 不支持的乘法
        sm = wb.create_sheet("园区汇总")
        sm["B2"] = "=SUM('1号楼'!B2:B3)+SUM('2号楼'!B2:B4)"
        return wb

    def test_external_range_folded_to_single_constant(self):
        wb = self._wb_with_external_ranges()
        result = WorkbookAnalyzer(wb, "1号楼", "园区汇总", "B2").analyze()
        self.assertEqual(result.external_constant, 7 + 11 + 13)
        self.assertEqual(len(result.building_terms), 1)
        self.assertTrue(result.building_terms[0].ref.is_range)

        out = _roundtrip(build_workbook(wb, result))
        self.assertNotIn("2号楼", out.sheetnames)
        formula = out[SUMMARY_SHEET]["B1"].value
        # 外部底数只以一个常数入表，不出现 2 号楼引用
        self.assertNotIn("2号楼", formula)
        values = {("1号楼", "B2"): 10, ("1号楼", "B3"): 20}
        self.assertEqual(
            _eval_simple(formula, values,
                         constants={(SUMMARY_SHEET, "A1"): 31}),
            31 + 30,
        )

    def test_unsupported_operator_reports_cell_and_stops(self):
        wb = self._wb_with_external_ranges()
        # 让总计依赖到含乘法的单元格
        wb["园区汇总"]["B2"] = "='2号楼'!C2+'1号楼'!B2"
        with self.assertRaises(FormulaError) as ctx:
            WorkbookAnalyzer(wb, "1号楼", "园区汇总", "B2").analyze()
        self.assertIn("2号楼!C2", str(ctx.exception))
        self.assertIn("不支持", str(ctx.exception))

    def test_unsupported_function(self):
        wb = self._wb_with_external_ranges()
        wb["园区汇总"]["B2"] = "=IF('1号楼'!B2>0,1,0)"
        with self.assertRaises(FormulaError) as ctx:
            WorkbookAnalyzer(wb, "1号楼", "园区汇总", "B2").analyze()
        self.assertIn("园区汇总!B2", str(ctx.exception))

    def test_cycle_reports_chain(self):
        wb = self._wb_with_external_ranges()
        wb["1号楼"]["B2"] = "=B3+1"
        wb["1号楼"]["B3"] = "=B2-1"
        with self.assertRaises(FormulaError) as ctx:
            WorkbookAnalyzer(wb, "1号楼", "园区汇总", "B2").analyze()
        message = str(ctx.exception)
        self.assertIn("循环引用", message)
        self.assertIn("1号楼!B2", message)
        self.assertIn("1号楼!B3", message)

    def test_non_integer_base_rejected(self):
        wb = self._wb_with_external_ranges()
        wb["2号楼"]["B3"] = 11.5
        with self.assertRaises(FormulaError) as ctx:
            WorkbookAnalyzer(wb, "1号楼", "园区汇总", "B2").analyze()
        self.assertIn("2号楼!B3", str(ctx.exception))
        self.assertIn("整数", str(ctx.exception))

    def test_integer_float_accepted(self):
        wb = self._wb_with_external_ranges()
        wb["2号楼"]["B3"] = 11.0
        result = WorkbookAnalyzer(wb, "1号楼", "园区汇总", "B2").analyze()
        self.assertEqual(result.external_constant, 7 + 11 + 13)

    def test_cross_workbook_ref_rejected(self):
        with self.assertRaises(FormulaError):
            parse_formula("=[外部.xlsx]Sheet1!A1+B2")

    def test_decimal_literal_rejected(self):
        with self.assertRaises(FormulaError):
            parse_formula("=A1+1.5")

    def test_total_target_parsing(self):
        sheet, coord = parse_total_target("'园区 汇总'!$D$10", "1号楼")
        self.assertEqual((sheet, coord), ("园区 汇总", "D10"))
        sheet, coord = parse_total_target("D2", "1号楼")
        self.assertEqual((sheet, coord), ("1号楼", "D2"))


class BuildingInternalFormulaTest(unittest.TestCase):
    def test_internal_formula_rewritten_and_editable(self):
        wb = Workbook()
        s1 = wb.active
        s1.title = "1号楼"
        s1["B2"] = 100
        s1["B3"] = 30
        s1["B4"] = "=B2-B3"  # 楼内减法
        s2 = wb.create_sheet("2号楼")
        s2["B2"] = 500
        sm = wb.create_sheet("园区汇总")
        sm["B2"] = "='1号楼'!B4+'2号楼'!B2"

        result = WorkbookAnalyzer(wb, "1号楼", "园区汇总", "B2").analyze()
        out = _roundtrip(build_workbook(wb, result))
        s1_out = out["1号楼"]
        # 外部常数 500；楼内公式重写为引用本楼栋读数
        self.assertEqual(out[SUMMARY_SHEET]["A1"].value, 500)
        self.assertEqual(s1_out["B4"].value, "='1号楼'!B2-'1号楼'!B3")

        values = {("1号楼", "B2"): 140, ("1号楼", "B3"): 25, ("2号楼", "B2"): 500}
        # 原表同样修改后的重算值
        expected = (140 - 25) + 500
        self.assertEqual(
            _eval_simple(
                out[SUMMARY_SHEET]["B1"].value, values,
                constants={(SUMMARY_SHEET, "A1"): 500},
            ),
            expected,
        )

    def test_external_formula_referencing_building_reading_stays_live(self):
        """外部工作表的公式里引用本楼栋读数：该读数必须保持可联动，不得折进常数。"""
        wb = Workbook()
        s1 = wb.active
        s1.title = "1号楼"
        s1["B2"] = 100  # 本楼栋读数
        s2 = wb.create_sheet("2号楼")
        s2["B2"] = 70
        # 外部公式 = 自身底数 + 本楼栋读数（跨表反向引用）
        s2["B3"] = "=B2+'1号楼'!B2"
        sm = wb.create_sheet("园区汇总")
        sm["B2"] = "='2号楼'!B3"

        result = WorkbookAnalyzer(wb, "1号楼", "园区汇总", "B2").analyze()
        # 外部常数只含 2 号楼自身底数 70；本楼栋 B2 必须作为引用项保留
        self.assertEqual(result.external_constant, 70)
        self.assertEqual(
            [(t.sign, t.ref.sheet, t.ref.coord) for t in result.building_terms],
            [(1, "1号楼", "B2")],
        )
        out = _roundtrip(build_workbook(wb, result))
        formula = out[SUMMARY_SHEET]["B1"].value
        values = {("1号楼", "B2"): 160, ("2号楼", "B2"): 70}
        # 原表把 B2 改成 160 后：70 + 160 = 230
        self.assertEqual(
            _eval_simple(formula, values, constants={(SUMMARY_SHEET, "A1"): 70}),
            230,
        )


class LayoutTest(unittest.TestCase):
    def test_merged_cells_and_styles_survive_export(self):
        wb = Workbook()
        s1 = wb.active
        s1.title = "1号楼"
        s1.merge_cells("A1:B1")
        s1["A1"] = "1号楼月度电量"
        s1["B2"] = 100
        s1["B3"] = 200
        s1.column_dimensions["A"].width = 18
        sm = wb.create_sheet("园区汇总")
        sm["B2"] = "=SUM('1号楼'!B2:B3)"

        result = WorkbookAnalyzer(wb, "1号楼", "园区汇总", "B2").analyze()
        # 含合并单元格时导出不应抛 MergedCell 只读异常
        out = _roundtrip(build_workbook(wb, result))
        ws = out["1号楼"]
        self.assertEqual(ws["A1"].value, "1号楼月度电量")
        self.assertIn("A1:B1", [str(r) for r in ws.merged_cells.ranges])
        self.assertEqual(ws.column_dimensions["A"].width, 18)
        self.assertEqual(ws["B2"].value, 100)

    def test_building_sheet_named_like_summary_is_remapped(self):
        wb = Workbook()
        s1 = wb.active
        s1.title = SUMMARY_SHEET  # 本楼栋表与保留汇总表同名
        s1["B2"] = 100
        sm = wb.create_sheet("园区汇总")
        sm["B2"] = f"='{SUMMARY_SHEET}'!B2+5"

        result = WorkbookAnalyzer(wb, SUMMARY_SHEET, "园区汇总", "B2").analyze()
        out = _roundtrip(build_workbook(wb, result))
        self.assertEqual(len(out.sheetnames), 2)
        self.assertIn(SUMMARY_SHEET, out.sheetnames)
        building_name = [n for n in out.sheetnames if n != SUMMARY_SHEET][0]
        formula = out[SUMMARY_SHEET]["B1"].value
        # 公式引用必须指向改名后的楼栋表，不得悬空
        self.assertIn(building_name, formula)
        values = {(building_name, "B2"): 140, (SUMMARY_SHEET, "B2"): 140}
        self.assertEqual(
            _eval_simple(formula, values, constants={(SUMMARY_SHEET, "A1"): 5}),
            145,
        )


class WebFlowTest(unittest.TestCase):
    def setUp(self):
        from app.server import create_app

        self.tmp_data = os.environ.get("TEST_DATA_DIR")
        os.environ["DATA_DIR"] = self.tmp_data
        self.app = create_app()
        self.client = self.app.test_client()

    def _upload(self, wb_bytes: bytes):
        return self.client.post(
            "/upload",
            data={"file": (io.BytesIO(wb_bytes), "园区电量.xlsx")},
            content_type="multipart/form-data",
        )

    def test_upload_export_and_error_pages(self):
        wb = Workbook()
        s1 = wb.active
        s1.title = "1号楼"
        s1["B2"] = 100
        s2 = wb.create_sheet("2号楼")
        s2["B2"] = 200
        sm = wb.create_sheet("园区汇总")
        sm["B2"] = "='1号楼'!B2+'2号楼'!B2"

        resp = self._upload(save_workbook(wb))
        self.assertEqual(resp.status_code, 200)
        self.assertIn("1号楼", resp.get_data(as_text=True))

        import re

        m = re.search(r'name="file_id"[^>]*value="([^"]+)"', resp.get_data(as_text=True))
        self.assertIsNotNone(m)
        file_id = m.group(1)

        resp = self.client.post(
            "/export",
            data={
                "file_id": file_id,
                "filename": "园区电量.xlsx",
                "building_sheet": "1号楼",
                "total_cell": "园区汇总!B2",
            },
        )
        self.assertEqual(resp.status_code, 200)
        self.assertIn("attachment", resp.headers.get("Content-Disposition", ""))
        out = load_workbook(io.BytesIO(resp.data))
        self.assertEqual(out[SUMMARY_SHEET]["A1"].value, 200)

        # 不支持公式 -> 422 且页面指出单元格
        wb["园区汇总"]["B2"] = "=SUM('1号楼'!B2)*2"
        resp = self._upload(save_workbook(wb))
        m = re.search(r'name="file_id"[^>]*value="([^"]+)"', resp.get_data(as_text=True))
        resp_err = self.client.post(
            "/export",
            data={
                "file_id": m.group(1),
                "filename": "园区电量.xlsx",
                "building_sheet": "1号楼",
                "total_cell": "园区汇总!B2",
            },
        )
        self.assertEqual(resp_err.status_code, 422)
        self.assertIn("园区汇总!B2", resp_err.get_data(as_text=True))

    def test_reject_non_xlsx(self):
        resp = self.client.post(
            "/upload",
            data={"file": (io.BytesIO(b"not excel"), "bad.txt")},
            content_type="multipart/form-data",
            follow_redirects=True,
        )
        self.assertEqual(resp.status_code, 200)
        self.assertIn(".xlsx", resp.get_data(as_text=True))


if __name__ == "__main__":
    unittest.main(verbosity=2)
