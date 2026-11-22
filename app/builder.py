"""根据分析结果生成试算工作簿。

输出工作簿只有两张表：

1. 与本楼栋同名的工作表：布局、样式、底数原样保留，公式重写为
   “本楼栋引用 + 常数”的加和形式（外部底数已并入常数）；
2. ``试算汇总``：一个汇总常数单元格 + 园区总计公式。
"""
from __future__ import annotations

import io
import re

from openpyxl import Workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.utils import quote_sheetname

from .analyzer import AnalyzeResult, Term

SUMMARY_SHEET = "试算汇总"


def _q(sheet: str) -> str:
    return quote_sheetname(sheet)


def render_ref(ref, building_name: str | None = None) -> str:
    """渲染引用。building_name 给定时，把分析期的本楼栋表名映射为输出表名。"""
    if ref.sheet is None:
        base = ref.coord
    else:
        sheet = building_name if building_name is not None and ref.sheet else ref.sheet
        base = f"{_q(sheet)}!{ref.coord}"
    return f"SUM({base})" if ref.is_range else base


def render_terms(terms: list[Term], building_name: str | None = None) -> str:
    """把带符号项列表渲染成只含 +、-、SUM 的公式。"""
    parts: list[str] = []
    constant = 0
    for term in terms:
        if term.is_number:
            constant += term.sign * term.number
        else:
            rendered = render_ref(term.ref, building_name)
            if not parts:
                parts.append(rendered if term.sign > 0 else f"-{rendered}")
            else:
                parts.append(("+" if term.sign > 0 else "-") + rendered)
    head = "".join(parts)
    if constant != 0 or not head:
        token = str(abs(constant))
        if not head:
            head = token if constant > 0 else f"-{token}"
        else:
            head += ("+" if constant > 0 else "-") + token
    return "=" + head


def render_total_formula(result: AnalyzeResult, building_name: str) -> str:
    """园区总计 = 汇总常数单元格 + 本楼栋各引用项。"""
    const_ref = f"{_q(SUMMARY_SHEET)}!$A$1"
    parts = [const_ref]
    for term in result.building_terms:
        rendered = render_ref(term.ref, building_name)
        parts.append(("+" if term.sign > 0 else "-") + rendered)
    return "=" + "".join(parts)


def _safe_building_name(name: str) -> str:
    """输出楼栋表名：清洗非法字符，并避开保留的汇总表名。"""
    base = re.sub(r"[\\/?*\[\]:]", "_", name)[:31] or "楼栋"
    if base != SUMMARY_SHEET:
        return base
    suffix = "_楼栋"
    return base[: 31 - len(suffix)] + suffix


def build_workbook(src_wb, result: AnalyzeResult) -> Workbook:
    out_wb = Workbook()
    default = out_wb.active
    out_wb.remove(default)

    src_ws = src_wb[result.building_sheet]
    building_name = _safe_building_name(result.building_sheet)
    out_ws = out_wb.create_sheet(building_name)

    # 布局：行高列宽、冻结窗格、合并单元格
    out_ws.sheet_format = src_ws.sheet_format
    out_ws.freeze_panes = src_ws.freeze_panes
    for key, dim in src_ws.column_dimensions.items():
        out_ws.column_dimensions[key].width = dim.width
        out_ws.column_dimensions[key].hidden = dim.hidden
    for key, dim in src_ws.row_dimensions.items():
        out_ws.row_dimensions[key].height = dim.height
    for merged in src_ws.merged_cells.ranges:
        out_ws.merge_cells(str(merged))

    rewritten = result.building_formulas
    for row in src_ws.iter_rows():
        for cell in row:
            # 合并区域只有左上角单元格可写，其余为只读 MergedCell
            if isinstance(cell, MergedCell):
                continue
            key = (result.building_sheet, cell.coordinate)
            target = out_ws[cell.coordinate]
            if key in rewritten:
                target.value = render_terms(rewritten[key], building_name)
            else:
                target.value = cell.value
            if cell.has_style:
                target.font = cell.font.copy()
                target.fill = cell.fill.copy()
                target.border = cell.border.copy()
                target.alignment = cell.alignment.copy()
                target.number_format = cell.number_format
                target.protection = cell.protection.copy()

    # 汇总表（第 1 行放常数与总计，第 3 行起放说明）
    ws = out_wb.create_sheet(SUMMARY_SHEET)
    ws["A1"] = result.external_constant
    ws["B1"] = render_total_formula(result, building_name)
    ws["A3"] = "A1：外部底数汇总常数（其余楼栋/工作表，只读）"
    ws["B3"] = "B1：园区总计（公式，随本楼栋读数联动）"
    ws["A4"] = "本楼栋工作表"
    ws["B4"] = building_name
    ws["A5"] = "原总计位置"
    ws["B5"] = f"{result.total_sheet}!{result.total_coord}"
    ws.column_dimensions["A"].width = 42
    ws.column_dimensions["B"].width = 42

    # 打开时强制重算，保证读数改动后总计立即刷新
    out_wb.calculation.fullCalcOnLoad = True
    return out_wb


def save_workbook(wb: Workbook) -> bytes:
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def build_filename(source_name: str, building_sheet: str) -> str:
    stem = re.sub(r"[\\/:*?\"<>|]", "_", source_name.rsplit(".", 1)[0])[:60] or "电量表"
    safe_sheet = re.sub(r"[\\/:*?\"<>|]", "_", building_sheet)[:20]
    return f"{stem}_{safe_sheet}_试算.xlsx"
