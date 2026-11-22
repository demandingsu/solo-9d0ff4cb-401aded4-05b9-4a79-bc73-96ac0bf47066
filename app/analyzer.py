"""依赖分析：展开总计公式、检测超范围/成环、拆分本楼栋与外部底数。

产出（:meth:`WorkbookAnalyzer.analyze`）:

* ``total_terms``：总计公式展平后的带符号叶子项；
* ``external_constant``：全部外部底数折算后的**一个**整数常数；
* ``building_terms``：本楼栋相关的带符号引用项（读数改这里即可联动总计）；
* ``building_formulas``：本楼栋表内公式单元格的重写结果。
"""
from __future__ import annotations

from dataclasses import dataclass

from openpyxl.utils import range_boundaries

from .formula import BinOp, FormulaError, Func, Node, Number, Ref, UnaryOp, parse_formula

CellKey = tuple[str, str]  # (工作表名, 单元格坐标)


@dataclass(frozen=True)
class Term:
    """展平后的带符号叶子项。"""

    sign: int  # +1 / -1
    ref: Ref | None  # None 表示数字常数项
    number: int | None = None

    @property
    def is_number(self) -> bool:
        return self.ref is None


@dataclass
class AnalyzeResult:
    total_sheet: str
    total_coord: str
    building_sheet: str
    total_terms: list[Term]
    external_constant: int
    building_terms: list[Term]
    # 本楼栋工作表内每个公式单元格 -> 重写后的带符号项
    building_formulas: dict[CellKey, list[Term]]


class WorkbookAnalyzer:
    def __init__(self, wb, building_sheet: str, total_sheet: str, total_coord: str):
        self.wb = wb
        self.building_sheet = building_sheet
        self.total_sheet = total_sheet
        self.total_coord = total_coord.upper().replace("$", "")
        self._parse_cache: dict[CellKey, Node] = {}
        self._expand_cache: dict[CellKey, Node] = {}
        self._eval_cache: dict[CellKey, int] = {}

    # ------------------------------------------------------------------ #
    # 基础工具
    # ------------------------------------------------------------------ #
    def _resolve_sheet(self, name: str | None, current: str) -> str:
        sheet = current if name is None else name
        if sheet not in self.wb.sheetnames:
            raise FormulaError(f"工作表 '{sheet}' 不存在于本工作簿")
        return sheet

    def _cell_value(self, key: CellKey):
        sheet, coord = key
        return self.wb[sheet][coord].value

    def _is_formula(self, key: CellKey) -> bool:
        value = self._cell_value(key)
        return isinstance(value, str) and value.startswith("=")

    def _parse(self, key: CellKey) -> Node:
        if key not in self._parse_cache:
            try:
                self._parse_cache[key] = parse_formula(self._cell_value(key))
            except FormulaError as exc:
                raise FormulaError(f"单元格 {key[0]}!{key[1]} 的公式不受支持：{exc}") from exc
        return self._parse_cache[key]

    @staticmethod
    def _iter_range(sheet: str, coord: str):
        try:
            min_col, min_row, max_col, max_row = range_boundaries(coord)
        except (ValueError, TypeError) as exc:
            raise FormulaError(f"非法区域引用 {sheet}!{coord}") from exc
        from openpyxl.utils import get_column_letter

        for row in range(min_row, max_row + 1):
            for col in range(min_col, max_col + 1):
                yield sheet, f"{get_column_letter(col)}{row}"

    def _refs_in_formula(self, key: CellKey) -> list[Ref]:
        from .formula import iter_refs

        return list(iter_refs(self._parse(key)))

    # ------------------------------------------------------------------ #
    # 依赖展开（仅沿“本楼栋”公式单元格向内替换；外部公式在求值阶段处理）
    # ------------------------------------------------------------------ #
    def expand(self, key: CellKey, stack: tuple[CellKey, ...] = ()) -> Node:
        """把公式单元格展开成仅含 Number/Ref(叶子)/SUM(区域)/BinOp/UnaryOp 的树。"""
        if key in self._expand_cache:
            return self._expand_cache[key]
        if key in stack:
            cycle = list(stack[stack.index(key):]) + [key]
            path = " -> ".join(f"{s}!{c}" for s, c in cycle)
            raise FormulaError(f"检测到循环引用，涉及单元格：{path}")
        node = self._parse(key)
        stack = stack + (key,)
        expanded = self._expand_node(node, key[0], stack)
        self._expand_cache[key] = expanded
        return expanded

    def _expand_node(self, node: Node, current_sheet: str, stack: tuple[CellKey, ...]) -> Node:
        if isinstance(node, Number):
            return node
        if isinstance(node, Ref):
            sheet = self._resolve_sheet(node.sheet, current_sheet)
            if node.is_range:
                return Ref(sheet, node.coord, True)
            key = (sheet, node.coord)
            if self._is_formula(key):
                # 只继续展开本楼栋工作表内的公式；外部公式整体留给 eval_value
                if sheet == self.building_sheet:
                    return self.expand(key, stack)
                return Ref(sheet, node.coord, False)
            return Ref(sheet, node.coord, False)
        if isinstance(node, UnaryOp):
            return UnaryOp(node.op, self._expand_node(node.operand, current_sheet, stack))
        if isinstance(node, BinOp):
            return BinOp(
                node.op,
                self._expand_node(node.left, current_sheet, stack),
                self._expand_node(node.right, current_sheet, stack),
            )
        if isinstance(node, Func):  # SUM
            return Func(
                node.name,
                tuple(self._expand_node(arg, current_sheet, stack) for arg in node.args),
            )
        raise FormulaError(f"未知公式节点: {node!r}")  # pragma: no cover

    # ------------------------------------------------------------------ #
    # 外部底数求值（支持区域里的公式单元格，同样限制在支持的语法内）
    # ------------------------------------------------------------------ #
    def eval_value(self, key: CellKey, stack: tuple[CellKey, ...] = ()) -> int:
        if key in self._eval_cache:
            return self._eval_cache[key]
        if key in stack:
            cycle = list(stack[stack.index(key):]) + [key]
            path = " -> ".join(f"{s}!{c}" for s, c in cycle)
            raise FormulaError(f"检测到循环引用，涉及单元格：{path}")

        if self._is_formula(key):
            node = self._parse(key)
            stack = stack + (key,)
            result = self._eval_node(node, key[0], stack)
        else:
            result = self._as_int(self._cell_value(key), key)
        self._eval_cache[key] = result
        return result

    def _eval_node(self, node: Node, current_sheet: str, stack: tuple[CellKey, ...]) -> int:
        if isinstance(node, Number):
            return node.value
        if isinstance(node, Ref):
            sheet = self._resolve_sheet(node.sheet, current_sheet)
            if node.is_range:
                return sum(self.eval_value(k, stack) for k in self._iter_range(sheet, node.coord))
            return self.eval_value((sheet, node.coord), stack)
        if isinstance(node, UnaryOp):
            value = self._eval_node(node.operand, current_sheet, stack)
            return value if node.op == "+" else -value
        if isinstance(node, BinOp):
            left = self._eval_node(node.left, current_sheet, stack)
            right = self._eval_node(node.right, current_sheet, stack)
            return left + right if node.op == "+" else left - right
        if isinstance(node, Func):  # SUM
            return sum(self._eval_node(arg, current_sheet, stack) for arg in node.args)
        raise FormulaError(f"未知公式节点: {node!r}")  # pragma: no cover

    @staticmethod
    def _as_int(value, key: CellKey) -> int:
        if value is None:
            return 0
        if isinstance(value, bool):
            raise FormulaError(f"单元格 {key[0]}!{key[1]} 是布尔值，底数必须是整数")
        if isinstance(value, int):
            return value
        if isinstance(value, float) and value.is_integer():
            return int(value)
        if isinstance(value, str):
            stripped = value.strip()
            if stripped == "":
                return 0
            if stripped.lstrip("-").isdigit():
                return int(stripped)
        raise FormulaError(f"单元格 {key[0]}!{key[1]} 的底数不是整数：{value!r}")

    # ------------------------------------------------------------------ #
    # 拆分：本楼栋引用保留，其余折算成常数
    #
    # 外部工作表的公式单元格整体求值，但其内部对本楼栋工作表的引用必须
    # 作为可编辑引用保留（否则改本楼栋读数后无法联动）。
    # ------------------------------------------------------------------ #
    def _split_ref(self, ref: Ref, current_sheet: str, stack: tuple[CellKey, ...]) -> list[Term]:
        sheet = self._resolve_sheet(ref.sheet, current_sheet)
        if sheet == self.building_sheet:
            if ref.is_range:
                return [Term(1, Ref(sheet, ref.coord, True))]
            key = (sheet, ref.coord)
            if key in stack:
                cycle = list(stack[stack.index(key):]) + [key]
                path = " -> ".join(f"{s}!{c}" for s, c in cycle)
                raise FormulaError(f"检测到循环引用，涉及单元格：{path}")
            # 本楼栋公式单元格内联到原始读数，使总计直接联动可编辑读数
            if self._is_formula(key):
                return self._split_node(self._parse(key), sheet, stack + (key,))
            return [Term(1, Ref(sheet, ref.coord, False))]
        if ref.is_range:
            out: list[Term] = []
            for key in self._iter_range(sheet, ref.coord):
                out.extend(self._split_ref(Ref(sheet, key[1], False), sheet, stack))
            return out
        key = (sheet, ref.coord)
        if self._is_formula(key):
            if key in stack:
                cycle = list(stack[stack.index(key):]) + [key]
                path = " -> ".join(f"{s}!{c}" for s, c in cycle)
                raise FormulaError(f"检测到循环引用，涉及单元格：{path}")
            return self._split_node(self._parse(key), sheet, stack + (key,))
        return [Term(1, None, self.eval_value(key, stack))]

    def _split_node(self, node: Node, current_sheet: str, stack: tuple[CellKey, ...]) -> list[Term]:
        if isinstance(node, Number):
            return [Term(1, None, node.value)]
        if isinstance(node, Ref):
            return self._split_ref(node, current_sheet, stack)
        if isinstance(node, UnaryOp):
            terms = self._split_node(node.operand, current_sheet, stack)
            if node.op == "-":
                terms = [Term(-t.sign, t.ref, t.number) for t in terms]
            return terms
        if isinstance(node, BinOp):
            left = self._split_node(node.left, current_sheet, stack)
            right = self._split_node(node.right, current_sheet, stack)
            if node.op == "-":
                right = [Term(-t.sign, t.ref, t.number) for t in right]
            return left + right
        if isinstance(node, Func):  # SUM
            out: list[Term] = []
            for arg in node.args:
                out.extend(self._split_node(arg, current_sheet, stack))
            return out
        raise FormulaError(f"未知公式节点: {node!r}")  # pragma: no cover

    # ------------------------------------------------------------------ #
    # 主流程
    # ------------------------------------------------------------------ #
    def analyze(self) -> AnalyzeResult:
        if self.building_sheet not in self.wb.sheetnames:
            raise FormulaError(f"工作表 '{self.building_sheet}' 不存在于本工作簿")
        if self.total_sheet not in self.wb.sheetnames:
            raise FormulaError(f"工作表 '{self.total_sheet}' 不存在于本工作簿")
        total_key = (self.total_sheet, self.total_coord)
        try:
            # 坐标合法性（非法时抛 CellCoordinatesException）
            from openpyxl.utils.cell import coordinate_from_string

            coordinate_from_string(self.total_coord)
        except (ValueError, AttributeError) as exc:
            raise FormulaError(f"非法总计单元格坐标：{self.total_coord}") from exc

        # 1. 展开总计（含成环、超范围检测）
        if self._is_formula(total_key):
            expanded = self.expand(total_key)
            total_terms = self._split_node(expanded, self.total_sheet, ())
        else:
            # 总计本身是常数：没有可联动的读数
            total_terms = [Term(1, None, self._as_int(self._cell_value(total_key), total_key))]

        # 2. 本楼栋工作表内所有公式单元格统一重写（保证改读数后表内中间结果仍正确）
        building_formulas: dict[CellKey, list[Term]] = {}
        building_ws = self.wb[self.building_sheet]
        for row in building_ws.iter_rows():
            for cell in row:
                key = (self.building_sheet, cell.coordinate)
                if self._is_formula(key):
                    expanded = self.expand(key)
                    building_formulas[key] = self._split_node(
                        expanded, self.building_sheet, ()
                    )

        # 3. 拆分外部常数与本楼栋项
        external_constant = 0
        building_terms: list[Term] = []
        for term in total_terms:
            if term.is_number:
                external_constant += term.sign * term.number
            else:
                building_terms.append(term)

        return AnalyzeResult(
            total_sheet=self.total_sheet,
            total_coord=self.total_coord,
            building_sheet=self.building_sheet,
            total_terms=total_terms,
            external_constant=external_constant,
            building_terms=building_terms,
            building_formulas=building_formulas,
        )


def parse_total_target(raw: str, default_sheet: str) -> tuple[str, str]:
    """解析用户填写的总计单元格，支持 ``Sheet!B2`` / ``'Sheet 1'!B2`` / ``B2``。"""
    text = raw.strip().lstrip("=")
    if "!" in text:
        sheet_part, coord = text.split("!", 1)
        sheet = sheet_part.strip("'")
    else:
        sheet, coord = default_sheet, text
    coord = coord.replace("$", "").strip().upper()
    if not coord:
        raise FormulaError("总计单元格为空")
    return sheet, coord
