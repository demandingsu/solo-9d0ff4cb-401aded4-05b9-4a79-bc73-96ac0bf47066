"""园区电量试算工具 —— 公式词法/语法分析。

只接受以下语法（超出范围由 :class:`FormulaError` 报错）::

    expr     := term (('+' | '-') term)*
    term     := ['-' | '+'] operand
    operand  := NUMBER | REF | func | '(' expr ')'
    func     := NAME '(' [expr (',' expr)*] ')'
    REF      := [SHEET '!'] (CELL | RANGE)   # CELL/RANGE 中允许 $ 绝对标记

函数只允许 SUM；运算只允许一元正负、加法、减法；引用必须位于同一工作簿。
"""
from __future__ import annotations

import re
from dataclasses import dataclass

_CELL_RE = re.compile(r"^\$?[A-Za-z]{1,3}\$?[0-9]+$")
_RANGE_RE = re.compile(
    r"^\$?[A-Za-z]{1,3}\$?[0-9]+:\$?[A-Za-z]{1,3}\$?[0-9]+$"
)
_NAME_START = re.compile(r"[A-Za-z_一-鿿]")
_NAME_PART = re.compile(r"[A-Za-z0-9_.一-鿿]")


class FormulaError(ValueError):
    """公式不在支持范围内或语法错误。"""


@dataclass(frozen=True)
class Number:
    value: int


@dataclass(frozen=True)
class Ref:
    """单元格或区域引用。

    sheet 为 None 表示与公式所在工作表相同。
    is_range 区分单点与区域；coord 已去除 ``$`` 并大写。
    """

    sheet: str | None
    coord: str
    is_range: bool


@dataclass(frozen=True)
class Func:
    name: str
    args: tuple["Node", ...]


@dataclass(frozen=True)
class BinOp:
    op: str  # '+' 或 '-'
    left: "Node"
    right: "Node"


@dataclass(frozen=True)
class UnaryOp:
    op: str  # '+' 或 '-'
    operand: "Node"


Node = Number | Ref | Func | BinOp | UnaryOp


# --------------------------------------------------------------------------- #
# 词法
# --------------------------------------------------------------------------- #
@dataclass(frozen=True)
class Token:
    kind: str  # NUMBER / REF / NAME / ( / ) , + -
    text: str
    pos: int


def _is_cell_like(coord: str) -> bool:
    clean = coord.replace("$", "").upper()
    return bool(_CELL_RE.match(clean) or _RANGE_RE.match(clean))


def _scan_quoted_sheet_ref(text: str, i: int) -> tuple[Token, int]:
    """读取 '工作表名'!单元格/区域。"""
    start = i
    j = i + 1
    buf = []
    while j < len(text) and text[j] != "'":
        buf.append(text[j])
        j += 1
    if j >= len(text):
        raise FormulaError("工作表名引号未闭合")
    sheet = "".join(buf)
    j += 1
    if j >= len(text) or text[j] != "!":
        raise FormulaError("带引号的工作表名后缺少 '!'")
    j += 1
    ref_start = j
    while j < len(text) and (text[j].isalnum() or text[j] in "$:"):
        j += 1
    coord = text[ref_start:j]
    if not coord or not _is_cell_like(coord):
        raise FormulaError(f"无法识别的单元格引用: {coord}")
    clean = coord.replace("$", "").upper()
    return Token("REF", f"{sheet}!{clean}", start), j


def _scan_bareword(text: str, i: int) -> tuple[Token, int]:
    """读取字母/下划线/中文开头的记号：名称、单元格、区域或 工作表!引用。"""
    start = i
    j = i
    while j < len(text) and _NAME_PART.match(text[j]):
        j += 1
    word = text[i:j]

    if j < len(text) and text[j] == "!":
        # 不带引号的工作表名
        j += 1
        ref_start = j
        while j < len(text) and (text[j].isalnum() or text[j] in "$:"):
            j += 1
        coord = text[ref_start:j]
        if not coord or not _is_cell_like(coord):
            raise FormulaError(f"无法识别的单元格引用: {word}!{coord}")
        clean = coord.replace("$", "").upper()
        return Token("REF", f"{word}!{clean}", start), j

    # 单元格后面紧跟 ':'：区域引用，如 B2:B3
    if _is_cell_like(word) and j < len(text) and text[j] == ":":
        j += 1
        part_start = j
        while j < len(text) and (text[j].isalnum() or text[j] == "$"):
            j += 1
        second = text[part_start:j]
        coord = f"{word}:{second}"
        if not _is_cell_like(coord):
            raise FormulaError(f"无法识别的区域引用: {coord}")
        return Token("REF", coord.replace("$", "").upper(), start), j

    if _is_cell_like(word):
        return Token("REF", word.replace("$", "").upper(), start), j
    return Token("NAME", word, start), j


def tokenize(formula: str) -> list[Token]:
    tokens: list[Token] = []
    text = formula
    i = 0
    n = len(text)
    while i < n:
        ch = text[i]
        if ch.isspace():
            i += 1
            continue
        if ch in "()+-,":
            tokens.append(Token(ch, ch, i))
            i += 1
            continue
        if ch in "*/&=<>":
            raise FormulaError(f"不支持的运算符 '{ch}'（仅支持 +、-）")
        if ch == '"':
            raise FormulaError("不支持字符串常量")
        if ch.isdigit():
            j = i + 1
            while j < n and text[j].isdigit():
                j += 1
            if j < n and text[j] == ".":
                raise FormulaError("底数必须是整数，不支持小数")
            tokens.append(Token("NUMBER", text[i:j], i))
            i = j
            continue
        if ch == "[" or ch == "]":
            raise FormulaError("不支持跨工作簿引用（[ ]），只允许同一工作簿内的引用")
        if ch == "'":
            tok, i = _scan_quoted_sheet_ref(text, i)
            tokens.append(tok)
            continue
        if ch == "$":
            # $ 开头的单元格/区域引用
            j = i + 1
            while j < n and (text[j].isalnum() or text[j] in "$:"):
                j += 1
            coord = text[i:j]
            if not _is_cell_like(coord):
                raise FormulaError(f"无法识别的单元格引用: {coord}")
            tokens.append(Token("REF", coord.replace("$", "").upper(), i))
            i = j
            continue
        if _NAME_START.match(ch):
            tok, i = _scan_bareword(text, i)
            tokens.append(tok)
            continue
        raise FormulaError(f"不支持的字符 '{ch}'（位置 {i}）")
    return tokens


# --------------------------------------------------------------------------- #
# 语法
# --------------------------------------------------------------------------- #
class _Parser:
    def __init__(self, tokens: list[Token]):
        self.tokens = tokens
        self.pos = 0

    def peek(self) -> Token | None:
        return self.tokens[self.pos] if self.pos < len(self.tokens) else None

    def next(self) -> Token:
        tok = self.peek()
        if tok is None:
            raise FormulaError("公式意外结束")
        self.pos += 1
        return tok

    def expect(self, kind: str) -> Token:
        tok = self.next()
        if tok.kind != kind:
            raise FormulaError(f"期望 '{kind}'，实际为 '{tok.text}'")
        return tok

    def parse(self) -> Node:
        node = self.parse_expr()
        if self.peek() is not None:
            raise FormulaError(f"公式末尾存在多余内容: {self.peek().text}")
        return node

    def parse_expr(self) -> Node:
        node = self.parse_term()
        while True:
            tok = self.peek()
            if tok is not None and tok.kind in ("+", "-"):
                self.next()
                right = self.parse_term()
                node = BinOp(tok.kind, node, right)
            else:
                return node

    def parse_term(self) -> Node:
        tok = self.peek()
        if tok is not None and tok.kind in ("+", "-"):
            self.next()
            return UnaryOp(tok.kind, self.parse_term())
        return self.parse_operand()

    def parse_operand(self) -> Node:
        tok = self.next()
        if tok.kind == "NUMBER":
            return Number(int(tok.text))
        if tok.kind == "REF":
            sheet, coord, is_range = _split_ref(tok.text)
            return Ref(sheet, coord, is_range)
        if tok.kind == "NAME":
            return self.parse_func(tok)
        if tok.kind == "(":
            node = self.parse_expr()
            self.expect(")")
            return node
        raise FormulaError(f"意外的记号 '{tok.text}'")

    def parse_func(self, name_tok: Token) -> Node:
        upper = name_tok.text.upper()
        if upper != "SUM":
            raise FormulaError(f"不支持的函数 {name_tok.text}（仅支持 SUM）")
        self.expect("(")
        args: list[Node] = []
        if self.peek() and self.peek().kind != ")":
            args.append(self.parse_expr())
            while self.peek() and self.peek().kind == ",":
                self.next()
                args.append(self.parse_expr())
        self.expect(")")
        return Func(upper, tuple(args))


def _split_ref(text: str) -> tuple[str | None, str, bool]:
    if "!" in text:
        sheet, coord = text.split("!", 1)
    else:
        sheet, coord = None, text
    is_range = ":" in coord
    return sheet, coord, is_range


def parse_formula(formula: str) -> Node:
    """把 ``=A1+SUM(...)`` 形式的公式解析为 AST。"""
    if not isinstance(formula, str) or not formula:
        raise FormulaError("空公式")
    body = formula[1:] if formula.startswith("=") else formula
    tokens = tokenize(body)
    if not tokens:
        raise FormulaError("空公式")
    return _Parser(tokens).parse()


def walk(node: Node):
    """深度优先遍历 AST。"""
    yield node
    if isinstance(node, (BinOp,)):
        yield from walk(node.left)
        yield from walk(node.right)
    elif isinstance(node, UnaryOp):
        yield from walk(node.operand)
    elif isinstance(node, Func):
        for arg in node.args:
            yield from walk(arg)


def iter_refs(node: Node):
    for sub in walk(node):
        if isinstance(sub, Ref):
            yield sub
