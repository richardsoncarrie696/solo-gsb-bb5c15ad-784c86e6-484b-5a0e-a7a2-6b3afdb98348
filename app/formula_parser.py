"""Excel 公式引用解析。

从公式文本中抽取单元格引用 / 矩形区域引用（含跨工作表、带引号工作表名、
整行/整列区域）以及**定义名称**（名称管理器中的名称，含跨表限定
``Sheet!Name``），输出未展开的 RawRef / RawName 列表，由 graph 模块结合
工作表实际尺寸与定义名称表展开为具体单元格。
"""
from __future__ import annotations

import re
from dataclasses import dataclass

from openpyxl.utils import column_index_from_string

# 工作表名前缀：'月报 2026'! 或 Sheet1! 或 参数!
_SHEET = r"(?:'(?P<qsheet>[^']+(?:''[^']*)*)'|(?P<usheet>[A-Za-z0-9_.一-鿿]+))!"
# 限定名称处的第二份表名前缀（命名组不能重名）
_SHEET2 = r"(?:'(?P<qsheet2>[^']+(?:''[^']*)*)'|(?P<usheet2>[A-Za-z0-9_.一-鿿]+))!"
_CELL = r"\$?[A-Za-z]{1,3}\$?[0-9]{1,7}"
_COL = r"\$?[A-Za-z]{1,3}"
_ROW = r"\$?[0-9]{1,7}"
# 名称主体：首字符为字母 / 下划线 / 反斜杠 / CJK，后续可含 . 等；不含 $
_NAME_HEAD = r"(?:[A-Za-z_\\一-鿿][A-Za-z0-9_.一-鿿]*)"

_TOKEN_RE = re.compile(
    r"""
      (?P<string>"(?:[^"]|"")*")                        # 字符串字面量，忽略
    | (?P<error>\#[A-Za-z0-9!/?]+)                      # #REF! 等错误字面量，忽略
    | (?P<sheetref>{sheet}(?:{cell}(?::{cell})?|{col}:{col}|{row}:{row})(?![A-Za-z0-9_.一-鿿$]))
    | (?P<sheetname>{sheet2}{name})                     # 跨表限定的定义名称
    | (?P<structured>{name}?\[(?:[^\[\]]|\[[^\[\]]*\])*\])  # 结构化引用 表[列] / [[#头],[列]]
    | (?P<range>(?<![0-9.]){cell}(?![A-Za-z0-9_.一-鿿$]):{cell}(?![A-Za-z0-9_.一-鿿$])
                 |{col}(?![0-9$]):{col}(?![0-9$])|{row}(?![A-Za-z$]):{row}(?![A-Za-z$]))
    | (?P<func>[A-Za-z_\\一-鿿][A-Za-z0-9_.一-鿿]*(?=\s*\())  # 函数名，忽略
    | (?P<cell>(?<![0-9.]){cell})(?![A-Za-z0-9_.一-鿿$])     # 本表单单元格（排除 1E2 科学计数法）
    | (?P<name>(?<![0-9.]){name})(?!\s*\()              # 本表定义名称（不含函数调用）
    | (?P<other>.)
    """.format(sheet=_SHEET, sheet2=_SHEET2, cell=_CELL, col=_COL, row=_ROW, name=_NAME_HEAD),
    re.VERBOSE,
)
_CELL_RE = re.compile(r"^\$?([A-Za-z]{1,3})\$?([0-9]{1,7})$")


@dataclass(frozen=True)
class RawRef:
    """一条尚未展开的单元格引用。行/列为 None 表示整列/整行区域，需按工作表尺寸截断。"""

    sheet: str | None  # None 表示公式所在工作表
    kind: str  # "cell" | "range" | "cols" | "rows"
    col1: int | None
    row1: int | None
    col2: int | None
    row2: int | None
    text: str  # 原始文本，用于展示


@dataclass(frozen=True)
class RawName:
    """公式中出现的定义名称（名称管理器名称）。

    qualifier 非 None 表示跨表限定，如 ``计算!单价``；为 None 时按公式所在
    工作表先查工作表级名称、再查工作簿级名称。
    """

    name: str
    qualifier: str | None
    text: str


@dataclass(frozen=True)
class RawToken:
    """扫描得到的单个 token：引用或名称。"""

    kind: str  # "ref" | "name"
    ref: RawRef | None = None
    name: RawName | None = None


# 布尔等保留字不是定义名称，公式中直接跳过
_RESERVED_NAMES = {"TRUE", "FALSE"}


def _split_cell(token: str) -> tuple[int, int]:
    m = _CELL_RE.match(token)
    if not m:
        raise ValueError(f"非法单元格引用: {token!r}")
    return column_index_from_string(m.group(1).upper()), int(m.group(2))


def _split_col(token: str) -> int:
    return column_index_from_string(token.replace("$", "").upper())


def _split_row(token: str) -> int:
    return int(token.replace("$", ""))


def _build_ref(sheet: str | None, body: str, full_text: str) -> RawRef | None:
    """把引用主体（不含工作表前缀）解析为 RawRef。"""
    if ":" in body:
        left, right = body.split(":", 1)
        if _CELL_RE.match(left) and _CELL_RE.match(right):
            c1, r1 = _split_cell(left)
            c2, r2 = _split_cell(right)
            return RawRef(sheet, "range", min(c1, c2), min(r1, r2), max(c1, c2), max(r1, r2), full_text)
        if re.fullmatch(_COL, left) and re.fullmatch(_COL, right):
            c1, c2 = _split_col(left), _split_col(right)
            return RawRef(sheet, "cols", min(c1, c2), None, max(c1, c2), None, full_text)
        if re.fullmatch(_ROW, left) and re.fullmatch(_ROW, right):
            r1, r2 = _split_row(left), _split_row(right)
            return RawRef(sheet, "rows", None, min(r1, r2), None, max(r1, r2), full_text)
        return None
    if _CELL_RE.match(body):
        c, r = _split_cell(body)
        return RawRef(sheet, "cell", c, r, c, r, full_text)
    return None


def _sheet_of(m: re.Match) -> str | None:
    sheet = m.group("qsheet")
    if sheet is not None:
        return sheet.replace("''", "'")
    return m.group("usheet")


def tokenize_formula(formula: str) -> list[RawToken]:
    """词法扫描公式，返回引用 / 名称 token（字符串、函数名等忽略）。"""
    if not formula:
        return []
    text = formula[1:] if formula.startswith("=") else formula
    tokens: list[RawToken] = []
    for m in _TOKEN_RE.finditer(text):
        kind = m.lastgroup
        if kind == "sheetref":
            sheet = _sheet_of(m)
            body = m.group("sheetref")
            body = body[body.index("!") + 1:]
            ref = _build_ref(sheet, body, m.group("sheetref"))
            if ref is not None:
                tokens.append(RawToken(kind="ref", ref=ref))
        elif kind == "sheetname":
            sheet = m.group("qsheet2")
            if sheet is not None:
                sheet = sheet.replace("''", "'")
            else:
                sheet = m.group("usheet2")
            name = m.group("sheetname").rsplit("!", 1)[-1]
            tokens.append(RawToken(
                kind="name", name=RawName(name, sheet, m.group("sheetname"))))
        elif kind == "range":
            ref = _build_ref(None, m.group("range"), m.group("range"))
            if ref is not None:
                tokens.append(RawToken(kind="ref", ref=ref))
        elif kind == "cell":
            tokens.append(RawToken(kind="ref", ref=_build_ref(
                None, m.group("cell"), m.group("cell"))))
        elif kind == "name":
            nm = m.group("name")
            if nm.upper() in _RESERVED_NAMES:
                continue
            tokens.append(RawToken(kind="name", name=RawName(nm, None, nm)))
        else:
            continue
    return tokens


def parse_formula(formula: str) -> list[RawRef]:
    """解析公式文本，返回单元格引用列表（未展开，定义名称不在其中）。"""
    return [t.ref for t in tokenize_formula(formula) if t.kind == "ref"]


def parse_names(formula: str) -> list[RawName]:
    """解析公式文本，返回其中出现的定义名称。"""
    return [t.name for t in tokenize_formula(formula) if t.kind == "name"]
