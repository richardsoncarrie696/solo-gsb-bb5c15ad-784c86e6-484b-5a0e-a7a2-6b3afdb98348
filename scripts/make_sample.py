"""生成示例工作簿 samples/sample_impact.xlsx。

覆盖场景：
  - 跨工作表引用（=参数!B1*参数!B2）
  - 带引号/空格的工作表名（='月报 2026'!...）
  - 矩形区域引用（=SUM(计算!B1:B4)）
  - 整列引用（=SUM(参数!B:B)）
  - 多级依赖链（参数 -> 计算 -> 汇总 -> 月报）
  - 循环引用（循环!A1 <-> 循环!B1，以及 循环!C1 自环）
  - 名称管理器：工作簿级 / 工作表级同名遮蔽、区域 / 多区域 / 嵌套 / 公式 /
    常量名称，以及未定义、#REF!、外部工作簿等可审计问题
  - 依赖健康审计（审计示例：缺失工作表 / 越界引用 / 空引用 / 高扇出 / 孤立公式）
"""
from pathlib import Path

from openpyxl import Workbook
from openpyxl.workbook.defined_name import DefinedName

OUT = Path(__file__).resolve().parent.parent / "samples" / "sample_impact.xlsx"


def main() -> None:
    wb = Workbook()

    # ---- 参数表：纯输入 ----
    ws = wb.active
    ws.title = "参数"
    rows = [
        ("项目", "值"),
        ("单价", 10),
        ("数量", 5),
        ("折扣率", 0.9),
        ("税率", 0.13),
        ("月数", 12),
    ]
    for r in rows:
        ws.append(r)

    # ---- 计算表：跨表引用 + 链式依赖 ----
    ws = wb.create_sheet("计算")
    ws.append(["项目", "值"])
    ws["A2"] = "小计"
    ws["B2"] = "=参数!B2*参数!B3"          # 单价 * 数量
    ws["A3"] = "折后"
    ws["B3"] = "=B2*参数!B4"               # 小计 * 折扣率
    ws["A4"] = "税额"
    ws["B4"] = "=B3*参数!B5"               # 折后 * 税率
    ws["A5"] = "月合计"
    ws["B5"] = "=B3+B4"                    # 折后 + 税额
    ws["A6"] = "年合计"
    ws["B6"] = "=B5*参数!B6"               # 月合计 * 月数

    # ---- 汇总表：跨表矩形区域引用 ----
    ws = wb.create_sheet("汇总")
    ws.append(["指标", "值"])
    ws["A2"] = "年合计(跨表)"
    ws["B2"] = "=计算!B6"
    ws["A3"] = "季度合计"
    ws["B3"] = "=计算!B6/4"
    ws["A4"] = "计算区求和(区域)"
    ws["B4"] = "=SUM(计算!B2:B5)"
    ws["A5"] = "参数区求和(区域)"
    ws["B5"] = "=SUM(参数!B2:B6)"
    ws["A6"] = "双区合计"
    ws["B6"] = "=SUM(计算!B2:B5)+SUM(参数!B2:B5)"
    ws["A7"] = "整列引用"
    ws["B7"] = "=SUM(参数!B:B)"

    # ---- 月报 2026：带空格的工作表名，公式内需引号 ----
    ws = wb.create_sheet("月报 2026")
    ws.append(["月份", "月均"])
    ws["A2"] = "月均(来自年合计)"
    ws["B2"] = "=计算!B6/12"
    ws["A3"] = "月均(来自汇总)"
    ws["B3"] = "=汇总!B2/12"
    ws["A4"] = "校验"
    ws["B4"] = "=B2-B3"

    # ---- 循环表：故意制造循环引用，用于演示检测 ----
    ws = wb.create_sheet("循环")
    ws["A1"] = "=B1+1"
    ws["B1"] = "=A1+1"
    ws["C1"] = "=C1*2"                     # 自环
    ws["D1"] = "=A1+10"                    # 受循环影响的下游单元格

    # ---- 审计数据表：200 行输入，供高扇出公式引用 ----
    ws = wb.create_sheet("数据")
    ws["A1"] = "序号"
    ws["B1"] = "值"
    for i in range(1, 201):
        ws.cell(row=i + 1, column=1, value=i)
        ws.cell(row=i + 1, column=2, value=i * 1.5)

    # ---- 审计示例表：刻意覆盖全部审计问题类型 ----
    ws = wb.create_sheet("审计示例")
    ws["A1"] = "缺失工作表"
    ws["B1"] = "=不存在的表!A1+1"           # error: missing_sheet
    ws["A2"] = "越界引用"
    ws["B2"] = "=数据!B999+1"               # warning: out_of_range
    ws["A3"] = "空单元格引用"
    ws["B3"] = "=参数!B8*2"                 # warning: empty_ref（参数表已用至第 7 行，B8 留空）
    ws["A4"] = "高扇出"
    ws["B4"] = "=SUM(数据!B2:B201)"         # warning: high_fanout（200 个直接依赖）
    ws["A5"] = "孤立公式"
    ws["B5"] = "=1+1"                       # info: orphan_formula
    ws["A6"] = "正常公式"
    ws["B6"] = "=参数!B2+1"                 # 引用合法非空
    ws["A7"] = "正常公式(下游)"
    ws["B7"] = "=B6*2"                      # 无问题：引用 B6，使 B6 不是孤立公式

    # 参数表追加标签行，使 B8 落在已用范围内；B8 本身留空，供空引用审计演示
    wb["参数"]["A8"] = "（空值行）"

    # ---- 名称表：演示名称管理器（工作簿级 / 工作表级同名、各类解析问题）----
    ws = wb.create_sheet("名称")
    ws["A1"] = "单价（工作簿级名称 单价 → 参数!B2）"
    ws["B1"] = "=单价"
    ws["A2"] = "数量（工作簿级名称 数量 → 参数!B3）"
    ws["B2"] = "=单价*数量"
    ws["A3"] = "折扣金额（工作表级名称 折扣 → 参数!B4，遮蔽全局）"
    ws["B3"] = "=折扣"
    ws["A4"] = "跨表引用计算表本地名称（计算!系数）"
    ws["B4"] = "=计算!系数"
    ws["A5"] = "区域名称（数据列 → 数据!B2:B201）"
    ws["B5"] = "=SUM(数据列)"
    ws["A6"] = "嵌套名称（年化系数 → 月数 → 参数!B6）"
    ws["B6"] = "=年化系数"
    ws["A7"] = "公式表达式名称（折后小计 = 单价*数量*折扣率）"
    ws["B7"] = "=折后小计"
    ws["A8"] = "常量名称（税率常量 0.13，不产生依赖边）"
    ws["B8"] = "=税率常量"
    ws["A9"] = "未定义名称（error: undefined_name）"
    ws["B9"] = "=不存在的名称+1"
    ws["A10"] = "已删除引用名称（error: invalid_name，#REF!）"
    ws["B10"] = "=失效名称"
    ws["A11"] = "指向不存在工作表的名称（error: missing_sheet）"
    ws["B11"] = "=幽灵表名称"
    ws["A12"] = "限定到不存在工作表的名称（error: undefined_name）"
    ws["B12"] = "=不存在的表!某名称"
    ws["A13"] = "外部工作簿名称（warning: external_name）"
    ws["B13"] = "=外部预算"
    ws["A14"] = "经名称越界引用（warning: out_of_range）"
    ws["B14"] = "=越界数据名"
    ws["A15"] = "名称循环（环甲→环乙→环甲，error: invalid_name）"
    ws["B15"] = "=环甲"
    ws["A16"] = "名称链下游（验证影响面经名称传播）"
    ws["B16"] = "=B2*年化系数"

    # 工作簿级定义名称
    wb.defined_names["单价"] = DefinedName("单价", attr_text="参数!$B$2")
    wb.defined_names["数量"] = DefinedName("数量", attr_text="参数!$B$3")
    # 与「名称」表工作表级名称同名：在该表内使用时本地遮蔽全局
    wb.defined_names["折扣"] = DefinedName("折扣", attr_text="参数!$B$2")
    wb["名称"].defined_names["折扣"] = DefinedName("折扣", attr_text="参数!$B$4")
    # 计算表工作表级名称（供 名称!B4 跨表限定演示）
    wb["计算"].defined_names["系数"] = DefinedName("系数", attr_text="参数!$B$4")
    wb.defined_names["数据列"] = DefinedName("数据列", attr_text="数据!$B$2:$B$201")
    # 嵌套名称：年化系数 -> 月数 -> 参数!B6
    wb.defined_names["月数"] = DefinedName("月数", attr_text="参数!$B$6")
    wb.defined_names["年化系数"] = DefinedName("年化系数", attr_text="月数")
    # 公式表达式名称（其中的引用仍展开为依赖边）
    wb.defined_names["折后小计"] = DefinedName(
        "折后小计", attr_text="参数!$B$2*参数!$B$3*参数!$B$4")
    # 常量名称：不产生依赖边，仅记录 resolved=constant
    wb.defined_names["税率常量"] = DefinedName("税率常量", attr_text="0.13")
    # 解析问题类名称
    wb.defined_names["失效名称"] = DefinedName("失效名称", attr_text="#REF!")
    wb.defined_names["幽灵表名称"] = DefinedName(
        "幽灵表名称", attr_text="不存在的表!$A$1")
    wb.defined_names["外部预算"] = DefinedName(
        "外部预算", attr_text="[1]年度预算.xlsx!汇总!$B$2")
    wb.defined_names["越界数据名"] = DefinedName("越界数据名", attr_text="数据!$B$999")
    wb.defined_names["环甲"] = DefinedName("环甲", attr_text="环乙")
    wb.defined_names["环乙"] = DefinedName("环乙", attr_text="环甲")

    OUT.parent.mkdir(parents=True, exist_ok=True)
    wb.save(OUT)
    print(f"已生成 {OUT}")


if __name__ == "__main__":
    main()
