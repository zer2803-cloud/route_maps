"""Build PAB product schedule XLSX from the extracted Word plan."""

from __future__ import annotations

from datetime import date
from pathlib import Path

from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

HEADERS = ("任务", "任务摘要", "预估完成节点", "实际完成节点", "超时原因", "任务责任人")
OUTPUT_NAME = "PAB产品计划表.xlsx"

# (task, summary, planned, actual, overtime_reason, owner)
# actual/overtime empty: source Word file has planned dates only.
ROWS: list[tuple[str, str, date, date | None, str, str]] = [
    (
        "第一期·零件图纸整理归档",
        "四批型号单级和双级所涉零件图纸整理完毕并分类归档。范围：PAB060、PAB090、PAB115、PAB142。协助人：吴晨阳。",
        date(2026, 8, 24),
        None,
        "",
        "王杰",
    ),
    (
        "第一期·零件下单采购",
        "各零件形成采购表单并填写数量，单级和双级各10台。范围：PAB060、PAB090、PAB115、PAB142。",
        date(2026, 8, 25),
        None,
        "",
        "王杰",
    ),
    (
        "第一期·标准件下单采购",
        "各标准件形成采购表单并填写数量。范围：PAB060、PAB090、PAB115、PAB142。",
        date(2026, 8, 25),
        None,
        "",
        "王杰",
    ),
    (
        "装配线布局",
        "装配线布局完成。协助人：吴金平、王杰、吴晨阳。",
        date(2026, 9, 15),
        None,
        "",
        "莫华华",
    ),
    (
        "装配工具到位",
        "轴用卡钳、孔用卡钳、镊子、压机、润滑脂、螺纹禁锢胶、密封厌氧胶等装配工具到位。协助人：王杰。",
        date(2026, 9, 10),
        None,
        "",
        "吴晨阳",
    ),
    (
        "PAB060·零件到位",
        "PAB060 零件到位。协助人：吴晨阳。",
        date(2026, 9, 24),
        None,
        "",
        "王杰",
    ),
    (
        "PAB060·工装图纸绘制",
        "PAB060 工装图纸绘制完成。协助人：吴晨阳。",
        date(2026, 8, 28),
        None,
        "",
        "王杰",
    ),
    (
        "PAB060·工装实物到位",
        "PAB060 工装实物到位。协助人：吴晨阳。",
        date(2026, 9, 15),
        None,
        "",
        "王杰",
    ),
    (
        "PAB060·排针机安装调试",
        "排针机1安装到位并调试完毕。协助人：王杰。",
        date(2026, 9, 24),
        None,
        "",
        "吴金平",
    ),
    (
        "PAB060·单级样机",
        "完成 PAB060 单级样机。协助人：吴晨阳。",
        date(2026, 9, 26),
        None,
        "",
        "王杰",
    ),
    (
        "PAB060·双级样机",
        "完成 PAB060 双级样机。协助人：吴晨阳。",
        date(2026, 9, 28),
        None,
        "",
        "王杰",
    ),
    (
        "PAB090·零件到位",
        "PAB090 零件到位。协助人：吴晨阳。",
        date(2026, 9, 27),
        None,
        "",
        "王杰",
    ),
    (
        "PAB090·工装图纸绘制",
        "PAB090 工装图纸绘制完成。协助人：吴晨阳。",
        date(2026, 9, 1),
        None,
        "",
        "王杰",
    ),
    (
        "PAB090·工装实物到位",
        "PAB090 工装实物到位。协助人：吴晨阳。",
        date(2026, 9, 20),
        None,
        "",
        "王杰",
    ),
    (
        "PAB090·排针机安装调试",
        "排针机1安装到位并调试完毕。协助人：王杰。",
        date(2026, 9, 27),
        None,
        "",
        "吴金平",
    ),
    (
        "PAB090·单级样机",
        "完成 PAB090 单级样机。协助人：吴晨阳。",
        date(2026, 9, 29),
        None,
        "",
        "王杰",
    ),
    (
        "PAB090·双级样机",
        "完成 PAB090 双级样机。协助人：吴晨阳。",
        date(2026, 9, 30),
        None,
        "",
        "王杰",
    ),
    (
        "PAB115·零件到位",
        "PAB115 零件到位（原文标注：国庆假期）。协助人：吴晨阳。",
        date(2026, 10, 8),
        None,
        "",
        "王杰",
    ),
    (
        "PAB115·工装图纸绘制",
        "PAB115 工装图纸绘制完成。协助人：吴晨阳。",
        date(2026, 9, 4),
        None,
        "",
        "王杰",
    ),
    (
        "PAB115·工装实物到位",
        "PAB115 工装实物到位。协助人：吴晨阳。",
        date(2026, 9, 25),
        None,
        "",
        "王杰",
    ),
    (
        "PAB115·排针机安装调试",
        "排针机1安装到位并调试完毕。协助人：王杰。",
        date(2026, 10, 8),
        None,
        "",
        "吴金平",
    ),
    (
        "PAB115·单级样机",
        "完成 PAB115 单级样机。协助人：吴晨阳。",
        date(2026, 10, 9),
        None,
        "",
        "王杰",
    ),
    (
        "PAB115·双级样机",
        "完成 PAB115 双级样机。协助人：吴晨阳。",
        date(2026, 10, 11),
        None,
        "",
        "王杰",
    ),
    (
        "PAB142·零件到位",
        "PAB142 零件到位。协助人：吴晨阳。",
        date(2026, 10, 12),
        None,
        "",
        "王杰",
    ),
    (
        "PAB142·工装图纸绘制",
        "PAB142 工装图纸绘制完成。协助人：吴晨阳。",
        date(2026, 9, 4),
        None,
        "",
        "王杰",
    ),
    (
        "PAB142·工装实物到位",
        "PAB142 工装实物到位。协助人：吴晨阳。",
        date(2026, 9, 29),
        None,
        "",
        "王杰",
    ),
    (
        "PAB142·排针机安装调试",
        "排针机1、排针机2安装到位并调试完毕。协助人：王杰。",
        date(2026, 10, 12),
        None,
        "",
        "吴金平",
    ),
    (
        "PAB142·单级样机",
        "完成 PAB142 单级样机。协助人：吴晨阳。",
        date(2026, 10, 13),
        None,
        "",
        "王杰",
    ),
    (
        "PAB142·双级样机",
        "完成 PAB142 双级样机。协助人：吴晨阳。",
        date(2026, 10, 15),
        None,
        "",
        "王杰",
    ),
    (
        "第二期·零件图纸整理归档",
        "四批型号单级和双级所涉零件图纸整理完毕并分类归档。范围：PAB180、PAB220。协助人：吴晨阳。",
        date(2026, 9, 7),
        None,
        "",
        "王杰",
    ),
    (
        "第二期·零件下单采购",
        "各零件形成采购表单并填写数量，单级和双级各10台。范围：PAB180、PAB220。",
        date(2026, 9, 10),
        None,
        "",
        "王杰",
    ),
    (
        "第二期·标准件下单采购",
        "各标准件形成采购表单并填写数量。范围：PAB180、PAB220。",
        date(2026, 9, 10),
        None,
        "",
        "王杰",
    ),
    (
        "PAB180·零件到位",
        "PAB180 零件到位（原文标注：国庆假期）。协助人：吴晨阳。",
        date(2026, 10, 15),
        None,
        "",
        "王杰",
    ),
    (
        "PAB180·工装图纸绘制",
        "PAB180 工装图纸绘制完成。协助人：吴晨阳。",
        date(2026, 9, 11),
        None,
        "",
        "王杰",
    ),
    (
        "PAB180·工装实物到位",
        "PAB180 工装实物到位。协助人：吴晨阳。",
        date(2026, 10, 8),
        None,
        "",
        "王杰",
    ),
    (
        "PAB180·排针机安装调试",
        "排针机1、排针机2安装到位并调试完毕。协助人：王杰。",
        date(2026, 10, 15),
        None,
        "",
        "吴金平",
    ),
    (
        "PAB180·单级样机",
        "完成 PAB180 单级样机。协助人：吴晨阳。",
        date(2026, 10, 16),
        None,
        "",
        "王杰",
    ),
    (
        "PAB180·双级样机",
        "完成 PAB180 双级样机。协助人：吴晨阳。",
        date(2026, 10, 19),
        None,
        "",
        "王杰",
    ),
    (
        "PAB220·零件到位",
        "PAB220 零件到位（原文标注：国庆假期）。协助人：吴晨阳。",
        date(2026, 10, 20),
        None,
        "",
        "王杰",
    ),
    (
        "PAB220·工装图纸绘制",
        "PAB220 工装图纸绘制完成。协助人：吴晨阳。",
        date(2026, 9, 14),
        None,
        "",
        "王杰",
    ),
    (
        "PAB220·工装实物到位",
        "PAB220 工装实物到位。协助人：吴晨阳。",
        date(2026, 10, 10),
        None,
        "",
        "王杰",
    ),
    (
        "PAB220·排针机安装调试",
        "排针机1、排针机2安装到位并调试完毕。协助人：王杰。",
        date(2026, 10, 20),
        None,
        "",
        "吴金平",
    ),
    (
        "PAB220·单级样机",
        "完成 PAB220 单级样机。协助人：吴晨阳。",
        date(2026, 10, 21),
        None,
        "",
        "王杰",
    ),
    (
        "PAB220·双级样机",
        "完成 PAB220 双级样机。协助人：吴晨阳。",
        date(2026, 10, 23),
        None,
        "",
        "王杰",
    ),
]


def _project_root() -> Path:
    return Path(__file__).resolve().parent.parent


def build_workbook(output_path: Path | None = None) -> Path:
    wb = Workbook()
    ws = wb.active
    ws.title = "时间节点计划表"

    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(name="微软雅黑", bold=True, color="FFFFFF", size=11)
    body_font = Font(name="微软雅黑", size=10)
    wrap = Alignment(wrap_text=True, vertical="center")
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin = Border(
        left=Side(style="thin", color="B7C9D6"),
        right=Side(style="thin", color="B7C9D6"),
        top=Side(style="thin", color="B7C9D6"),
        bottom=Side(style="thin", color="B7C9D6"),
    )
    alt_fill = PatternFill("solid", fgColor="F3F7FA")
    overtime_fill = PatternFill("solid", fgColor="FF0000")

    for col, header in enumerate(HEADERS, start=1):
        cell = ws.cell(1, col, header)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center
        cell.border = thin

    last_data_row = 1 + len(ROWS)
    for i, (task, summary, planned, actual, reason, owner) in enumerate(ROWS, start=2):
        values = (task, summary, planned, actual, reason, owner)
        for col, value in enumerate(values, start=1):
            cell = ws.cell(i, col, value)
            cell.font = body_font
            cell.border = thin
            cell.alignment = wrap if col in (1, 2, 5) else center
            if i % 2 == 0:
                cell.fill = alt_fill
            if col in (3, 4) and value is not None:
                cell.number_format = "YYYY-MM-DD"

    # D > C → red fill on the actual-date cell. Extends past current rows for later fills.
    rule_range = f"D2:D{max(last_data_row, 200)}"
    ws.conditional_formatting.add(
        rule_range,
        FormulaRule(
            formula=["AND(ISNUMBER(D2),ISNUMBER(C2),D2>C2)"],
            fill=overtime_fill,
            font=Font(name="微软雅黑", color="FFFFFF", bold=True, size=10),
        ),
    )

    widths = (28, 62, 16, 16, 28, 14)
    for idx, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(idx)].width = width
    ws.row_dimensions[1].height = 22
    for row in range(2, last_data_row + 1):
        ws.row_dimensions[row].height = 36

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:F{last_data_row}"
    ws.auto_filter.add_sort_condition(f"C2:C{last_data_row}")

    note_row = last_data_row + 2
    ws.merge_cells(start_row=note_row, start_column=1, end_row=note_row, end_column=6)
    note = ws.cell(
        note_row,
        1,
        "说明：数据来自《PAB产品计划表》Word原文。原文仅有预估节点，实际完成节点留空待填；"
        "填写后若实际完成节点晚于预估完成节点，D列单元格自动红色底纹。"
        "协助人写入任务摘要；任务责任人取原文「责任人」。",
    )
    note.font = Font(name="微软雅黑", size=9, color="6B7280")
    note.alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[note_row].height = 40

    date_dv = DataValidation(type="date", operator="greaterThan", formula1="DATE(2020,1,1)", allow_blank=True)
    date_dv.add(f"C2:D{max(last_data_row, 200)}")
    ws.add_data_validation(date_dv)

    target = output_path or (_project_root() / OUTPUT_NAME)
    wb.save(target)
    return target


if __name__ == "__main__":
    path = build_workbook()
    print(f"已生成：{path}")
