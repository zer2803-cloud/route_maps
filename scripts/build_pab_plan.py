"""Build PAB product schedule XLSX from the extracted Word plan."""

from __future__ import annotations

from datetime import date, timedelta
from pathlib import Path

from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

HEADERS = (
    "阶段",
    "任务",
    "任务简要",
    "任务摘要",
    "预估完成节点",
    "实际完成节点",
    "超时原因",
    "任务责任人",
    "偏移天数",
)
OUTPUT_NAME = "PAB产品计划表.xlsx"

BRIEF_ALIASES = {
    "第一期·零件图纸整理归档": "一期零件图纸整理归档",
    "第一期·零件下单采购": "一期零件下单采购",
    "第一期·标准件下单采购": "一期标准件下单采购",
    "装配线布局": "装配线布局完成",
    "第二期·零件图纸整理归档": "二期零件图纸整理归档",
    "第二期·零件下单采购": "二期零件下单采购",
    "第二期·标准件下单采购": "二期标准件下单采购",
}


def phase_for_task(task: str) -> str:
    if task.startswith("第二期") or task.startswith(("PAB180", "PAB220")):
        return "第二期"
    return "第一期"


def group_for_task(task: str) -> str:
    if task.startswith(("第一期", "第二期")) or task in {"装配线布局", "装配工具到位"}:
        return "前期准备"
    return task.split("·", 1)[0]


def brief_for_task(task: str) -> str:
    return BRIEF_ALIASES.get(task, task)

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


def _thin_border() -> Border:
    return Border(
        left=Side(style="thin", color="B7C9D6"),
        right=Side(style="thin", color="B7C9D6"),
        top=Side(style="thin", color="B7C9D6"),
        bottom=Side(style="thin", color="B7C9D6"),
    )


DAY_COL_START = 7  # Gantt calendar days begin at column G.


def _calendar_range() -> tuple[date, date]:
    dates = [row[2] for row in ROWS]
    dates.extend(row[3] for row in ROWS if row[3] is not None)
    return min(dates), max(dates)


def _calendar_days() -> list[date]:
    start, end = _calendar_range()
    return [start + timedelta(days=i) for i in range((end - start).days + 1)]


def _offset_label(planned: date, actual: date | None) -> str:
    if actual is None:
        return "待填"
    delta = (actual - planned).days
    if delta > 0:
        return f"+{delta}天"
    if delta < 0:
        return f"{delta}天"
    return "0天"


def _fill_bar(ws, row: int, start: date, end: date, color: str) -> None:
    """Paint one cell per calendar day from project start through `end` (inclusive)."""
    fill = PatternFill("solid", fgColor=color)
    last_idx = (end - start).days
    for idx in range(last_idx + 1):
        cell = ws.cell(row, DAY_COL_START + idx)
        cell.fill = fill
    finish = ws.cell(row, DAY_COL_START + last_idx, end.day)
    finish.font = Font(name="微软雅黑", size=7, color="FFFFFF", bold=True)
    finish.alignment = Alignment(horizontal="center", vertical="center")
    finish.number_format = "0"


def _add_gantt_sheet(wb) -> None:
    """Calendar-day Gantt: equal-width columns, bar length = days from project start to finish."""
    ws = wb.create_sheet("甘特图")
    days = _calendar_days()
    start, end = days[0], days[-1]
    last_day_col = DAY_COL_START + len(days) - 1
    thin = _thin_border()
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(name="微软雅黑", bold=True, color="FFFFFF", size=10)
    month_fill = PatternFill("solid", fgColor="2E75B6")
    planned_bar = "5B8BD5"
    actual_bar = "F59E0B"
    delay_bar = "DC2626"
    early_bar = "16A34A"
    pending_fill = PatternFill("solid", fgColor="F8FAFC")
    alt_label = PatternFill("solid", fgColor="F8FAFC")
    body_font = Font(name="微软雅黑", size=9)
    small = Font(name="微软雅黑", size=8, color="334155")

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=6)
    title = ws.cell(
        1,
        1,
        f"PAB 甘特图（{start.isoformat()} → {end.isoformat()}，每格 1 天；条长 = 日历天数）",
    )
    title.font = Font(name="微软雅黑", bold=True, size=12, color="1F4E78")
    title.alignment = left
    ws.row_dimensions[1].height = 22

    # Month bands above the day axis.
    month_start = 0
    for idx, day in enumerate(days):
        nxt = days[idx + 1] if idx + 1 < len(days) else None
        if nxt is None or nxt.month != day.month:
            col1 = DAY_COL_START + month_start
            col2 = DAY_COL_START + idx
            if col2 > col1:
                ws.merge_cells(start_row=1, start_column=col1, end_row=1, end_column=col2)
            cell = ws.cell(1, col1, f"{day.year}年{day.month}月")
            cell.fill = month_fill
            cell.font = header_font
            cell.alignment = center
            for col in range(col1, col2 + 1):
                painted = ws.cell(1, col)
                painted.fill = month_fill
                painted.border = thin
            month_start = idx + 1

    labels = ("类型", "阶段", "任务", "预估完成", "实际完成", "偏移")
    for col, name in enumerate(labels, start=1):
        cell = ws.cell(2, col, name)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center
        cell.border = thin

    for idx, day in enumerate(days):
        col = DAY_COL_START + idx
        cell = ws.cell(2, col, day)
        cell.number_format = "D"
        cell.font = Font(name="微软雅黑", size=7, color="FFFFFF" if day.weekday() < 5 else "FDE68A")
        cell.fill = header_fill if day.weekday() < 5 else PatternFill("solid", fgColor="334155")
        cell.alignment = center
        cell.border = thin
        ws.column_dimensions[get_column_letter(col)].width = 2.6
    ws.row_dimensions[2].height = 18

    for idx, width in enumerate((8, 8, 22, 12, 12, 10), start=1):
        ws.column_dimensions[get_column_letter(idx)].width = width

    row = 3
    for task, _summary, planned, actual, _reason, _owner in ROWS:
        brief = brief_for_task(task)
        phase = phase_for_task(task)
        offset = _offset_label(planned, actual)
        for kind in ("预估", "实际"):
            values = (
                kind,
                phase,
                brief,
                planned,
                actual if actual is not None else "待填",
                offset if kind == "实际" else "",
            )
            for col, value in enumerate(values, start=1):
                cell = ws.cell(row, col, value)
                cell.font = body_font
                cell.border = thin
                cell.alignment = left if col == 3 else center
                if col in (4, 5) and isinstance(value, date):
                    cell.number_format = "YYYY-MM-DD"
                if row % 2 == 0:
                    cell.fill = alt_label
            finish = planned if kind == "预估" else actual
            for idx in range(len(days)):
                ws.cell(row, DAY_COL_START + idx).border = thin
            if finish is not None:
                color = planned_bar
                if kind == "实际":
                    delta = (actual - planned).days
                    if delta > 0:
                        color = delay_bar
                    elif delta < 0:
                        color = early_bar
                    else:
                        color = actual_bar
                _fill_bar(ws, row, start, finish, color)
            else:
                pending = ws.cell(row, DAY_COL_START, "待填（尚无实际完成日期）")
                pending.font = small
                pending.alignment = left
                pending.fill = pending_fill
                ws.merge_cells(
                    start_row=row,
                    start_column=DAY_COL_START,
                    end_row=row,
                    end_column=min(DAY_COL_START + 14, last_day_col),
                )
            if kind == "实际" and actual is not None:
                off_cell = ws.cell(row, 6)
                if offset.startswith("+"):
                    off_cell.font = Font(name="微软雅黑", size=9, color="B91C1C", bold=True)
                elif offset.startswith("-"):
                    off_cell.font = Font(name="微软雅黑", size=9, color="15803D", bold=True)
            ws.row_dimensions[row].height = 16
            row += 1

    legend_row = row + 1
    ws.merge_cells(start_row=legend_row, start_column=1, end_row=legend_row, end_column=6)
    legend = ws.cell(
        legend_row,
        1,
        "图例：蓝条=预估（从项目首日铺到预估完成日，格数=日历天数）；"
        "橙/红/绿=实际（准时/延期/提前）。未填实际完成节点时显示「待填」。深色列为周末。",
    )
    legend.font = Font(name="微软雅黑", size=9, color="6B7280")
    legend.alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[legend_row].height = 32
    samples = (
        (planned_bar, "预估"),
        (actual_bar, "实际·准时"),
        (delay_bar, "实际·延期"),
        (early_bar, "实际·提前"),
    )
    for offset, (color, text) in enumerate(samples):
        cell = ws.cell(legend_row, DAY_COL_START + offset * 4, text)
        cell.fill = PatternFill("solid", fgColor=color)
        cell.font = Font(name="微软雅黑", size=8, color="FFFFFF", bold=True)
        cell.alignment = center

    ws.freeze_panes = "G3"
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToPage = False
    ws.page_setup.paperSize = ws.PAPERSIZE_A3
    ws.print_title_rows = "1:2"
    ws.print_title_cols = "A:F"
    ws.sheet_view.zoomScale = 90


def _merge_consecutive(ws, col: int, labels: list[str], fill, font, align, border) -> None:
    sentinel = object()
    start = 2
    current = labels[0]
    for idx, label in enumerate(labels[1:] + [sentinel], start=3):
        if label == current:
            continue
        end = idx - 1
        if end > start:
            ws.merge_cells(start_row=start, start_column=col, end_row=end, end_column=col)
        fill_value = fill[current] if isinstance(fill, dict) else fill
        cell = ws.cell(start, col, current)
        cell.fill = fill_value
        cell.font = font
        cell.alignment = align
        for row in range(start, end + 1):
            painted = ws.cell(row, col)
            painted.border = border
            painted.fill = fill_value
            painted.font = font
            painted.alignment = align
        if label is sentinel:
            break
        start = idx
        current = label


def build_workbook(output_path: Path | None = None) -> Path:
    wb = Workbook()
    ws = wb.active
    ws.title = "时间节点计划表"

    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(name="微软雅黑", bold=True, color="FFFFFF", size=11)
    body_font = Font(name="微软雅黑", size=10)
    wrap = Alignment(wrap_text=True, vertical="center")
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin = _thin_border()
    alt_fill = PatternFill("solid", fgColor="F3F7FA")
    overtime_fill = PatternFill("solid", fgColor="FF0000")
    early_fill = PatternFill("solid", fgColor="86EFAC")
    group_fill = PatternFill("solid", fgColor="E2E8F0")
    group_font = Font(name="微软雅黑", bold=True, size=10, color="1F4E78")

    for col, header in enumerate(HEADERS, start=1):
        cell = ws.cell(1, col, header)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center
        cell.border = thin

    last_data_row = 1 + len(ROWS)
    groups: list[str] = []
    for i, (task, summary, planned, actual, reason, owner) in enumerate(ROWS, start=2):
        groups.append(group_for_task(task))
        values = (
            phase_for_task(task),
            group_for_task(task),
            brief_for_task(task),
            summary,
            planned,
            actual,
            reason,
            owner,
            f'=IF(OR(E{i}="",F{i}=""),"待填",F{i}-E{i})',
        )
        for col, value in enumerate(values, start=1):
            cell = ws.cell(i, col, value)
            cell.font = body_font
            cell.border = thin
            cell.alignment = wrap if col in (3, 4, 7) else center
            if i % 2 == 0 and col not in (1, 2):
                cell.fill = alt_fill
            if col in (5, 6) and isinstance(value, date):
                cell.number_format = "YYYY-MM-DD"
            if col == 9:
                cell.number_format = "0"

    phase_fills = {
        "第一期": PatternFill("solid", fgColor="D6EAF8"),
        "第二期": PatternFill("solid", fgColor="FDEBD0"),
    }
    phase_font = Font(name="微软雅黑", bold=True, size=12, color="1F4E78")
    phase_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    _merge_consecutive(
        ws,
        1,
        [phase_for_task(row[0]) for row in ROWS],
        phase_fills,
        phase_font,
        phase_align,
        thin,
    )
    _merge_consecutive(ws, 2, groups, group_fill, group_font, phase_align, thin)

    ws.conditional_formatting.add(
        f"F2:F{max(last_data_row, 200)}",
        FormulaRule(
            formula=["AND(ISNUMBER(F2),ISNUMBER(E2),F2>E2)"],
            fill=overtime_fill,
            font=Font(name="微软雅黑", color="FFFFFF", bold=True, size=10),
        ),
    )
    ws.conditional_formatting.add(
        f"I2:I{last_data_row}",
        FormulaRule(formula=['AND(ISNUMBER(I2),I2>0)'], fill=overtime_fill, font=Font(name="微软雅黑", color="FFFFFF", size=10)),
    )
    ws.conditional_formatting.add(
        f"I2:I{last_data_row}",
        FormulaRule(formula=['AND(ISNUMBER(I2),I2<0)'], fill=early_fill, font=Font(name="微软雅黑", color="14532D", size=10)),
    )

    widths = (10, 12, 24, 62, 16, 16, 16, 12, 12)
    for idx, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(idx)].width = width
    ws.row_dimensions[1].height = 22
    for row in range(2, last_data_row + 1):
        ws.row_dimensions[row].height = 32

    ws.freeze_panes = "C2"
    ws.auto_filter.ref = f"A1:I{last_data_row}"

    note_row = last_data_row + 2
    ws.merge_cells(start_row=note_row, start_column=1, end_row=note_row, end_column=9)
    note = ws.cell(
        note_row,
        1,
        "说明：A列按第一期/第二期合并，B列按任务组合并。偏移天数=实际−预估（正值延期）。"
        "进度条见工作表「甘特图」：横轴每一格为 1 个自然日，色条长度等于从项目起始日到完成节点的日历天数。",
    )
    note.font = Font(name="微软雅黑", size=9, color="6B7280")
    note.alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[note_row].height = 40

    date_dv = DataValidation(type="date", operator="greaterThan", formula1="DATE(2020,1,1)", allow_blank=True)
    date_dv.add(f"E2:F{max(last_data_row, 200)}")
    ws.add_data_validation(date_dv)

    _add_gantt_sheet(wb)

    target = output_path or (_project_root() / OUTPUT_NAME)
    wb.save(target)
    return target



if __name__ == "__main__":
    path = build_workbook()
    print(f"已生成：{path}")
