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


def _to_date(value: date | None) -> date | None:
    if value is None:
        return None
    if isinstance(value, date) and not hasattr(value, "hour"):
        return value
    if hasattr(value, "date"):
        return value.date()
    return value


def _axis_column(value: date, start: date, end: date, first_col: int, n_cols: int) -> int:
    span = (end - start).days
    if span <= 0:
        return first_col
    offset = round((value - start).days / span * (n_cols - 1))
    return first_col + max(0, min(n_cols - 1, offset))


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


def _draw_arrow_axis(ws, *, title_row: int, title: str, subtitle: str, points: list[tuple[date, str, str]], start: date, end: date, axis_color: str, show_offset: bool) -> int:
    """Draw a left-to-right dated axis. Returns the last row used."""
    first_col = 2
    n_cols = (end - start).days + 1
    last_col = first_col + n_cols - 1
    title_font = Font(name="微软雅黑", bold=True, size=12, color="1F4E78")
    sub_font = Font(name="微软雅黑", size=9, color="6B7280")
    label_font = Font(name="微软雅黑", size=8, color="1F2937")
    date_font = Font(name="微软雅黑", size=7, color="334155")
    axis_fill = PatternFill("solid", fgColor=axis_color)
    ws.merge_cells(start_row=title_row, start_column=1, end_row=title_row, end_column=min(8, last_col))
    ws.cell(title_row, 1, title).font = title_font
    ws.merge_cells(start_row=title_row + 1, start_column=1, end_row=title_row + 1, end_column=min(8, last_col))
    ws.cell(title_row + 1, 1, subtitle).font = sub_font

    buckets: dict[int, list[tuple[str, str]]] = {}
    for when, name, offset_text in points:
        col = _axis_column(when, start, end, first_col, n_cols)
        buckets.setdefault(col, []).append((name, offset_text))
    stack = max((len(items) for items in buckets.values()), default=1)
    label_top = title_row + 2
    for col, items in buckets.items():
        for idx, (name, offset_text) in enumerate(items):
            cell = ws.cell(label_top + idx, col, name)
            cell.font = label_font
            cell.alignment = Alignment(horizontal="center", vertical="bottom", wrap_text=True, textRotation=90)
            if show_offset:
                off = ws.cell(label_top + stack + idx, col, offset_text)
                if offset_text.startswith("+"):
                    off.font = Font(name="微软雅黑", size=8, color="B91C1C", bold=True)
                elif offset_text.startswith("-"):
                    off.font = Font(name="微软雅黑", size=8, color="15803D", bold=True)
                else:
                    off.font = Font(name="微软雅黑", size=8, color="6B7280")
                off.alignment = Alignment(horizontal="center", vertical="center", textRotation=90)

    tick_row = label_top + stack + (stack if show_offset else 0)
    axis_row = tick_row + 1
    date_row = axis_row + 1
    for col in range(first_col, last_col + 1):
        day = start + timedelta(days=col - first_col)
        tick = ws.cell(tick_row, col, "▼" if col in buckets or day in {start, end} else "·")
        tick.font = Font(name="微软雅黑", size=8, color=axis_color)
        tick.alignment = Alignment(horizontal="center")
        bar = ws.cell(axis_row, col, "▶" if col == last_col else "━")
        bar.fill = axis_fill
        bar.font = Font(name="微软雅黑", color="FFFFFF", bold=True, size=10)
        bar.alignment = Alignment(horizontal="center", vertical="center")
        if col == first_col or col == last_col or day.day in {1, 15} or col in buckets:
            date_cell = ws.cell(date_row, col, day)
            date_cell.number_format = "M/D"
            date_cell.font = date_font
            date_cell.alignment = Alignment(horizontal="center", textRotation=90)
        ws.column_dimensions[get_column_letter(col)].width = 3.2
    ws.cell(axis_row, 1, f"{start.isoformat()} →").font = Font(name="微软雅黑", size=8, color="1F4E78")
    ws.cell(axis_row, last_col + 1, start.isoformat() if False else f"→ {end.isoformat()}").font = Font(name="微软雅黑", size=8, color="1F4E78")
    ws.column_dimensions[get_column_letter(last_col + 1)].width = 14
    for row in range(label_top, date_row + 1):
        ws.row_dimensions[row].height = 48 if row < tick_row else 18
    return date_row


def _add_timeline_sheet(wb: Workbook) -> None:
    from datetime import timedelta

    ws = wb.create_sheet("时间轴")
    planned_points = [(row[2], brief_for_task(row[0]), "目标") for row in ROWS]
    actual_points: list[tuple[date, str, str]] = []
    for task, _summary, planned, actual, _reason, _owner in ROWS:
        brief = brief_for_task(task)
        if actual is None:
            continue
        delta = (actual - planned).days
        if delta > 0:
            offset = f"+{delta}天"
        elif delta < 0:
            offset = f"{delta}天"
        else:
            offset = "0天"
        actual_points.append((actual, brief, offset))

    start = min(row[2] for row in ROWS)
    end = max(row[2] for row in ROWS)
    if actual_points:
        start = min(start, min(item[0] for item in actual_points))
        end = max(end, max(item[0] for item in actual_points))

    last = _draw_arrow_axis(
        ws,
        title_row=1,
        title="目标时间轴（按预估完成节点）",
        subtitle=f"左端为首个任务 {start.isoformat()}，右端为最终任务 {end.isoformat()}，其余任务按日历天数分段落在轴上。",
        points=planned_points,
        start=start,
        end=end,
        axis_color="1F4E78",
        show_offset=False,
    )
    actual_title = "实际时间轴（按实际完成节点；偏移度 = 实际 − 预估，正值延期）"
    actual_sub = (
        "尚无实际完成节点。填写主表「实际完成节点」后重新生成本表即可刷新本轴。"
        if not actual_points
        else "节点位置按实际日期落轴；轴下旋转文字为相对预估的偏移天数。"
    )
    _draw_arrow_axis(
        ws,
        title_row=last + 2,
        title=actual_title,
        subtitle=actual_sub,
        points=actual_points,
        start=start,
        end=end,
        axis_color="C2410C",
        show_offset=True,
    )
    legend_row = last + 2
    # legend placed after second axis by scanning used rows
    legend_row = ws.max_row + 2
    ws.cell(legend_row, 1, "偏移图例：+N天=延期（红）；-N天=提前（绿）；0天=按期；待填=主表实际完成节点为空。两轴共用同一起止日期，便于对照时间偏移。").font = Font(
        name="微软雅黑", size=9, color="6B7280"
    )
    ws.column_dimensions["A"].width = 22
    ws.freeze_panes = "B3"
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1


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
        "下方为时间轴图；同文件工作表「时间轴」为按日历天数展开的带箭头轴（目标轴在上、实际轴在下，两轴共用起止日期以对照偏移）。",
    )
    note.font = Font(name="微软雅黑", size=9, color="6B7280")
    note.alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[note_row].height = 40

    date_dv = DataValidation(type="date", operator="greaterThan", formula1="DATE(2020,1,1)", allow_blank=True)
    date_dv.add(f"E2:F{max(last_data_row, 200)}")
    ws.add_data_validation(date_dv)

    _add_main_sheet_timelines(ws, note_row + 2, last_data_row)
    _add_timeline_sheet(wb)

    target = output_path or (_project_root() / OUTPUT_NAME)
    wb.save(target)
    return target


def _add_main_sheet_timelines(ws, start_row: int, last_data_row: int) -> None:
    from openpyxl.chart import ScatterChart, Reference, Series
    from openpyxl.chart.marker import Marker
    from openpyxl.chart.series import SeriesLabel
    from openpyxl.chart.shapes import GraphicalProperties
    from openpyxl.drawing.line import LineProperties
    from openpyxl.chart.axis import DateAxis

    ws.cell(start_row, 1, "目标时间轴 / 实际时间轴").font = Font(name="微软雅黑", bold=True, size=12, color="1F4E78")
    ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=9)
    ws.cell(
        start_row + 1,
        1,
        "下图按预估日期生成从左到右的目标轴；橙色点为实际完成（若已填写）。X 轴左端=首个任务日期，右端=最终任务日期。点相对目标点的左右位移即时间偏移。",
    ).font = Font(name="微软雅黑", size=9, color="6B7280")
    ws.merge_cells(start_row=start_row + 1, start_column=1, end_row=start_row + 1, end_column=9)

    # Chart data block to the right of the table, not printed as primary view.
    data_row0 = start_row + 3
    ws.cell(data_row0, 11, "任务简要")
    ws.cell(data_row0, 12, "预估")
    ws.cell(data_row0, 13, "目标Y")
    ws.cell(data_row0, 14, "实际")
    ws.cell(data_row0, 15, "实际Y")
    ws.cell(data_row0, 16, "轴Y")
    for idx, (task, _summary, planned, actual, _reason, _owner) in enumerate(ROWS, start=1):
        r = data_row0 + idx
        ws.cell(r, 11, brief_for_task(task))
        ws.cell(r, 12, planned).number_format = "YYYY-MM-DD"
        ws.cell(r, 13, 2)
        if actual is not None:
            ws.cell(r, 14, actual).number_format = "YYYY-MM-DD"
        ws.cell(r, 15, 1)
        ws.cell(r, 16, 1.5)
    n = len(ROWS)
    end_r = data_row0 + n

    chart = ScatterChart()
    chart.title = "PAB 目标轴（蓝）与实际轴（橙）"
    chart.x_axis.title = "完成节点"
    chart.y_axis.title = None
    chart.y_axis.scaling.min = 0.5
    chart.y_axis.scaling.max = 2.5
    chart.y_axis.delete = True
    chart.style = 10
    chart.height = 9
    chart.width = 22
    chart.legend.position = "b"

    x_planned = Reference(ws, min_col=12, min_row=data_row0 + 1, max_row=end_r)
    y_planned = Reference(ws, min_col=13, min_row=data_row0, max_row=end_r)
    s_planned = Series(y_planned, x_planned, title="目标时间轴")
    s_planned.marker = Marker(symbol="diamond", size=8)
    s_planned.graphicalProperties = GraphicalProperties(ln=LineProperties(prstDash="solid", w=12000, solidFill="1F4E78"))
    chart.series.append(s_planned)

    x_actual = Reference(ws, min_col=14, min_row=data_row0 + 1, max_row=end_r)
    y_actual = Reference(ws, min_col=15, min_row=data_row0, max_row=end_r)
    s_actual = Series(y_actual, x_actual, title="实际时间轴")
    s_actual.marker = Marker(symbol="circle", size=8)
    s_actual.graphicalProperties = GraphicalProperties(ln=LineProperties(prstDash="dash", w=10000, solidFill="C2410C"))
    chart.series.append(s_actual)

    x_axis_line = Reference(ws, min_col=12, min_row=data_row0 + 1, max_row=end_r)
    y_axis_line = Reference(ws, min_col=16, min_row=data_row0, max_row=end_r)
    s_axis = Series(y_axis_line, x_axis_line, title="时间方向 →")
    s_axis.marker = Marker(symbol="none")
    s_axis.graphicalProperties = GraphicalProperties(ln=LineProperties(solidFill="94A3B8", w=8000))
    chart.series.append(s_axis)

    chart.anchor = f"A{start_row + 3}"
    ws.add_chart(chart)
    for col in range(11, 17):
        ws.column_dimensions[get_column_letter(col)].hidden = True

    _ = (DateAxis, SeriesLabel, last_data_row)


if __name__ == "__main__":
    path = build_workbook()
    print(f"已生成：{path}")
