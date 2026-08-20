"""Build PAB product schedule XLSX from the extracted Word plan."""

from __future__ import annotations

from datetime import date
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


def _unique_dates() -> list[date]:
    dates = [row[2] for row in ROWS]
    dates.extend(row[3] for row in ROWS if row[3] is not None)
    return sorted(set(dates))


def _offset_label(planned: date, actual: date | None) -> str:
    if actual is None:
        return "待填"
    delta = (actual - planned).days
    if delta > 0:
        return f"+{delta}天"
    if delta < 0:
        return f"{delta}天"
    return "0天"


def _points_by_date(*, use_actual: bool) -> dict[date, list[tuple[str, str]]]:
    grouped: dict[date, list[tuple[str, str]]] = {}
    for task, _summary, planned, actual, _reason, _owner in ROWS:
        when = actual if use_actual else planned
        if when is None:
            continue
        grouped.setdefault(when, []).append((brief_for_task(task), _offset_label(planned, actual)))
    return grouped


def _draw_segment_axis(
    ws,
    start_row: int,
    *,
    title: str,
    subtitle: str,
    dates: list[date],
    points: dict[date, list[tuple[str, str]]],
    show_offset: bool,
    bar_color: str,
) -> int:
    """Draw a wrapping left-to-right axis using columns A–I (one unique date per cell)."""
    cols = 9
    title_font = Font(name="微软雅黑", bold=True, size=12, color="1F4E78")
    sub_font = Font(name="微软雅黑", size=9, color="6B7280")
    label_font = Font(name="微软雅黑", size=8, color="1F2937")
    date_font = Font(name="微软雅黑", size=8, color="334155")
    bar_fill = PatternFill("solid", fgColor=bar_color)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)

    ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=cols)
    ws.cell(start_row, 1, title).font = title_font
    ws.merge_cells(start_row=start_row + 1, start_column=1, end_row=start_row + 1, end_column=cols)
    ws.cell(start_row + 1, 1, subtitle).font = sub_font
    ws.cell(start_row + 1, 1).alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[start_row + 1].height = 28

    row = start_row + 2
    total = len(dates)
    for chunk_start in range(0, total, cols):
        chunk = dates[chunk_start : chunk_start + cols]
        last_chunk = chunk_start + cols >= total
        stack = max((len(points.get(day, [])) for day in chunk), default=1)
        label_top = row
        for offset, day in enumerate(chunk):
            col = 1 + offset
            for idx, (name, _off) in enumerate(points.get(day, [])):
                cell = ws.cell(label_top + idx, col, name)
                cell.font = label_font
                cell.alignment = Alignment(horizontal="center", vertical="bottom", wrap_text=True)
        tick_row = label_top + stack
        axis_row = tick_row + 1
        date_row = axis_row + 1
        offset_row = date_row + 1 if show_offset else date_row
        for offset, day in enumerate(chunk):
            col = 1 + offset
            tick = ws.cell(tick_row, col, "▼" if day in points else "·")
            tick.font = Font(name="微软雅黑", size=9, color=bar_color)
            tick.alignment = center
            last_cell = last_chunk and offset == len(chunk) - 1
            bar = ws.cell(axis_row, col, "▶" if last_cell else "━")
            bar.fill = bar_fill
            bar.font = Font(name="微软雅黑", color="FFFFFF", bold=True, size=11)
            bar.alignment = center
            date_cell = ws.cell(date_row, col, day)
            date_cell.number_format = "YYYY-MM-DD"
            date_cell.font = date_font
            date_cell.alignment = center
            if show_offset:
                labels = [off for _name, off in points.get(day, [])]
                text = " / ".join(labels) if labels else "待填"
                off_cell = ws.cell(offset_row, col, text)
                if text.startswith("+"):
                    off_cell.font = Font(name="微软雅黑", size=8, color="B91C1C", bold=True)
                elif text.startswith("-"):
                    off_cell.font = Font(name="微软雅黑", size=8, color="15803D", bold=True)
                else:
                    off_cell.font = Font(name="微软雅黑", size=8, color="6B7280")
                off_cell.alignment = center
        for r in range(label_top, offset_row + 1):
            ws.row_dimensions[r].height = 40 if r < tick_row else 18
        row = offset_row + 2
    return row


def _add_main_sheet_timelines(ws, start_row: int, last_data_row: int) -> None:
    dates = _unique_dates()
    start, end = dates[0], dates[-1]
    after_planned = _draw_segment_axis(
        ws,
        start_row,
        title="目标时间轴（按预估完成节点）",
        subtitle=(
            f"左端为首个任务 {start.isoformat()}，右端为最终任务 {end.isoformat()}。"
            "按不重复日期从左到右分段（每行最多 9 个节点，下一行续接）；同一天多个任务上下叠放。"
        ),
        dates=dates,
        points=_points_by_date(use_actual=False),
        show_offset=False,
        bar_color="1F4E78",
    )
    _draw_segment_axis(
        ws,
        after_planned + 1,
        title="实际时间轴（按实际完成节点）",
        subtitle=(
            "与目标轴共用同一组日期刻度，便于对照偏移。偏移 = 实际 − 预估（正值延期）。"
            "主表「实际完成节点」为空时，刻度保留并显示「待填」。"
        ),
        dates=dates,
        points=_points_by_date(use_actual=True),
        show_offset=True,
        bar_color="C2410C",
    )
    _ = last_data_row


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
        "下方为目标时间轴与实际时间轴：按不重复日期从左到右分段，同一刻度对照偏移。",
    )
    note.font = Font(name="微软雅黑", size=9, color="6B7280")
    note.alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[note_row].height = 40

    date_dv = DataValidation(type="date", operator="greaterThan", formula1="DATE(2020,1,1)", allow_blank=True)
    date_dv.add(f"E2:F{max(last_data_row, 200)}")
    ws.add_data_validation(date_dv)

    _add_main_sheet_timelines(ws, note_row + 2, last_data_row)

    target = output_path or (_project_root() / OUTPUT_NAME)
    wb.save(target)
    return target



if __name__ == "__main__":
    path = build_workbook()
    print(f"已生成：{path}")
