"""Tests for the PAB schedule workbook."""

from __future__ import annotations

from datetime import date
from pathlib import Path

from openpyxl import load_workbook

from scripts.build_pab_plan import DATE_FORMAT, HEADERS, build_workbook


def test_headers_and_overdue_rule(tmp_path: Path) -> None:
    path = build_workbook(tmp_path / "PAB产品计划表.xlsx")
    wb = load_workbook(path)
    ws = wb.active

    assert [ws.cell(1, col).value for col in range(1, 10)] == list(HEADERS)
    assert ws["A1"].value == "阶段"
    assert ws["B1"].value == "任务"
    assert ws["C1"].value == "任务简要"
    assert ws["D1"].value == "任务摘要"
    assert ws["E1"].value == "预估完成节点"
    assert ws["F1"].value == "实际完成节点"
    assert ws["G1"].value == "超时原因"
    assert ws["H1"].value == "任务责任人"
    assert ws["I1"].value == "偏移天数"

    tasks = {ws.cell(row, 3).value: row for row in range(2, 46) if ws.cell(row, 3).value}
    assert "PAB060·零件到位" in tasks
    row = tasks["PAB060·零件到位"]
    assert ws.cell(row, 5).value.date() == date(2026, 9, 24)
    assert ws.cell(row, 6).value is None
    assert ws.cell(row, 8).value == "王杰"
    assert "协助人：吴晨阳" in ws.cell(row, 4).value
    assert "待填" in str(ws.cell(row, 9).value)

    assert ws.cell(row, 5).number_format.lower() == DATE_FORMAT
    assert ws.cell(row, 6).number_format.lower() == DATE_FORMAT

    row_180 = tasks["PAB180·零件到位"]
    assert ws.cell(row_180, 5).value.date() == date(2026, 10, 15)
    assert "国庆假期" in ws.cell(row_180, 4).value

    formulas = []
    font_rgbs = []
    for cf_range in ws.conditional_formatting._cf_rules:
        for rule in ws.conditional_formatting._cf_rules[cf_range]:
            formulas.extend(rule.formula)
            color = getattr(getattr(getattr(rule, "dxf", None), "font", None), "color", None)
            rgb = getattr(color, "rgb", None) if color is not None else None
            if rgb:
                font_rgbs.append(str(rgb).upper())
    assert any("F2>E2" in formula for formula in formulas)
    assert any("ISNUMBER(E2)" in formula for formula in formulas)
    assert all(not rgb.endswith("FFFFFF") for rgb in font_rgbs)


def test_phase_column_is_merged(tmp_path: Path) -> None:
    path = build_workbook(tmp_path / "plan.xlsx")
    ws = load_workbook(path).active
    merged = {str(range_) for range_ in ws.merged_cells.ranges}
    assert "A2:A30" in merged
    assert "A31:A45" in merged
    assert "B2:B6" in merged
    assert "B7:B12" in merged
    assert ws["A2"].value == "第一期"
    assert ws["A31"].value == "第二期"
    assert ws["B2"].value == "前期准备"
    assert ws["B7"].value == "PAB060"


def test_source_covers_both_phases(tmp_path: Path) -> None:
    path = build_workbook(tmp_path / "plan.xlsx")
    ws = load_workbook(path).active
    names = [ws.cell(row, 3).value for row in range(2, 46) if ws.cell(row, 3).value]
    for model in ("PAB060", "PAB090", "PAB115", "PAB142", "PAB180", "PAB220"):
        assert any(str(name).startswith(f"{model}·") for name in names)
    assert names[0] == "一期零件图纸整理归档"
    assert names[-1] == "PAB220·双级样机"


def _cell_date(value):
    if value is None:
        return None
    if isinstance(value, date) and not hasattr(value, "hour"):
        return value
    if hasattr(value, "date"):
        return value.date()
    return None


def test_gantt_sheet_uses_calendar_day_bars(tmp_path: Path) -> None:
    wb = load_workbook(build_workbook(tmp_path / "plan.xlsx"))
    assert "甘特图" in wb.sheetnames
    ws = wb["甘特图"]
    # Header row 2 contains consecutive calendar dates.
    start = _cell_date(ws.cell(2, 7).value)
    assert start == date(2026, 8, 24)
    dates = []
    col = 7
    while True:
        d = _cell_date(ws.cell(2, col).value)
        if d is None:
            break
        dates.append(d)
        col += 1
    assert dates[-1] == date(2026, 10, 23)
    assert dates == [start + __import__("datetime").timedelta(days=i) for i in range(len(dates))]

    # Bar length for the last task equals calendar span from project start.
    last_task_row = None
    for row in range(3, ws.max_row + 1):
        if ws.cell(row, 3).value == "PAB220·双级样机" and ws.cell(row, 1).value == "预估":
            last_task_row = row
            break
    assert last_task_row is not None
    # Dates come from the main sheet; each day cell has a formula that draws the bar.
    assert "E45" in str(ws.cell(last_task_row, 4).value)
    assert "F45" in str(ws.cell(last_task_row, 5).value)
    assert ws.cell(last_task_row, 4).number_format.lower() == DATE_FORMAT
    assert ws.cell(last_task_row, 5).number_format.lower() == DATE_FORMAT
    first_bar = str(ws.cell(last_task_row, 7).value)
    last_bar = str(ws.cell(last_task_row, 6 + len(dates)).value)
    assert first_bar.startswith("=")
    assert '"#"' not in first_bar
    assert "CHAR(" not in first_bar
    assert "$D" in first_bar
    assert "N(G$2)" in first_bar
    assert first_bar.endswith('G$2,"")')
    assert "$D" in last_bar
    assert last_bar.endswith('$2,"")')
    assert ws.cell(last_task_row, 7).number_format.lower() == DATE_FORMAT
    actual_bar = str(ws.cell(last_task_row + 1, 7).value)
    assert actual_bar.startswith("=")
    assert "$E" in actual_bar
    formulas = []
    font_rgbs = []
    for cf_range in ws.conditional_formatting._cf_rules:
        for rule in ws.conditional_formatting._cf_rules[cf_range]:
            formulas.extend(rule.formula)
            color = getattr(getattr(getattr(rule, "dxf", None), "font", None), "color", None)
            rgb = getattr(color, "rgb", None) if color is not None else None
            if rgb:
                font_rgbs.append(str(rgb).upper())
    assert any("ISNUMBER(G3)" in formula and "预估" in formula for formula in formulas)
    assert any("ISNUMBER(G3)" in formula and "实际" in formula for formula in formulas)
    assert all(not rgb.endswith("FFFFFF") for rgb in font_rgbs)
    assert (date(2026, 10, 23) - date(2026, 8, 24)).days + 1 == len(dates)


def test_main_sheet_points_to_gantt(tmp_path: Path) -> None:
    ws = load_workbook(build_workbook(tmp_path / "plan.xlsx")).active
    note = " ".join(str(c.value) for row in ws.iter_rows(min_row=46, max_row=50) for c in row if c.value)
    assert "甘特图" in note
    assert "手工涂色" in note
    assert not any(c.value == "▶" for row in ws.iter_rows(min_row=47, max_row=90) for c in row)

