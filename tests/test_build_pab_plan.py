"""Tests for the PAB schedule workbook."""

from __future__ import annotations

from datetime import date
from pathlib import Path

from openpyxl import load_workbook

from scripts.build_pab_plan import HEADERS, build_workbook


def test_headers_and_overdue_rule(tmp_path: Path) -> None:
    path = build_workbook(tmp_path / "PAB产品计划表.xlsx")
    wb = load_workbook(path)
    ws = wb.active

    assert [ws.cell(1, col).value for col in range(1, 8)] == list(HEADERS)
    assert ws["A1"].value == "阶段"
    assert ws["B1"].value == "任务"
    assert ws["C1"].value == "任务摘要"
    assert ws["D1"].value == "预估完成节点"
    assert ws["E1"].value == "实际完成节点"
    assert ws["F1"].value == "超时原因"
    assert ws["G1"].value == "任务责任人"

    tasks = {ws.cell(row, 2).value: row for row in range(2, ws.max_row + 1) if ws.cell(row, 2).value}
    assert "PAB060·零件到位" in tasks
    row = tasks["PAB060·零件到位"]
    assert ws.cell(row, 4).value.date() == date(2026, 9, 24)
    assert ws.cell(row, 5).value is None
    assert ws.cell(row, 7).value == "王杰"
    assert "协助人：吴晨阳" in ws.cell(row, 3).value

    row_180 = tasks["PAB180·零件到位"]
    assert ws.cell(row_180, 4).value.date() == date(2026, 10, 15)
    assert "国庆假期" in ws.cell(row_180, 3).value

    formulas = []
    for cf_range in ws.conditional_formatting._cf_rules:
        assert "E2" in str(cf_range)
        for rule in ws.conditional_formatting._cf_rules[cf_range]:
            formulas.extend(rule.formula)
    assert any("E2>D2" in formula for formula in formulas)


def test_phase_column_is_merged(tmp_path: Path) -> None:
    path = build_workbook(tmp_path / "plan.xlsx")
    ws = load_workbook(path).active
    merged = {str(range_) for range_ in ws.merged_cells.ranges}
    assert "A2:A30" in merged
    assert "A31:A45" in merged
    assert ws["A2"].value == "第一期"
    assert ws["A31"].value == "第二期"


def test_source_covers_both_phases(tmp_path: Path) -> None:
    path = build_workbook(tmp_path / "plan.xlsx")
    ws = load_workbook(path).active
    names = [
        ws.cell(row, 2).value
        for row in range(2, ws.max_row + 1)
        if ws.cell(row, 2).value and not str(ws.cell(row, 2).value).startswith("说明")
    ]
    for model in ("PAB060", "PAB090", "PAB115", "PAB142", "PAB180", "PAB220"):
        assert any(name.startswith(f"{model}·") for name in names)
    assert names[0] == "第一期·零件图纸整理归档"
    assert names[-1] == "PAB220·双级样机"
