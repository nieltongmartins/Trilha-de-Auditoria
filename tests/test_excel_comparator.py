from pathlib import Path

from app.excel.comparator import CellChange, compare_snapshots
from app.excel.reader import read_snapshot
from app.models import ChangeType


def test_comparison_detects_required_changes_and_preserves_sheet(
    local_workbook_versions: Path,
) -> None:
    previous = read_snapshot(local_workbook_versions / "0.84.xlsx")
    current = read_snapshot(local_workbook_versions / "0.85.xlsx")

    assert compare_snapshots(previous, current) == [
        CellChange("Apoio", "A1", ChangeType.MOD, "Original", "Atualizado"),
        CellChange("Dados", "A1", ChangeType.MOD, 10, 15),
        CellChange("Dados", "B2", ChangeType.ADD, None, "OK"),
        CellChange("Dados", "C3", ChangeType.DEL, "Pendente", None),
        CellChange("Dados", "D4", ChangeType.ADD, None, 0),
        CellChange("Dados", "F6", ChangeType.MOD, "=SUM(A1:A10)", "=SUM(A1:A20)"),
    ]


def test_false_and_zero_are_not_treated_as_empty_or_equal(
    local_workbook_versions: Path,
) -> None:
    previous = read_snapshot(local_workbook_versions / "0.85.xlsx")
    current = read_snapshot(local_workbook_versions / "0.86.xlsx")

    assert compare_snapshots(previous, current) == [
        CellChange("Dados", "D4", ChangeType.MOD, 0, False),
        CellChange("Dados", "E5", ChangeType.DEL, False, None),
    ]


def test_added_and_removed_sheets_are_cell_changes(
    local_workbook_versions: Path,
) -> None:
    previous = read_snapshot(local_workbook_versions / "0.86.xlsx")
    current = read_snapshot(local_workbook_versions / "0.87.xlsx")

    assert compare_snapshots(previous, current) == [
        CellChange("Apoio", "A1", ChangeType.DEL, "Atualizado", None),
        CellChange("Nova Aba", "B2", ChangeType.ADD, None, "Nova"),
    ]


def test_equal_snapshots_have_no_changes(
    local_workbook_versions: Path,
) -> None:
    snapshot = read_snapshot(local_workbook_versions / "0.87.xlsx")

    assert compare_snapshots(snapshot, snapshot) == []


def test_result_is_deterministic_by_sheet_row_and_column() -> None:
    previous = {"Z": {"B10": "x", "C2": "x"}, "A": {"Z1": "x"}}
    current = {"Z": {"A10": "x"}, "A": {}}

    first_result = compare_snapshots(previous, current)

    assert first_result == compare_snapshots(previous, current)
    assert [(item.sheet, item.address) for item in first_result] == [
        ("A", "Z1"),
        ("Z", "C2"),
        ("Z", "A10"),
        ("Z", "B10"),
    ]
