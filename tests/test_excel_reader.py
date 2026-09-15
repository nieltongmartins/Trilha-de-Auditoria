from pathlib import Path

import pytest

from app.excel.reader import read_snapshot


def test_reader_preserves_values_formulas_and_all_sheets(
    local_workbook_versions: Path,
) -> None:
    snapshot = read_snapshot(local_workbook_versions / "0.84.xlsx")

    assert list(snapshot) == ["Dados", "Apoio"]
    assert snapshot["Dados"] == {
        "A1": 10,
        "C3": "Pendente",
        "E5": False,
        "F6": "=SUM(A1:A10)",
    }
    assert snapshot["Apoio"] == {"A1": "Original"}


def test_reader_does_not_modify_source_file(
    local_workbook_versions: Path,
) -> None:
    path = local_workbook_versions / "0.84.xlsx"
    contents_before = path.read_bytes()

    read_snapshot(path)

    assert path.read_bytes() == contents_before


def test_reader_rejects_non_xlsx_file(tmp_path: Path) -> None:
    path = tmp_path / "version.xls"
    path.touch()

    with pytest.raises(ValueError, match="somente arquivos .xlsx"):
        read_snapshot(path)
