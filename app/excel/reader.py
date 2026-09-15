"""Leitor de pastas de trabalho Excel para snapshots lógicos."""

from __future__ import annotations

from datetime import date, datetime, time, timedelta
from decimal import Decimal
from pathlib import Path
from typing import TypeAlias

from openpyxl import load_workbook


CellValue: TypeAlias = (
    str | int | float | bool | date | datetime | time | timedelta | Decimal
)
SheetSnapshot: TypeAlias = dict[str, CellValue]
Snapshot: TypeAlias = dict[str, SheetSnapshot]


def is_empty(value: object) -> bool:
    """Informa se um valor representa uma célula vazia para a auditoria."""
    return value is None or value == ""


def read_snapshot(path: str | Path) -> Snapshot:
    """Lê um ``.xlsx`` sem alterá-lo e preserva fórmulas no snapshot."""
    workbook_path = Path(path)
    if workbook_path.suffix.lower() != ".xlsx":
        raise ValueError("O leitor aceita somente arquivos .xlsx.")

    workbook = load_workbook(
        filename=workbook_path,
        read_only=True,
        data_only=False,
        keep_links=False,
    )
    try:
        snapshot: Snapshot = {}
        for worksheet in workbook.worksheets:
            cells: SheetSnapshot = {}
            for row in worksheet.iter_rows():
                for cell in row:
                    if not is_empty(cell.value):
                        cells[cell.coordinate] = cell.value
            snapshot[worksheet.title] = cells
        return snapshot
    finally:
        workbook.close()
