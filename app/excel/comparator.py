"""Comparação determinística de snapshots de versões Excel."""

from __future__ import annotations

from dataclasses import dataclass

from openpyxl.utils.cell import coordinate_to_tuple

from app.excel.reader import CellValue, Snapshot, is_empty
from app.models import ChangeType


@dataclass(frozen=True)
class CellChange:
    """Diferença observada em uma célula entre duas versões."""

    sheet: str
    address: str
    change_type: ChangeType
    previous_value: CellValue | None
    new_value: CellValue | None


def _values_equal(previous: object, current: object) -> bool:
    """Compara conteúdo e tipo, evitando considerar ``False`` igual a zero."""
    return type(previous) is type(current) and previous == current


def _change_type(previous: object, current: object) -> ChangeType:
    if is_empty(previous):
        return ChangeType.ADD
    if is_empty(current):
        return ChangeType.DEL
    return ChangeType.MOD


def compare_snapshots(previous: Snapshot, current: Snapshot) -> list[CellChange]:
    """Retorna diferenças em ordem estável de aba, linha e coluna."""
    changes: list[CellChange] = []

    for sheet_name in sorted(previous.keys() | current.keys()):
        previous_cells = previous.get(sheet_name, {})
        current_cells = current.get(sheet_name, {})
        addresses = previous_cells.keys() | current_cells.keys()

        for address in sorted(addresses, key=coordinate_to_tuple):
            previous_value = previous_cells.get(address)
            new_value = current_cells.get(address)
            if _values_equal(previous_value, new_value):
                continue
            changes.append(
                CellChange(
                    sheet=sheet_name,
                    address=address,
                    change_type=_change_type(previous_value, new_value),
                    previous_value=previous_value,
                    new_value=new_value,
                )
            )

    return changes
