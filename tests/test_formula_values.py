from pathlib import Path

from openpyxl import Workbook
from openpyxl.worksheet.formula import ArrayFormula, DataTableFormula

from app.audit_service import AuditService
from app.excel.comparator import CellChange, compare_snapshots
from app.excel.formulas import normalize_formula_value
from app.excel.reader import read_workbook
from app.models import ChangeType


def array_formula(ref: str = "A1:A2", text: str = "=SUM(B1:B2)") -> ArrayFormula:
    return ArrayFormula(ref=ref, text=text)


def test_normal_formulas_remain_strings_and_compare_by_value() -> None:
    previous = {"Dados": {"A1": "=SUM(B1:B2)"}}

    assert normalize_formula_value("=SUM(B1:B2)") == "=SUM(B1:B2)"
    assert compare_snapshots(previous, {"Dados": {"A1": "=SUM(B1:B2)"}}) == []
    assert compare_snapshots(previous, {"Dados": {"A1": "=SUM(B1:B3)"}}) == [
        CellChange(
            "Dados", "A1", ChangeType.MOD, "=SUM(B1:B2)", "=SUM(B1:B3)"
        )
    ]


def test_distinct_array_formula_instances_compare_by_formula_and_ref() -> None:
    previous = {"Dados": {"A1": array_formula()}}

    assert compare_snapshots(previous, {"Dados": {"A1": array_formula()}}) == []

    changed_formula = compare_snapshots(
        previous, {"Dados": {"A1": array_formula(text="=SUM(B1:B3)")}}
    )
    changed_ref = compare_snapshots(
        previous, {"Dados": {"A1": array_formula(ref="A1:A3")}}
    )

    assert len(changed_formula) == 1
    assert len(changed_ref) == 1
    for change in (*changed_formula, *changed_ref):
        assert change.change_type is ChangeType.MOD
        assert change.previous_value.startswith("ARRAYFORMULA|")
        assert change.new_value.startswith("ARRAYFORMULA|")
        assert "object at 0x" not in change.previous_value + change.new_value


def test_formula_normalization_preserves_regular_value_types() -> None:
    values = (None, 0, False, "")

    normalized = tuple(normalize_formula_value(value) for value in values)

    assert normalized == values
    assert tuple(type(value) for value in normalized) == tuple(
        type(value) for value in values
    )


def test_data_table_formula_preserves_all_openpyxl_metadata() -> None:
    formula = DataTableFormula(
        "C1:D4", ca=True, dt2D=True, dtr=True, r1="A1", r2="B1",
        del1=True, del2=True,
    )

    assert normalize_formula_value(formula) == (
        'DATATABLEFORMULA|{"ca":true,"del1":true,"del2":true,'
        '"dt2D":true,"dtr":true,"r1":"A1","r2":"B1","ref":"C1:D4"}'
    )


def test_real_openpyxl_array_formula_is_normalized_while_reading(
    tmp_path: Path,
) -> None:
    path = tmp_path / "array.xlsx"
    workbook = Workbook()
    worksheet = workbook.active
    worksheet["A1"] = array_formula()
    workbook.save(path)

    value = read_workbook(path)["Sheet"]["A1"]

    assert value == (
        'ARRAYFORMULA|{"ref":"A1:A2","text":"=SUM(B1:B2)"}'
    )
    assert "object at 0x" not in value


def test_persistence_serializer_never_uses_object_identity() -> None:
    first = AuditService._serialize(array_formula())
    second = AuditService._serialize(array_formula())

    assert first == second == (
        'ARRAYFORMULA|{"ref":"A1:A2","text":"=SUM(B1:B2)"}'
    )
    assert "object at 0x" not in first
