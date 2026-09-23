"""Representa fórmulas especiais do openpyxl de modo estável e legível."""

from __future__ import annotations

import json
from typing import Any

from openpyxl.worksheet.formula import ArrayFormula, DataTableFormula


def _canonical(kind: str, attributes: dict[str, object]) -> str:
    payload = json.dumps(
        attributes,
        ensure_ascii=False,
        sort_keys=True,
        separators=(",", ":"),
    )
    return f"{kind}|{payload}"


def normalize_formula_value(value: Any) -> Any:
    """Converte somente objetos de fórmula especiais em texto canônico.

    Fórmulas comuns já são strings iniciadas por ``=`` no openpyxl e todos os
    demais valores devem manter tipo e valor.  A fórmula matricial inclui seu
    texto e sua referência, ambos semanticamente relevantes.  Fórmulas de
    tabela de dados não possuem texto no openpyxl 3.1; todos os metadados
    expostos pela classe são preservados para evitar equivalências indevidas.
    """

    if isinstance(value, ArrayFormula):
        return _canonical(
            "ARRAYFORMULA",
            {"ref": value.ref, "text": value.text},
        )
    if isinstance(value, DataTableFormula):
        return _canonical(
            "DATATABLEFORMULA",
            {
                "ca": value.ca,
                "del1": value.del1,
                "del2": value.del2,
                "dt2D": value.dt2D,
                "dtr": value.dtr,
                "r1": value.r1,
                "r2": value.r2,
                "ref": value.ref,
            },
        )
    return value
