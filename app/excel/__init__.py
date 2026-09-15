"""Leitura e comparação determinística de arquivos Excel."""

from app.excel.comparator import CellChange, compare_snapshots
from app.excel.reader import Snapshot, read_snapshot

__all__ = ["CellChange", "Snapshot", "compare_snapshots", "read_snapshot"]
