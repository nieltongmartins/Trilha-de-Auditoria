from dataclasses import replace
from pathlib import Path

from app.database import Database
from app.sources.base import SpreadsheetInfo, VersionInfo
from app.version_catalog import VersionCatalog


BOOK = SpreadsheetInfo("site", "drive", "book-id", "CQLPA123.xlsx", "/CQLPA123.xlsx")


def version(identifier: int, label: str, *, current: bool = False) -> VersionInfo:
    return VersionInfo(str(identifier), label, f"2026-09-24T00:{identifier:02d}:00Z", is_current=current)


class CatalogSource:
    def __init__(self, complete, current=None, delta=None):
        self.complete = tuple(complete)
        self.current = current or self.complete[-1]
        self.delta = tuple(delta or ())
        self.full_calls = 0
        self.current_calls = 0
        self.delta_calls = 0

    def list_versions(self, _spreadsheet):
        self.full_calls += 1
        return self.complete

    def get_current_version(self, _spreadsheet):
        self.current_calls += 1
        return self.current

    def list_versions_delta(self, _spreadsheet, anchor_id, anchor_label, _progress=None):
        self.delta_calls += 1
        assert (anchor_id, anchor_label) == ("3", "1.2")
        return self.delta


def database(tmp_path: Path) -> Database:
    result = Database(tmp_path / "catalog.sqlite3")
    result.initialize()
    return result


def test_second_open_uses_one_current_query_and_zero_historical_pages(tmp_path: Path) -> None:
    db = database(tmp_path)
    initial = (version(1, "1.0"), version(2, "1.1"), version(3, "1.2", current=True))
    first_source = CatalogSource(initial)
    catalog = VersionCatalog(db)
    first = catalog.sync(BOOK, first_source)
    assert first.source == "full_rebuild"
    assert first_source.full_calls == 1

    second_source = CatalogSource(initial)
    second = catalog.sync(BOOK, second_source)
    assert second.source == "local_only"
    assert [item.id for item in second.versions] == ["1", "2", "3"]
    assert second_source.current_calls == 1
    assert second_source.delta_calls == 0
    assert second_source.full_calls == 0


def test_new_current_fetches_only_tail_and_promotes_cached_current_to_history(tmp_path: Path) -> None:
    db = database(tmp_path)
    old = (version(1, "1.0"), version(2, "1.1"), version(3, "1.2", current=True))
    catalog = VersionCatalog(db)
    catalog.sync(BOOK, CatalogSource(old))
    new_current = version(4, "1.3", current=True)
    tail = (replace(old[-1], is_current=False), new_current)
    source = CatalogSource(old, current=new_current, delta=tail)

    result = catalog.sync(BOOK, source)

    assert result.source == "delta"
    assert result.new_versions == 1
    assert [(item.id, item.is_current) for item in result.versions] == [
        ("1", False), ("2", False), ("3", False), ("4", True)
    ]
    assert source.delta_calls == 1
    assert source.full_calls == 0


def test_unprovable_anchor_is_the_only_path_to_full_rebuild(tmp_path: Path) -> None:
    db = database(tmp_path)
    old = (version(1, "1.0"), version(2, "1.1"), version(3, "1.2", current=True))
    catalog = VersionCatalog(db)
    catalog.sync(BOOK, CatalogSource(old))
    rebuilt = old[:-1] + (replace(old[-1], is_current=False), version(4, "1.3", current=True))
    source = CatalogSource(rebuilt, current=rebuilt[-1], delta=(rebuilt[-1],))

    result = catalog.sync(BOOK, source)

    assert result.source == "full_rebuild"
    assert source.delta_calls == 1
    assert source.full_calls == 1


def test_checkpoint_is_not_used_as_catalog_frontier(tmp_path: Path) -> None:
    db = database(tmp_path)
    versions = (version(1, "6.7"), version(2, "40.510"), version(3, "40.511", current=True))
    catalog = VersionCatalog(db)
    catalog.sync(BOOK, CatalogSource(versions))
    db.connection.execute(
        "INSERT INTO planilha (drive_item_id,nome_atual,site_id,drive_id) VALUES (?,?,?,?)",
        (BOOK.drive_item_id, BOOK.name, BOOK.site_id, BOOK.drive_id),
    )
    planilha_id = db.connection.execute("SELECT id FROM planilha").fetchone()[0]
    db.connection.execute(
        "INSERT INTO checkpoint (planilha_id,versao_id,versao_numero) VALUES (?,?,?)",
        (planilha_id, "1", "6.7"),
    )
    db.connection.commit()
    source = CatalogSource(versions)

    result = catalog.sync(BOOK, source)

    assert result.source == "local_only"
    assert source.full_calls == source.delta_calls == 0
    assert [item.number for item in result.versions[1:]] == ["40.510", "40.511"]
