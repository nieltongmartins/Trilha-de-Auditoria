"""Catálogo SQLite de metadados de versões e sincronização por watermark."""

from __future__ import annotations

from dataclasses import dataclass, replace
from datetime import datetime, timezone
import logging
import time
from collections.abc import Callable, Sequence

from app.database import Database
from app.sources.base import SpreadsheetInfo, VersionInfo, VersionSource


logger = logging.getLogger("auditoria_excel.version_catalog")


class CatalogInconsistencyError(RuntimeError):
    """Indica que a fronteira incremental não pôde ser comprovada."""


@dataclass(frozen=True, slots=True)
class CatalogSyncResult:
    versions: tuple[VersionInfo, ...]
    source: str
    new_versions: int


class VersionCatalog:
    """Mantém somente metadata; arquivos XLSX nunca são persistidos aqui."""

    def __init__(self, database: Database) -> None:
        self.database = database

    @staticmethod
    def identity(spreadsheet: SpreadsheetInfo) -> str:
        return "|".join(
            (spreadsheet.site_id, spreadsheet.drive_id, spreadsheet.drive_item_id)
        )

    def load(self, spreadsheet: SpreadsheetInfo) -> tuple[VersionInfo, ...] | None:
        started = time.perf_counter()
        identity = self.identity(spreadsheet)
        state = self.database.connection.execute(
            "SELECT * FROM version_catalog_state WHERE workbook_identity=?",
            (identity,),
        ).fetchone()
        if state is None or state["catalog_status"] != "VALID":
            return None
        rows = self.database.connection.execute(
            """SELECT * FROM version_catalog WHERE workbook_identity=?
               ORDER BY CAST(technical_version_id AS INTEGER), id""",
            (identity,),
        ).fetchall()
        try:
            versions = tuple(self._from_row(row) for row in rows)
            self._validate_complete(versions, state)
        except (ValueError, CatalogInconsistencyError):
            with self.database.connection:
                self.database.connection.execute(
                    "UPDATE version_catalog_state SET catalog_status='INVALID' "
                    "WHERE workbook_identity=?", (identity,)
                )
            logger.exception("Catálogo local inválido workbook=%s", identity)
            return None
        logger.info(
            "VERSION_CATALOG_LOAD workbook=%s count=%d duration=%.3f "
            "max_known_id=%s current_id=%s",
            identity, len(versions), time.perf_counter() - started,
            versions[-1].id if versions else None, state["last_known_current_id"],
        )
        return versions

    def sync(
        self,
        spreadsheet: SpreadsheetInfo,
        source: VersionSource,
        progress_callback: Callable[[int], None] | None = None,
    ) -> CatalogSyncResult:
        started = time.perf_counter()
        cached = self.load(spreadsheet)
        state = self._state(spreadsheet) if cached is not None else None
        if cached is None or state is None:
            versions = self._list_full(source, spreadsheet, progress_callback)
            self._replace(spreadsheet, versions)
            self._log_complete(spreadsheet, "full_rebuild", versions, started)
            return CatalogSyncResult(versions, "full_rebuild", len(versions))

        current_started = time.perf_counter()
        get_current = getattr(source, "get_current_version", None)
        if not callable(get_current):
            # A sincronização leve é uma capacidade explícita. Fontes locais de
            # teste/legadas continuam funcionais, sem fingir que houve watermark.
            versions = self._list_full(source, spreadsheet, progress_callback)
            self._replace(spreadsheet, versions)
            self._log_complete(spreadsheet, "full_rebuild", versions, started)
            return CatalogSyncResult(versions, "full_rebuild", len(versions))
        remote = get_current(spreadsheet)
        cached_id = state["last_known_current_id"]
        changed = remote.id != cached_id or remote.number != state["last_known_current_label"]
        logger.info(
            "VERSION_CATALOG_REMOTE_CHECK cached_current_id=%s remote_current_id=%s "
            "changed=%s duration=%.3f",
            cached_id, remote.id, str(changed).lower(),
            time.perf_counter() - current_started,
        )
        if not changed:
            self._touch(spreadsheet)
            self._log_complete(spreadsheet, "local_only", cached, started)
            return CatalogSyncResult(cached, "local_only", 0)

        try:
            if int(remote.id) <= int(cached_id):
                raise CatalogInconsistencyError("watermark remoto não avançou")
            delta_started = time.perf_counter()
            list_delta = getattr(source, "list_versions_delta", None)
            if callable(list_delta):
                tail = tuple(list_delta(spreadsheet, cached_id, state["last_known_current_label"], progress_callback))
            else:
                raise CatalogInconsistencyError(
                    "fonte não oferece consulta incremental comprovável"
                )
            merged, added = self._merge(cached, tail, remote, cached_id)
            self._replace(spreadsheet, merged, rebuild=False)
            logger.info(
                "VERSION_CATALOG_DELTA anchor_id=%s new_historical_count=%d "
                "new_current_id=%s duration=%.3f",
                cached_id, max(added - 1, 0), remote.id,
                time.perf_counter() - delta_started,
            )
            self._log_complete(spreadsheet, "delta", merged, started)
            return CatalogSyncResult(merged, "delta", added)
        except (CatalogInconsistencyError, ValueError) as error:
            logger.warning("Catálogo requer reconstrução workbook=%s motivo=%s", self.identity(spreadsheet), error)
            self._mark(spreadsheet, "NEEDS_RECONCILIATION")
            versions = self._list_full(source, spreadsheet, progress_callback)
            self._replace(spreadsheet, versions)
            self._log_complete(spreadsheet, "full_rebuild", versions, started)
            return CatalogSyncResult(versions, "full_rebuild", len(versions))

    @staticmethod
    def _list_full(
        source: VersionSource,
        spreadsheet: SpreadsheetInfo,
        progress_callback: Callable[[int], None] | None,
    ) -> tuple[VersionInfo, ...]:
        try:
            return tuple(source.list_versions(  # type: ignore[call-arg]
                spreadsheet, progress_callback=progress_callback
            ))
        except TypeError:
            return tuple(source.list_versions(spreadsheet))

    def _state(self, spreadsheet: SpreadsheetInfo):
        return self.database.connection.execute(
            "SELECT * FROM version_catalog_state WHERE workbook_identity=?",
            (self.identity(spreadsheet),),
        ).fetchone()

    @staticmethod
    def _from_row(row) -> VersionInfo:
        return VersionInfo(
            id=row["technical_version_id"], number=row["version_label"],
            modified_at=row["created_at_sharepoint"], author=row["author"],
            author_email=row["author_email"], author_login=row["author_login"],
            comment=row["comment"], size=row["size"], source_url=row["source_url"],
            is_current=bool(row["is_current_snapshot"]),
        )

    @staticmethod
    def _validate_versions(versions: Sequence[VersionInfo]) -> None:
        if not versions:
            raise CatalogInconsistencyError("catálogo vazio")
        ids = [int(version.id) for version in versions]
        if ids != sorted(ids) or len(ids) != len(set(ids)):
            raise CatalogInconsistencyError("IDs não são únicos e monotônicos")
        labels = [version.number for version in versions]
        if len(labels) != len(set(labels)):
            raise CatalogInconsistencyError("VersionLabel duplicado")
        if sum(version.is_current for version in versions) != 1 or not versions[-1].is_current:
            raise CatalogInconsistencyError("current snapshot inválido")

    def _validate_complete(self, versions: Sequence[VersionInfo], state) -> None:
        self._validate_versions(versions)
        if len(versions) != state["catalog_count"]:
            raise CatalogInconsistencyError("contagem divergente")
        current = versions[-1]
        if (current.id, current.number) != (
            state["last_known_current_id"], state["last_known_current_label"]
        ):
            raise CatalogInconsistencyError("watermark local divergente")

    def _merge(self, cached, tail, remote, anchor_id):
        self._validate_versions(tail)
        anchor = next((item for item in tail if item.id == anchor_id), None)
        known = {item.id: item for item in cached}
        if anchor is None or anchor.number != known[anchor_id].number:
            raise CatalogInconsistencyError("anchor técnico ausente ou conflitante")
        for item in tail:
            previous = known.get(item.id)
            if previous is not None and previous.number != item.number:
                raise CatalogInconsistencyError("label conflitante para technical ID")
            known[item.id] = item
        if remote.id not in known or known[remote.id].number != remote.number:
            raise CatalogInconsistencyError("current metadata incoerente com a cauda")
        merged = tuple(
            replace(item, is_current=item.id == remote.id)
            for item in sorted(known.values(), key=lambda version: int(version.id))
        )
        self._validate_versions(merged)
        return merged, len([item for item in merged if int(item.id) > int(anchor_id)])

    def _replace(
        self, spreadsheet: SpreadsheetInfo, versions: Sequence[VersionInfo], *,
        rebuild: bool = True,
    ) -> None:
        self._validate_versions(versions)
        identity = self.identity(spreadsheet)
        now = datetime.now(timezone.utc).isoformat()
        historical = versions[-2] if len(versions) > 1 else None
        current = versions[-1]
        connection = self.database.connection
        with connection:
            if rebuild:
                connection.execute("DELETE FROM version_catalog WHERE workbook_identity=?", (identity,))
            else:
                connection.execute(
                    "UPDATE version_catalog SET is_current_snapshot=0 "
                    "WHERE workbook_identity=?", (identity,)
                )
            connection.executemany(
                """INSERT INTO version_catalog
                   (workbook_identity, technical_version_id, version_label,
                    created_at_sharepoint, is_current_snapshot, discovered_at,
                    last_verified_at, author, author_email, author_login, comment,
                    size, source_url)
                   VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                   ON CONFLICT(workbook_identity, technical_version_id) DO UPDATE SET
                   version_label=excluded.version_label,
                   created_at_sharepoint=COALESCE(excluded.created_at_sharepoint, created_at_sharepoint),
                   is_current_snapshot=excluded.is_current_snapshot,
                   last_verified_at=excluded.last_verified_at,
                   author=COALESCE(excluded.author, author),
                   author_email=COALESCE(excluded.author_email, author_email),
                   author_login=COALESCE(excluded.author_login, author_login),
                   comment=COALESCE(excluded.comment, comment),
                   size=COALESCE(excluded.size, size),
                   source_url=COALESCE(excluded.source_url, source_url)""",
                [(identity, v.id, v.number, v.modified_at, int(v.is_current), now,
                  now, v.author, v.author_email, v.author_login, v.comment, v.size,
                  v.source_url) for v in versions],
            )
            connection.execute(
                """INSERT INTO version_catalog_state VALUES (?, ?, ?, ?, ?, ?, ?, 'VALID')
                   ON CONFLICT(workbook_identity) DO UPDATE SET
                   last_historical_id=excluded.last_historical_id,
                   last_historical_label=excluded.last_historical_label,
                   last_known_current_id=excluded.last_known_current_id,
                   last_known_current_label=excluded.last_known_current_label,
                   catalog_count=excluded.catalog_count,
                   last_sync_at=excluded.last_sync_at, catalog_status='VALID'""",
                (identity, historical.id if historical else None,
                 historical.number if historical else None, current.id,
                 current.number, len(versions), now),
            )

    def _touch(self, spreadsheet: SpreadsheetInfo) -> None:
        identity = self.identity(spreadsheet)
        now = datetime.now(timezone.utc).isoformat()
        with self.database.connection:
            self.database.connection.execute(
                "UPDATE version_catalog SET last_verified_at=? WHERE workbook_identity=?",
                (now, identity),
            )
            self.database.connection.execute(
                "UPDATE version_catalog_state SET last_sync_at=? WHERE workbook_identity=?",
                (now, identity),
            )

    def _mark(self, spreadsheet: SpreadsheetInfo, status: str) -> None:
        with self.database.connection:
            self.database.connection.execute(
                "UPDATE version_catalog_state SET catalog_status=? WHERE workbook_identity=?",
                (status, self.identity(spreadsheet)),
            )

    @staticmethod
    def _log_complete(spreadsheet, source, versions, started) -> None:
        logger.info(
            "VERSION_CATALOG_SYNC_COMPLETE workbook=%s source=%s count=%d duration=%.3f",
            spreadsheet.drive_item_id, source, len(versions), time.perf_counter() - started,
        )
