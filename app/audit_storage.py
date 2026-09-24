"""Gerenciamento exclusivamente local das auditorias persistidas em SQLite."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, timezone
import hashlib
import hmac
import json
import os
from pathlib import Path
import re
import sqlite3
import threading
from typing import Any

from app.database import Database, SCHEMA, SCHEMA_VERSION


BACKUP_FORMAT_VERSION = 1
DEPENDENT_TABLES = (
    "alteracao",
    "versao_processada",
    "erro_processamento",
    "checkpoint",
    "execucao_auditoria",
)


class BackupError(RuntimeError):
    """Indica backup ausente, corrompido ou incompatível."""


class RestoreConflictError(BackupError):
    """Indica que uma restauração individual substituiria dados existentes."""


@dataclass(frozen=True)
class StoredAudit:
    id: int
    name: str
    path: str | None
    unique_id: str
    last_version: str | None
    processed_versions: int
    changes: int
    last_audit: str | None
    checkpoint: str | None


class AuditStorageManager:
    """Opera no banco local sem depender de qualquer fonte SharePoint."""

    def __init__(self, database: Database, backups_directory: str | Path) -> None:
        self.database = database
        self.backups_directory = Path(backups_directory)
        self._lock = threading.RLock()

    def list_audits(self) -> list[StoredAudit]:
        rows = self.database.connection.execute(
            """
            WITH versoes AS (
                SELECT planilha_id, COUNT(*) AS total
                  FROM versao_processada
                 GROUP BY planilha_id
            ), alteracoes AS (
                SELECT planilha_id, COUNT(*) AS total
                  FROM alteracao
                 GROUP BY planilha_id
            ), execucoes AS (
                SELECT planilha_id, MAX(fim) AS ultima
                  FROM execucao_auditoria
                 GROUP BY planilha_id
            )
            SELECT p.id, p.nome_atual, p.caminho_sharepoint, p.drive_item_id,
                   c.versao_numero,
                   COALESCE(v.total, 0) AS versoes,
                   COALESCE(a.total, 0) AS alteracoes,
                   e.ultima AS ultima_auditoria
              FROM planilha p
              LEFT JOIN checkpoint c ON c.planilha_id=p.id
              LEFT JOIN versoes v ON v.planilha_id=p.id
              LEFT JOIN alteracoes a ON a.planilha_id=p.id
              LEFT JOIN execucoes e ON e.planilha_id=p.id
             ORDER BY p.nome_atual COLLATE NOCASE, p.id
            """
        ).fetchall()
        return [
            StoredAudit(
                row["id"], row["nome_atual"], row["caminho_sharepoint"],
                row["drive_item_id"], row["versao_numero"], row["versoes"],
                row["alteracoes"], row["ultima_auditoria"], row["versao_numero"],
            )
            for row in rows
        ]

    def backup_individual(self, spreadsheet_id: int) -> Path:
        with self._lock:
            source = self.database.connection
            spreadsheet = source.execute(
                "SELECT * FROM planilha WHERE id=?", (spreadsheet_id,)
            ).fetchone()
            if spreadsheet is None:
                raise ValueError("Auditoria selecionada não existe.")
            path = self._new_path("individual", spreadsheet["drive_item_id"])
            target = sqlite3.connect(path)
            target.row_factory = sqlite3.Row
            try:
                target.execute("PRAGMA foreign_keys=ON")
                target.executescript(SCHEMA)
                target.execute(f"PRAGMA user_version={SCHEMA_VERSION}")
                target.execute(
                    """CREATE TABLE backup_metadata (
                    format_version INTEGER NOT NULL, backup_type TEXT NOT NULL,
                    created_utc TEXT NOT NULL, schema_version INTEGER NOT NULL,
                    spreadsheet_unique_id TEXT NOT NULL)"""
                )
                target.execute(
                    "INSERT INTO backup_metadata VALUES (?, 'individual', ?, ?, ?)",
                    (BACKUP_FORMAT_VERSION, self._utc_now(), SCHEMA_VERSION,
                     spreadsheet["drive_item_id"]),
                )
                self._copy_rows(source, target, "planilha", "id=?", (spreadsheet_id,))
                for table in ("checkpoint", "execucao_auditoria", "versao_processada",
                              "alteracao", "erro_processamento"):
                    self._copy_rows(source, target, table, "planilha_id=?", (spreadsheet_id,))
                identity = "|".join(
                    (spreadsheet["site_id"], spreadsheet["drive_id"], spreadsheet["drive_item_id"])
                )
                self._copy_rows(source, target, "version_catalog", "workbook_identity=?", (identity,))
                self._copy_rows(source, target, "version_catalog_state", "workbook_identity=?", (identity,))
                target.commit()
                self._check_integrity(target)
            except Exception:
                target.close()
                path.unlink(missing_ok=True)
                raise
            finally:
                if target:
                    target.close()
            self._write_checksum(path)
            return path

    def backup_full(self, *, prefix: str = "completo") -> Path:
        with self._lock:
            path = self._new_path(prefix)
            target = sqlite3.connect(path)
            try:
                self.database.connection.backup(target)
                self._check_integrity(target)
            except Exception:
                target.close()
                path.unlink(missing_ok=True)
                raise
            finally:
                target.close()
            self._write_metadata(path, "complete")
            self._write_checksum(path)
            return path

    def delete_individual(self, spreadsheet_id: int) -> None:
        with self._lock, self.database.connection:
            connection = self.database.connection
            spreadsheet = connection.execute(
                "SELECT site_id, drive_id, drive_item_id FROM planilha WHERE id=?",
                (spreadsheet_id,),
            ).fetchone()
            if spreadsheet is None:
                raise ValueError("Auditoria selecionada não existe.")
            identity = "|".join(tuple(spreadsheet))
            connection.execute("DELETE FROM version_catalog WHERE workbook_identity=?", (identity,))
            connection.execute("DELETE FROM version_catalog_state WHERE workbook_identity=?", (identity,))
            for table in DEPENDENT_TABLES:
                connection.execute(f"DELETE FROM {table} WHERE planilha_id=?", (spreadsheet_id,))
            deleted = connection.execute("DELETE FROM planilha WHERE id=?", (spreadsheet_id,))
            if deleted.rowcount != 1:
                raise ValueError("Auditoria selecionada não existe.")

    def delete_all(self) -> None:
        with self._lock, self.database.connection:
            connection = self.database.connection
            connection.execute("DELETE FROM version_catalog")
            connection.execute("DELETE FROM version_catalog_state")
            for table in DEPENDENT_TABLES:
                connection.execute(f"DELETE FROM {table}")
            connection.execute("DELETE FROM planilha")

    def restore_individual(self, path: str | Path, *, replace: bool = False) -> int:
        backup_path = Path(path)
        self._verify_checksum(backup_path)
        source = sqlite3.connect(f"file:{backup_path}?mode=ro", uri=True)
        source.row_factory = sqlite3.Row
        try:
            metadata = source.execute("SELECT * FROM backup_metadata").fetchone()
            if (metadata is None or metadata["format_version"] != BACKUP_FORMAT_VERSION
                    or metadata["backup_type"] != "individual"
                    or metadata["schema_version"] != SCHEMA_VERSION):
                raise BackupError("Backup individual incompatível.")
            self._check_integrity(source)
            spreadsheet = source.execute("SELECT * FROM planilha").fetchone()
            if spreadsheet is None or source.execute("SELECT COUNT(*) FROM planilha").fetchone()[0] != 1:
                raise BackupError("Backup individual inválido.")
            identity = (spreadsheet["site_id"], spreadsheet["drive_id"], spreadsheet["drive_item_id"])
            existing = self.database.connection.execute(
                "SELECT id FROM planilha WHERE site_id=? AND drive_id=? AND drive_item_id=?", identity
            ).fetchone()
            if existing and not replace:
                raise RestoreConflictError("Já existe auditoria local para esta planilha.")
            with self._lock, self.database.connection:
                if existing:
                    old = self.database.connection.execute(
                        "SELECT site_id, drive_id, drive_item_id FROM planilha WHERE id=?",
                        (existing["id"],),
                    ).fetchone()
                    old_identity = "|".join(tuple(old))
                    self.database.connection.execute(
                        "DELETE FROM version_catalog WHERE workbook_identity=?", (old_identity,)
                    )
                    self.database.connection.execute(
                        "DELETE FROM version_catalog_state WHERE workbook_identity=?", (old_identity,)
                    )
                    for table in DEPENDENT_TABLES:
                        self.database.connection.execute(
                            f"DELETE FROM {table} WHERE planilha_id=?", (existing["id"],)
                        )
                    self.database.connection.execute(
                        "DELETE FROM planilha WHERE id=?", (existing["id"],)
                    )
                destination = self.database.connection
                new_spreadsheet_id = self._insert_row(destination, "planilha", spreadsheet, {"id"})
                checkpoint = source.execute("SELECT * FROM checkpoint").fetchone()
                if checkpoint:
                    self._insert_row(destination, "checkpoint", checkpoint, {"id"}, {"planilha_id": new_spreadsheet_id})
                execution_ids: dict[int, int] = {}
                for row in source.execute("SELECT * FROM execucao_auditoria ORDER BY id"):
                    execution_ids[row["id"]] = self._insert_row(destination, "execucao_auditoria", row, {"id"}, {"planilha_id": new_spreadsheet_id})
                version_ids: dict[int, int] = {}
                for row in source.execute("SELECT * FROM versao_processada ORDER BY id"):
                    version_ids[row["id"]] = self._insert_row(destination, "versao_processada", row, {"id"}, {"planilha_id": new_spreadsheet_id, "execucao_id": execution_ids[row["execucao_id"]]})
                for row in source.execute("SELECT * FROM alteracao ORDER BY id"):
                    self._insert_row(destination, "alteracao", row, {"id"}, {"planilha_id": new_spreadsheet_id, "versao_processada_id": version_ids[row["versao_processada_id"]]})
                for row in source.execute("SELECT * FROM erro_processamento ORDER BY id"):
                    self._insert_row(destination, "erro_processamento", row, {"id"}, {"planilha_id": new_spreadsheet_id, "execucao_id": execution_ids[row["execucao_id"]]})
                for row in source.execute("SELECT * FROM version_catalog ORDER BY id"):
                    self._insert_row(destination, "version_catalog", row, {"id"})
                catalog_state = source.execute("SELECT * FROM version_catalog_state").fetchone()
                if catalog_state:
                    self._insert_row(destination, "version_catalog_state", catalog_state, set())
            return new_spreadsheet_id
        except sqlite3.DatabaseError as error:
            raise BackupError(f"Backup individual corrompido: {error}") from error
        finally:
            source.close()

    def restore_full(self, path: str | Path) -> Path:
        backup_path = Path(path)
        self._verify_checksum(backup_path)
        metadata = self._read_metadata(backup_path)
        if (metadata.get("format_version") != BACKUP_FORMAT_VERSION
                or metadata.get("backup_type") != "complete"
                or metadata.get("schema_version") != SCHEMA_VERSION):
            raise BackupError("Backup completo incompatível.")
        source: sqlite3.Connection | None = None
        try:
            source = sqlite3.connect(f"file:{backup_path}?mode=ro", uri=True)
            self._check_integrity(source)
            if source.execute("PRAGMA user_version").fetchone()[0] != SCHEMA_VERSION:
                raise BackupError("Versão de schema incompatível.")
            required = {"planilha", "checkpoint", "execucao_auditoria", "versao_processada", "alteracao", "erro_processamento", "version_catalog", "version_catalog_state"}
            present = {row[0] for row in source.execute("SELECT name FROM sqlite_master WHERE type='table'")}
            if not required <= present:
                raise BackupError("Backup completo não contém o schema canônico.")
        except sqlite3.DatabaseError as error:
            raise BackupError(f"Backup completo corrompido: {error}") from error
        finally:
            if source is not None:
                source.close()

        with self._lock:
            safety = self.backup_full(prefix="seguranca_pre_restauracao")
            staged = self.database.path.with_suffix(self.database.path.suffix + ".restore.tmp")
            staged.unlink(missing_ok=True)
            source = sqlite3.connect(f"file:{backup_path}?mode=ro", uri=True)
            target = sqlite3.connect(staged)
            try:
                source.backup(target)
                self._check_integrity(target)
            finally:
                source.close()
                target.close()
            self.database.close()
            try:
                os.replace(staged, self.database.path)
            finally:
                self.database.connect()
            return safety

    @staticmethod
    def _copy_rows(source: sqlite3.Connection, target: sqlite3.Connection,
                   table: str, where: str = "1=1", parameters: tuple[Any, ...] = ()) -> None:
        columns = [row[1] for row in source.execute(f"PRAGMA table_info({table})")]
        quoted = ", ".join(f'"{column}"' for column in columns)
        placeholders = ", ".join("?" for _ in columns)
        for row in source.execute(f"SELECT {quoted} FROM {table} WHERE {where}", parameters):
            target.execute(f"INSERT INTO {table} ({quoted}) VALUES ({placeholders})", tuple(row))

    @staticmethod
    def _insert_row(connection: sqlite3.Connection, table: str, row: sqlite3.Row,
                    excluded: set[str], replacements: dict[str, Any] | None = None) -> int:
        replacements = replacements or {}
        columns = [name for name in row.keys() if name not in excluded]
        quoted = ", ".join(f'"{name}"' for name in columns)
        values = [replacements.get(name, row[name]) for name in columns]
        cursor = connection.execute(
            f"INSERT INTO {table} ({quoted}) VALUES ({', '.join('?' for _ in columns)})",
            values,
        )
        assert cursor.lastrowid is not None
        return cursor.lastrowid

    def _new_path(self, kind: str, identifier: str = "") -> Path:
        self.backups_directory.mkdir(parents=True, exist_ok=True)
        safe = re.sub(r"[^A-Za-z0-9._-]+", "_", identifier).strip("._-")[:60]
        stamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%S%fZ")
        suffix = f"_{safe}" if safe else ""
        return self.backups_directory / f"auditoria_{kind}_{stamp}{suffix}.sqlite3"

    @staticmethod
    def _utc_now() -> str:
        return datetime.now(timezone.utc).isoformat().replace("+00:00", "Z")

    @staticmethod
    def _digest(path: Path) -> str:
        digest = hashlib.sha256()
        with path.open("rb") as stream:
            for chunk in iter(lambda: stream.read(1024 * 1024), b""):
                digest.update(chunk)
        return digest.hexdigest()

    def _write_checksum(self, path: Path) -> None:
        path.with_suffix(path.suffix + ".sha256").write_text(
            f"{self._digest(path)}  {path.name}\n", encoding="ascii"
        )

    def _verify_checksum(self, path: Path) -> None:
        if not path.is_file():
            raise BackupError("Arquivo de backup não encontrado.")
        checksum = path.with_suffix(path.suffix + ".sha256")
        try:
            expected = checksum.read_text(encoding="ascii").split()[0]
        except (OSError, IndexError, UnicodeError) as error:
            raise BackupError("Arquivo de integridade SHA-256 ausente ou inválido.") from error
        if not hmac.compare_digest(expected, self._digest(path)):
            raise BackupError("Backup corrompido ou alterado (SHA-256 divergente).")

    def _write_metadata(self, path: Path, backup_type: str) -> None:
        path.with_suffix(path.suffix + ".meta.json").write_text(json.dumps({
            "format_version": BACKUP_FORMAT_VERSION,
            "backup_type": backup_type,
            "created_utc": self._utc_now(),
            "schema_version": SCHEMA_VERSION,
        }, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")

    @staticmethod
    def _read_metadata(path: Path) -> dict[str, Any]:
        try:
            value = json.loads(path.with_suffix(path.suffix + ".meta.json").read_text(encoding="utf-8"))
        except (OSError, UnicodeError, json.JSONDecodeError) as error:
            raise BackupError("Metadados do backup ausentes ou inválidos.") from error
        if not isinstance(value, dict):
            raise BackupError("Metadados do backup incompatíveis.")
        return value

    @staticmethod
    def _check_integrity(connection: sqlite3.Connection) -> None:
        if connection.execute("PRAGMA integrity_check").fetchone()[0] != "ok":
            raise BackupError("Falha na verificação de integridade SQLite.")
        violations = connection.execute("PRAGMA foreign_key_check").fetchall()
        if violations:
            raise BackupError("Backup contém violações de chave estrangeira.")
