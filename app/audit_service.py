"""Orquestra a auditoria incremental independentemente da fonte de versões."""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime, timezone
import logging
import sqlite3
import threading
import time
from collections.abc import Callable
from uuid import uuid4

from app.database import Database
from app.excel.comparator import CellChange, compare_snapshots
from app.excel.formulas import normalize_formula_value
from app.excel.reader import CellValue, Snapshot, read_workbook
from app.integrity import sha256_file
from app.models import AuditExecutionStatus, ProcessedVersionStatus
from app.sources.base import SpreadsheetInfo, VersionInfo, VersionSource


logger = logging.getLogger("auditoria_excel.audit")


@dataclass(frozen=True, slots=True)
class AuditResult:
    execution_code: str
    status: AuditExecutionStatus
    processed_versions: int
    changes: int
    initial_checkpoint: str | None
    final_version: str | None


@dataclass(frozen=True, slots=True)
class VersionProgress:
    """Evento imutável de uma etapa real do processamento de uma versão."""

    version: str
    percent: int
    stage: str
    occurred_at: float = field(default_factory=time.monotonic)


class AuditService:
    """Executa comparações consecutivas e confirma cada uma atomicamente."""

    def __init__(
        self,
        database: Database,
        source: VersionSource,
        progress_callback: Callable[[int, int], None] | None = None,
        checkpoint_callback: Callable[[str, int, int], None] | None = None,
        version_progress_callback: Callable[[VersionProgress], None] | None = None,
        pause_event: threading.Event | None = None,
        stop_event: threading.Event | None = None,
        control_callback: Callable[[str, str | None], None] | None = None,
    ) -> None:
        self.database = database
        self.source = source
        self.progress_callback = progress_callback
        self.checkpoint_callback = checkpoint_callback
        self.version_progress_callback = version_progress_callback
        self.pause_event = pause_event or threading.Event()
        self.stop_event = stop_event or threading.Event()
        self.control_callback = control_callback

    def audit(
        self,
        spreadsheet: SpreadsheetInfo,
        versions: list[VersionInfo] | tuple[VersionInfo, ...] | None = None,
    ) -> AuditResult:
        connection = self.database.connection
        spreadsheet_id = self._upsert_spreadsheet(connection, spreadsheet)
        checkpoint = connection.execute(
            "SELECT versao_id, versao_numero FROM checkpoint WHERE planilha_id = ?",
            (spreadsheet_id,),
        ).fetchone()
        initial_checkpoint = checkpoint["versao_numero"] if checkpoint else None
        execution_code = f"AUD-{datetime.now(timezone.utc):%Y%m%d%H%M%S}-{uuid4().hex[:8]}"
        execution_id = connection.execute(
            """
            INSERT INTO execucao_auditoria
                (codigo_execucao, planilha_id, checkpoint_inicial, status)
            VALUES (?, ?, ?, ?)
            """,
            (
                execution_code,
                spreadsheet_id,
                initial_checkpoint,
                AuditExecutionStatus.RUNNING.value,
            ),
        ).lastrowid
        connection.commit()
        assert execution_id is not None
        logger.info(
            "Auditoria iniciada execucao=%s planilha=%s identidade=%s/%s/%s checkpoint=%s",
            execution_code,
            spreadsheet.name,
            spreadsheet.site_id,
            spreadsheet.drive_id,
            spreadsheet.drive_item_id,
            initial_checkpoint or "nenhum",
        )

        try:
            if versions is None:
                list_versions = getattr(self.source, "list_versions")
                try:
                    versions = list(
                        list_versions(
                            spreadsheet,
                            checkpoint_id=checkpoint["versao_id"] if checkpoint else None,
                            checkpoint_label=(
                                checkpoint["versao_numero"] if checkpoint else None
                            ),
                        )
                    )
                except TypeError:
                    # Fontes locais/alternativas conservam o contrato mínimo.
                    versions = list(list_versions(spreadsheet))
                logger.info(
                    "Lista de versões obtida na auditoria execucao=%s total=%d",
                    execution_code,
                    len(versions),
                )
            else:
                versions = list(versions)
                logger.info(
                    "Lista de versões reutilizada do cache da interface execucao=%s total=%d",
                    execution_code,
                    len(versions),
                )
            pairs = self._pending_pairs(versions, checkpoint["versao_id"] if checkpoint else None)
            logger.info(
                "Versões descobertas execucao=%s planilha=%s total=%d comparacoes_pendentes=%d",
                execution_code,
                spreadsheet.name,
                len(versions),
                len(pairs),
            )
            self._report_progress(0, len(pairs))
        except Exception as error:
            return self._record_failure(
                connection, execution_id, spreadsheet_id, execution_code,
                initial_checkpoint, 0, 0, None, None, error,
            )

        if not pairs:
            final = initial_checkpoint
            self._finish_execution(
                connection, execution_id, AuditExecutionStatus.COMPLETED_WITHOUT_UPDATES,
                final, 0, 0, None,
            )
            logger.info(
                "Auditoria sem novidades execucao=%s planilha=%s checkpoint=%s",
                execution_code,
                spreadsheet.name,
                final or "nenhum",
            )
            return AuditResult(
                execution_code, AuditExecutionStatus.COMPLETED_WITHOUT_UPDATES,
                0, 0, initial_checkpoint, final,
            )

        processed = 0
        total_changes = 0
        final = initial_checkpoint
        previous_snapshot = None
        audit_perf_started = time.perf_counter()
        if self._wait_at_safe_point(initial_checkpoint):
            return self._stop_execution(
                connection, execution_id, execution_code, initial_checkpoint,
                0, 0, initial_checkpoint,
            )
        for pair_index, (previous, current) in enumerate(pairs):
            pair_started = time.perf_counter()
            future_versions = tuple(pair[1] for pair in pairs[pair_index:])
            try:
                self._report_version_progress(
                    current.number, 0, f"Obtendo versão {current.number}..."
                )
                self._report_version_progress(current.number, 5, "Baixando dados...")
                logger.debug(
                    "Comparação iniciada execucao=%s planilha=%s versao_anterior=%s versao_atual=%s",
                    execution_code,
                    spreadsheet.name,
                    previous.number,
                    current.number,
                )
                previous_read_seconds = 0.0
                if previous_snapshot is None:
                    previous_started = time.perf_counter()
                    previous_snapshot, _ = self._read_temporary_version(
                        spreadsheet, previous,
                        prefetch_versions=self._planned_prefetch_versions(
                            previous, future_versions
                        ),
                    )
                    previous_read_seconds = time.perf_counter() - previous_started

                current_started = time.perf_counter()
                current_snapshot, current_hash = self._read_temporary_version(
                    spreadsheet, current,
                    prefetch_versions=self._planned_prefetch_versions(
                        current, future_versions[1:]
                    ),
                    report_stages=True,
                )
                current_read_seconds = time.perf_counter() - current_started

                compare_started = time.perf_counter()
                changes = compare_snapshots(previous_snapshot, current_snapshot)
                self._report_version_progress(current.number, 97, "Comparação concluída.")
                compare_seconds = time.perf_counter() - compare_started

                persist_started = time.perf_counter()
                self._report_version_progress(current.number, 97, "Salvando alterações...")
                self._persist_comparison(
                    connection, spreadsheet_id, execution_id, previous, current,
                    current_hash, changes,
                )
                persist_seconds = time.perf_counter() - persist_started
                pair_seconds = time.perf_counter() - pair_started

                logger.info(
                    "PERF comparacao planilha=%s anterior=%s atual=%s "
                    "leitura_anterior=%.3fs leitura_atual=%.3fs comparar=%.3fs "
                    "banco=%.3fs total=%.3fs alteracoes=%d",
                    spreadsheet.name,
                    previous.number,
                    current.number,
                    previous_read_seconds,
                    current_read_seconds,
                    compare_seconds,
                    persist_seconds,
                    pair_seconds,
                    len(changes),
                )
                logger.debug(
                    "Comparação concluída execucao=%s planilha=%s versao_atual=%s alteracoes=%d",
                    execution_code,
                    spreadsheet.name,
                    current.number,
                    len(changes),
                )
            except Exception as error:
                self._cancel_prefetch()
                connection.rollback()
                return self._record_failure(
                    connection, execution_id, spreadsheet_id, execution_code,
                    initial_checkpoint, processed, total_changes, previous, current, error,
                )
            self._report_version_progress(current.number, 100, "Checkpoint confirmado.")
            processed += 1
            self._report_progress(processed, len(pairs))
            self._report_checkpoint(current.number, processed, len(pairs) - processed)
            total_changes += len(changes)
            final = current.number
            previous_snapshot = current_snapshot

            # Este é o único ponto de controle dentro do loop: a persistência e
            # o checkpoint da versão atual já foram confirmados atomicamente.
            if self._wait_at_safe_point(final):
                return self._stop_execution(
                    connection, execution_id, execution_code, initial_checkpoint,
                    processed, total_changes, final,
                )

        self._cancel_prefetch()
        self._finish_execution(
            connection, execution_id, AuditExecutionStatus.COMPLETED,
            final, processed, total_changes, None,
        )
        audit_perf_seconds = time.perf_counter() - audit_perf_started
        logger.info(
            "PERF auditoria_resumo planilha=%s comparacoes=%d total=%.3fs media=%.3fs_por_comparacao",
            spreadsheet.name,
            processed,
            audit_perf_seconds,
            (audit_perf_seconds / processed) if processed else 0.0,
        )
        logger.info(
            "Auditoria concluída execucao=%s planilha=%s versoes_processadas=%d alteracoes=%d checkpoint=%s",
            execution_code,
            spreadsheet.name,
            processed,
            total_changes,
            final or "nenhum",
        )
        return AuditResult(
            execution_code, AuditExecutionStatus.COMPLETED, processed,
            total_changes, initial_checkpoint, final,
        )

    def _wait_at_safe_point(self, checkpoint: str | None) -> bool:
        """Obedece pausa/parada somente entre comparações confirmadas."""
        if self.stop_event.is_set():
            self._cancel_prefetch()
            return True
        if not self.pause_event.is_set():
            return False

        # Downloads especulativos não devem continuar ocupando o buffer durante
        # uma pausa. O WebDriver e a sessão autenticada não são encerrados.
        self._cancel_prefetch()
        self._report_control("paused", checkpoint)
        while self.pause_event.is_set():
            if self.stop_event.wait(0.1):
                return True
        self._report_control("resumed", checkpoint)
        return self.stop_event.is_set()

    def _report_control(self, state: str, checkpoint: str | None) -> None:
        if self.control_callback is not None:
            try:
                self.control_callback(state, checkpoint)
            except Exception:
                logger.warning("Falha ao publicar estado de controle", exc_info=True)

    def _stop_execution(
        self, connection: sqlite3.Connection, execution_id: int,
        execution_code: str, initial_checkpoint: str | None,
        processed: int, changes: int, final: str | None,
    ) -> AuditResult:
        self._cancel_prefetch()
        self._finish_execution(
            connection, execution_id, AuditExecutionStatus.STOPPED,
            final, processed, changes, "Interrompida pelo usuário.",
        )
        self._report_control("stopped", final)
        logger.info(
            "Auditoria interrompida pelo usuário execucao=%s checkpoint=%s",
            execution_code, final or "nenhum",
        )
        return AuditResult(
            execution_code, AuditExecutionStatus.STOPPED, processed, changes,
            initial_checkpoint, final,
        )

    def _report_version_progress(
        self, version: str, percent: int, stage: str
    ) -> None:
        if self.version_progress_callback is not None:
            try:
                self.version_progress_callback(VersionProgress(version, percent, stage))
            except Exception:
                logger.warning("Falha ao publicar progresso da versão", exc_info=True)

    def _report_progress(self, completed: int, total: int) -> None:
        if self.progress_callback is not None:
            try:
                self.progress_callback(completed, total)
            except Exception:
                logger.warning("Falha ao publicar progresso da auditoria", exc_info=True)

    def _report_checkpoint(
        self, version: str, processed: int, pending: int
    ) -> None:
        """Publica somente checkpoints cujo bloco transacional já foi confirmado."""
        if self.checkpoint_callback is not None:
            try:
                self.checkpoint_callback(version, processed, pending)
            except Exception:
                logger.warning(
                    "Falha ao publicar checkpoint confirmado da auditoria",
                    exc_info=True,
                )

    def _start_prefetch(
        self, spreadsheet: SpreadsheetInfo, version: VersionInfo
    ) -> None:
        prefetch = getattr(self.source, "prefetch_version", None)
        if not callable(prefetch):
            return
        try:
            prefetch(spreadsheet, version)
        except Exception as error:
            # Prefetch é apenas otimização. Uma falha aqui não invalida a
            # auditoria: get_version fará o download normal quando necessário.
            logger.warning(
                "Prefetch opcional falhou planilha=%s versao=%s erro=%s",
                spreadsheet.name,
                version.number,
                error,
            )

    def _planned_prefetch_versions(
        self,
        current: VersionInfo,
        candidates: tuple[VersionInfo, ...],
    ) -> tuple[VersionInfo, ...]:
        """Seleciona, em ordem, as próximas versões que cabem no buffer."""
        configured_size = getattr(self.source, "prefetch_buffer_size", 2)
        buffer_size = configured_size if isinstance(configured_size, int) else 2
        if buffer_size <= 0:
            return ()
        selected: list[VersionInfo] = []
        seen = {current.id}
        for candidate in candidates:
            if candidate.id in seen:
                continue
            seen.add(candidate.id)
            selected.append(candidate)
            if len(selected) >= buffer_size:
                break
        return tuple(selected)

    def _cancel_prefetch(self) -> None:
        cancel = getattr(self.source, "cancel_prefetch", None)
        if callable(cancel):
            try:
                cancel()
            except Exception:
                logger.debug("Falha ao cancelar prefetch opcional", exc_info=True)

    def _read_temporary_version(
        self,
        spreadsheet: SpreadsheetInfo,
        version: VersionInfo,
        *,
        prefetch_versions: tuple[VersionInfo, ...] = (),
        report_stages: bool = False,
    ) -> tuple[Snapshot, str]:
        total_started = time.perf_counter()
        download_started = time.perf_counter()
        path = self.source.get_version(spreadsheet, version)
        download_seconds = time.perf_counter() - download_started
        if report_stages:
            self._report_version_progress(version.number, 55, "Validando arquivo...")

        # Assim que o binário atual chegou ao disco, o Edge começa a buscar a
        # próxima versão em background. Enquanto isso, o Python calcula SHA,
        # lê o XLSX e compara a versão atual. Não há duas comparações paralelas
        # e o checkpoint continua sendo confirmado estritamente em ordem.
        logger.info(
            "PERF prefetch_planejado atual=%s futuras=[%s] quantidade=%d",
            version.number,
            ",".join(candidate.number for candidate in prefetch_versions),
            len(prefetch_versions),
        )
        for candidate in prefetch_versions:
            self._start_prefetch(spreadsheet, candidate)
        try:
            # O digest representa exatamente o binário adquirido nesta execução,
            # antes que o XLSX temporário seja descartado pela fonte.
            hash_started = time.perf_counter()
            digest = sha256_file(path)
            verify_digest = getattr(self.source, "verify_download_digest", None)
            if callable(verify_digest):
                verify_digest(path, digest)
            hash_seconds = time.perf_counter() - hash_started
            if report_stages:
                self._report_version_progress(version.number, 70, "Lendo XLSX...")

            read_started = time.perf_counter()
            snapshot = read_workbook(path)
            read_seconds = time.perf_counter() - read_started
            if report_stages:
                self._report_version_progress(version.number, 90, "Comparando...")
            total_seconds = time.perf_counter() - total_started
            try:
                file_size = path.stat().st_size
            except OSError:
                file_size = -1
            logger.info(
                "PERF versao planilha=%s versao=%s bytes=%d download=%.3fs sha256=%.3fs "
                "leitura_xlsx=%.3fs total=%.3fs",
                spreadsheet.name,
                version.number,
                file_size,
                download_seconds,
                hash_seconds,
                read_seconds,
                total_seconds,
            )
            return snapshot, digest
        finally:
            try:
                self.source.release_version(path)
            except Exception as error:
                logger.warning(
                    "Falha ao remover XLSX temporário planilha=%s versao=%s erro=%s",
                    spreadsheet.name,
                    version.number,
                    error,
                )

    @staticmethod
    def _pending_pairs(
        versions: list[VersionInfo], checkpoint_id: str | None
    ) -> list[tuple[VersionInfo, VersionInfo]]:
        ids = [version.id for version in versions]
        if len(ids) != len(set(ids)):
            raise ValueError("A fonte retornou identificadores de versão duplicados")
        if checkpoint_id is None:
            start = 0
        else:
            try:
                start = ids.index(checkpoint_id)
            except ValueError as error:
                raise ValueError("A versão do checkpoint não está disponível na fonte") from error
        return list(zip(versions[start:], versions[start + 1 :]))

    @staticmethod
    def _upsert_spreadsheet(
        connection: sqlite3.Connection, spreadsheet: SpreadsheetInfo
    ) -> int:
        connection.execute(
            """
            INSERT INTO planilha
                (site_id, drive_id, drive_item_id, nome_atual, caminho_sharepoint)
            VALUES (?, ?, ?, ?, ?)
            ON CONFLICT(site_id, drive_id, drive_item_id) DO UPDATE SET
                nome_atual = excluded.nome_atual,
                caminho_sharepoint = excluded.caminho_sharepoint,
                data_atualizacao = CURRENT_TIMESTAMP
            """,
            (spreadsheet.site_id, spreadsheet.drive_id, spreadsheet.drive_item_id,
             spreadsheet.name, spreadsheet.path),
        )
        row = connection.execute(
            """SELECT id FROM planilha
               WHERE site_id = ? AND drive_id = ? AND drive_item_id = ?""",
            (spreadsheet.site_id, spreadsheet.drive_id, spreadsheet.drive_item_id),
        ).fetchone()
        connection.commit()
        return int(row["id"])

    @staticmethod
    def _serialize(value: CellValue) -> str | None:
        normalized = normalize_formula_value(value)
        return None if normalized is None else str(normalized)

    def _persist_comparison(
        self,
        connection: sqlite3.Connection,
        spreadsheet_id: int,
        execution_id: int,
        previous: VersionInfo,
        current: VersionInfo,
        current_hash: str,
        changes: list[CellChange],
    ) -> None:
        status = (
            ProcessedVersionStatus.PROCESSED
            if changes
            else ProcessedVersionStatus.WITHOUT_CHANGES
        )
        with connection:
            cursor = connection.execute(
                """
                INSERT INTO versao_processada (
                    planilha_id, versao_anterior_id, versao_anterior_numero,
                    versao_atual_id, versao_atual_numero, data_hora_versao,
                    autor, autor_email, autor_login, comentario, tamanho,
                    url_origem, versao_atual, quantidade_alteracoes, status,
                    hash_origem, execucao_id
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (spreadsheet_id, previous.id, previous.number, current.id,
                 current.number, current.modified_at, current.author,
                 current.author_email, current.author_login, current.comment,
                 current.size, current.source_url, int(current.is_current),
                 len(changes), status.value, current_hash, execution_id),
            )
            processed_id = cursor.lastrowid
            assert processed_id is not None
            connection.executemany(
                """
                INSERT INTO alteracao (
                    versao_processada_id, planilha_id, tipo, aba, endereco,
                    valor_anterior, valor_novo
                ) VALUES (?, ?, ?, ?, ?, ?, ?)
                """,
                [
                    (processed_id, spreadsheet_id, change.change_type.value,
                     change.sheet, change.address,
                     self._serialize(change.previous_value),
                     self._serialize(change.new_value))
                    for change in changes
                ],
            )
            connection.execute(
                """
                INSERT INTO checkpoint
                    (planilha_id, versao_id, versao_numero, data_hora_versao)
                VALUES (?, ?, ?, ?)
                ON CONFLICT(planilha_id) DO UPDATE SET
                    versao_id = excluded.versao_id,
                    versao_numero = excluded.versao_numero,
                    data_hora_versao = excluded.data_hora_versao,
                    data_atualizacao = CURRENT_TIMESTAMP
                """,
                (spreadsheet_id, current.id, current.number, current.modified_at),
            )

    @staticmethod
    def _finish_execution(
        connection: sqlite3.Connection,
        execution_id: int,
        status: AuditExecutionStatus,
        final: str | None,
        processed: int,
        changes: int,
        message: str | None,
    ) -> None:
        # O schema histórico restringe os valores persistidos. Uma interrupção
        # graciosa é uma execução concluída (não uma falha); a mensagem preserva
        # a causa sem exigir migração destrutiva da tabela existente.
        persisted_status = (
            AuditExecutionStatus.COMPLETED
            if status is AuditExecutionStatus.STOPPED
            else status
        )
        connection.execute(
            """
            UPDATE execucao_auditoria SET
                fim = CURRENT_TIMESTAMP, versao_final = ?,
                versoes_processadas = ?, alteracoes_encontradas = ?,
                status = ?, mensagem = ?
            WHERE id = ?
            """,
            (final, processed, changes, persisted_status.value, message, execution_id),
        )
        connection.commit()

    def _record_failure(
        self, connection: sqlite3.Connection, execution_id: int,
        spreadsheet_id: int, execution_code: str,
        initial_checkpoint: str | None, processed: int, changes: int,
        previous: VersionInfo | None, current: VersionInfo | None,
        error: Exception,
    ) -> AuditResult:
        message = str(error)
        connection.execute(
            """
            INSERT INTO erro_processamento (
                execucao_id, planilha_id, versao_anterior, versao_atual,
                tipo_erro, mensagem
            ) VALUES (?, ?, ?, ?, ?, ?)
            """,
            (execution_id, spreadsheet_id,
             previous.number if previous else None,
             current.number if current else None,
             type(error).__name__, message),
        )
        checkpoint = connection.execute(
            "SELECT versao_numero FROM checkpoint WHERE planilha_id = ?",
            (spreadsheet_id,),
        ).fetchone()
        final = checkpoint["versao_numero"] if checkpoint else None
        self._finish_execution(
            connection, execution_id, AuditExecutionStatus.FAILED,
            final, processed, changes, message,
        )
        logger.error(
            "Auditoria falhou execucao=%s planilha_id=%d etapa=%s->%s tipo_erro=%s checkpoint_preservado=%s erro=%s",
            execution_code,
            spreadsheet_id,
            previous.number if previous else "listagem",
            current.number if current else "listagem",
            type(error).__name__,
            final or "nenhum",
            error,
        )
        return AuditResult(
            execution_code, AuditExecutionStatus.FAILED, processed, changes,
            initial_checkpoint, final,
        )
