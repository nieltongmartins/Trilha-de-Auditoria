"""Interface desktop responsiva para configurar, conectar e auditar."""

from __future__ import annotations

from collections.abc import Callable, Sequence
from pathlib import Path
import logging
import queue
import subprocess
import sys
import threading
import time
from typing import TYPE_CHECKING
from tkinter import filedialog, messagebox, ttk
import tkinter as tk

from app.audit_storage import AuditStorageManager, RestoreConflictError
from app.database import Database
from app.models import AuditExecutionStatus
from app.progress import SmoothVersionProgress
from app.report_artifacts import ReportArtifactManager
from app.sources.base import SpreadsheetInfo, VersionInfo, VersionSource
from app.version_catalog import VersionCatalog

if TYPE_CHECKING:
    from app.audit_service import AuditResult


logger = logging.getLogger("auditoria_excel.interface")
SourceFactory = Callable[[str, tuple[str, ...]], VersionSource]
ConfigurationSaver = Callable[[str, tuple[str, ...]], None]


class AuditApplication(ttk.Frame):
    """Tela operacional; todo trabalho demorado ocorre fora da thread Tk."""

    def __init__(
        self,
        master: tk.Misc,
        database: Database,
        source: VersionSource | None,
        reports_directory: str | Path,
        backups_directory: str | Path = "data/backups",
        *,
        site_url: str = "",
        scope_paths: Sequence[str] = (),
        connect_source: SourceFactory | None = None,
        save_configuration: ConfigurationSaver | None = None,
    ) -> None:
        self._startup_started = time.perf_counter()
        self._startup_last = self._startup_started
        self._startup_complete = False
        self._startup_log("entrada_init")
        super().__init__(master, padding=12)
        self._startup_log("frame_tk")
        self.database = database
        self.source = source
        self.connect_source = connect_source
        self.save_configuration = save_configuration
        self.reports_directory = Path(reports_directory)
        self.storage = AuditStorageManager(database, backups_directory)
        self.version_catalog = VersionCatalog(database)
        self._startup_log("criar_audit_storage")
        self.report_artifacts = ReportArtifactManager(
            getattr(database, "connection", database), self.reports_directory
        )
        self._startup_log("criar_report_artifacts")
        self.spreadsheets: list[SpreadsheetInfo] = []
        # Cache curto da enumeração de versões. A listagem pode levar minutos em
        # arquivos com dezenas de milhares de versões; reutilizá-la evita repetir
        # a mesma consulta ao clicar em "Continuar auditoria".
        self._version_cache: dict[str, tuple[float, tuple[VersionInfo, ...]]] = {}
        self._version_cache_ttl = 300.0
        self.last_report: Path | None = None
        self._busy = False
        self._audit_active = False
        self._audit_paused = False
        self._pause_event = threading.Event()
        self._stop_event = threading.Event()
        self._work_results: queue.SimpleQueue[tuple[bool, object]] = (
            queue.SimpleQueue()
        )
        self._progress_updates: queue.SimpleQueue[tuple[int, int]] = queue.SimpleQueue()
        self._checkpoint_updates: queue.SimpleQueue[tuple[str, int, int]] = (
            queue.SimpleQueue()
        )
        self._version_progress_updates: queue.SimpleQueue[object] = queue.SimpleQueue()
        self._report_updates: queue.SimpleQueue[str] = queue.SimpleQueue()
        self._control_updates: queue.SimpleQueue[tuple[str, str | None]] = queue.SimpleQueue()
        self._version_scan_updates: queue.SimpleQueue[int] = queue.SimpleQueue()
        self._version_scan_started_at: float | None = None
        self._version_scan_active = False
        self._version_scan_checkpoint_label: str | None = None
        self._audit_started_at: float | None = None
        self._progress_completed = 0
        self._progress_total = 0
        self._latest_available = "—"
        self._current_version_started_at: float | None = None
        self._current_version = "—"
        self._smooth_version_progress = SmoothVersionProgress()
        self._version_animation_interval = 0.15
        self._version_last_animation = 0.0
        self._version_last_clock_update = 0.0
        self._hidden_clicks: list[float] = []
        # AuditService, ReportService, leitor XLSX/openpyxl, fontes SharePoint e
        # Graph permanecem fora deste caminho e são importados somente nas ações.
        self._startup_log("imports_lazy_nao_disparados")
        self.site_url = tk.StringVar(value=site_url)
        self.scope_paths = tk.StringVar(value=";".join(scope_paths))
        # Preserve the public attribute used by installations upgraded from the
        # first folder-field implementation. It now backs the read-only selector.
        self.folder_names = tk.StringVar(value="")
        self.folder_paths: list[str] = []
        self.status = tk.StringVar(
            value=(
                "Conectado. Atualize a lista."
                if source is not None
                else "Desconectado. Confira a configuração e clique em Conectar."
            )
        )
        self.details = tk.StringVar(value="Selecione uma planilha.")
        self._startup_log("configurar_variaveis_tk")
        self._build()
        self._startup_log("finalizacao")
        self._startup_complete = True

    def _startup_log(self, stage: str) -> None:
        """Registra cada trecho síncrono percorrido antes da primeira pintura."""
        now = time.perf_counter()
        logger.info(
            "STARTUP_INTERFACE etapa=%s duracao=%.3fs acumulado=%.3fs",
            stage,
            now - self._startup_last,
            now - self._startup_started,
        )
        self._startup_last = now

    def _build(self) -> None:
        self.master.columnconfigure(0, weight=1)
        self.master.rowconfigure(0, weight=1)
        self.grid(sticky="nsew")
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)
        self.notebook = ttk.Notebook(self)
        self.notebook.grid(row=0, column=0, sticky="nsew")
        audit_tab = ttk.Frame(self.notebook, padding=4)
        stored_tab = ttk.Frame(self.notebook, padding=8)
        self.notebook.add(audit_tab, text="Auditoria")
        self.notebook.add(stored_tab, text="Auditorias armazenadas")
        self._startup_log("criar_abas")
        self._audit_tab = audit_tab
        self._comparator_tab = None
        audit_tab.columnconfigure(0, weight=1)
        title_label = ttk.Label(
            audit_tab, text="Auditor de Planilhas", font=("TkDefaultFont", 14, "bold")
        )
        title_label.grid(sticky="w")
        # Gesto deliberado e invisível: cinco cliques no título existente.
        title_label.bind("<Button-1>", self._hidden_comparator_gesture)

        config = ttk.LabelFrame(audit_tab, text="Conexão SharePoint", padding=8)
        config.grid(row=1, column=0, sticky="ew", pady=(8, 4))
        config.columnconfigure(1, weight=1)
        ttk.Label(config, text="URL do site:").grid(row=0, column=0, sticky="w")
        ttk.Entry(config, textvariable=self.site_url).grid(
            row=0, column=1, sticky="ew", padx=(8, 0)
        )
        ttk.Label(config, text="Escopo(s):").grid(
            row=1, column=0, sticky="w", pady=(6, 0)
        )
        ttk.Entry(config, textvariable=self.scope_paths).grid(
            row=1, column=1, sticky="ew", padx=(8, 0), pady=(6, 0)
        )
        ttk.Label(
            config,
            text="Use o caminho completo da biblioteca/pasta; separe vários por ';'.",
        ).grid(
            row=2, column=1, sticky="w"
        )
        self.connect_button = ttk.Button(config, text="Conectar", command=self.connect)
        self.connect_button.grid(row=4, column=1, sticky="e", pady=(8, 0))

        self.selector = ttk.Combobox(audit_tab, state="readonly", width=70)
        self.selector.grid(row=2, column=0, sticky="ew", pady=8)
        self.selector.bind("<<ComboboxSelected>>", lambda _event: self.show_status())
        ttk.Label(audit_tab, textvariable=self.details).grid(row=3, column=0, sticky="w")
        buttons = ttk.Frame(audit_tab)
        buttons.grid(row=4, column=0, sticky="w", pady=10)
        self.refresh_button = ttk.Button(
            buttons, text="Atualizar lista", command=self.refresh
        )
        self.refresh_button.pack(side="left", padx=(0, 6))
        self.audit_button = ttk.Button(
            buttons, text="Auditar histórico", command=self.audit
        )
        self.audit_button.pack(side="left", padx=(0, 6))
        self.report_button = ttk.Button(
            buttons, text="Gerar relatório", command=self.generate_report
        )
        self.report_button.pack(side="left", padx=6)
        ttk.Button(buttons, text="Abrir relatório", command=self.open_report).pack(
            side="left", padx=6
        )
        self.pause_button = ttk.Button(buttons, text="Pausar", command=self.toggle_pause)
        self.pause_button.pack(side="left", padx=6)
        self.stop_button = ttk.Button(buttons, text="Parar", command=self.stop_audit)
        self.stop_button.pack(side="left", padx=6)
        ttk.Label(audit_tab, textvariable=self.status, wraplength=720).grid(
            row=5, column=0, sticky="w"
        )
        progress = ttk.LabelFrame(audit_tab, text="Progresso geral", padding=8)
        progress.grid(row=6, column=0, sticky="ew", pady=(10, 2))
        progress.columnconfigure(0, weight=1)
        self.progress_value = tk.DoubleVar(value=0)
        self.progress_bar = ttk.Progressbar(
            progress, variable=self.progress_value, maximum=100, mode="determinate"
        )
        self.progress_bar.grid(row=0, column=0, sticky="ew")
        self.progress_text = tk.StringVar(value="0 / 0 (0%) | Tempo total: 00:00")
        ttk.Label(progress, textvariable=self.progress_text).grid(row=1, column=0, sticky="w")

        current = ttk.LabelFrame(audit_tab, text="Versão atual: —", padding=8)
        current.grid(row=7, column=0, sticky="ew", pady=(4, 0))
        current.columnconfigure(0, weight=1)
        self.current_version_frame = current
        self.version_progress_value = tk.DoubleVar(value=0)
        ttk.Progressbar(
            current, variable=self.version_progress_value, maximum=100, mode="determinate"
        ).grid(row=0, column=0, sticky="ew")
        self.version_stage_text = tk.StringVar(value="Aguardando.")
        ttk.Label(current, textvariable=self.version_stage_text).grid(row=1, column=0, sticky="w")
        self.version_timing_text = tk.StringVar(
            value="Tempo da versão: 00:00 | Média recente: calculando | Estimativa restante: calculando"
        )
        ttk.Label(current, textvariable=self.version_timing_text).grid(row=2, column=0, sticky="w")
        self._startup_log("widgets_auditoria")
        self._build_stored_tab(stored_tab)
        self._startup_log("widgets_auditorias_armazenadas")
        self._set_action_state()
        self._startup_log("configurar_callbacks_estado")
        self.refresh_stored()
        self._startup_log("refresh_inicial")

    def _build_stored_tab(self, tab: ttk.Frame) -> None:
        tab.columnconfigure(0, weight=1)
        tab.rowconfigure(1, weight=1)
        ttk.Label(tab, text="Auditorias armazenadas", font=("TkDefaultFont", 14, "bold")).grid(row=0, column=0, sticky="w")
        columns = ("nome", "caminho", "unique_id", "versao", "processadas", "alteracoes", "ultima", "checkpoint")
        self.stored_tree = ttk.Treeview(tab, columns=columns, show="headings", selectmode="browse")
        labels = ("Planilha", "Caminho", "UniqueId", "Última versão", "Versões", "Alterações", "Última auditoria", "Checkpoint")
        widths = (150, 210, 140, 95, 65, 70, 130, 95)
        for column, label, width in zip(columns, labels, widths):
            self.stored_tree.heading(column, text=label)
            self.stored_tree.column(column, width=width, minwidth=55)
        self.stored_tree.grid(row=1, column=0, sticky="nsew", pady=8)
        scroll = ttk.Scrollbar(tab, orient="vertical", command=self.stored_tree.yview)
        scroll.grid(row=1, column=1, sticky="ns", pady=8)
        self.stored_tree.configure(yscrollcommand=scroll.set)
        self.stored_tree.bind("<<TreeviewSelect>>", self._stored_selection_changed)
        actions = ttk.Frame(tab)
        actions.grid(row=2, column=0, columnspan=2, sticky="ew")
        definitions = (
            ("Atualizar", self.refresh_stored),
            ("Backup selecionado", self.backup_selected),
            ("Restaurar selecionado", self.restore_individual),
            ("Abrir relatório", self.open_stored_report),
            ("Excluir auditoria selecionada", self.delete_selected),
            ("Backup completo", self.backup_complete),
            ("Restaurar backup completo", self.restore_complete),
            ("Excluir todas as auditorias", self.delete_all_audits),
        )
        self.storage_buttons = []
        for index, (text, command) in enumerate(definitions):
            button = ttk.Button(actions, text=text, command=command)
            button.grid(row=index // 4, column=index % 4, sticky="ew", padx=(0, 5), pady=3)
            self.storage_buttons.append(button)
            if text == "Abrir relatório":
                self.open_stored_report_button = button
        self._stored_selection_changed()

    def _stored_selection_changed(self, _event: object = None) -> None:
        if hasattr(self, "open_stored_report_button"):
            state = "normal" if self.stored_tree.selection() and not self._busy else "disabled"
            self.open_stored_report_button.configure(state=state)

    def _hidden_comparator_gesture(self, _event: object = None) -> None:
        now = time.monotonic()
        self._hidden_clicks = [click for click in self._hidden_clicks if now - click <= 2.0]
        self._hidden_clicks.append(now)
        if len(self._hidden_clicks) >= 5:
            self._hidden_clicks.clear()
            self._show_comparator()

    def _show_comparator(self) -> None:
        from app.spreadsheet_comparator import ComparatorFrame

        if self._comparator_tab is None:
            self._comparator_tab = ComparatorFrame(
                self.notebook, on_close=self._hide_comparator
            )
            self.notebook.add(self._comparator_tab, text="Comparador de Planilhas")
        self.notebook.select(self._comparator_tab)

    def _hide_comparator(self) -> None:
        if self._comparator_tab is None:
            return
        self.notebook.select(self._audit_tab)
        self.notebook.forget(self._comparator_tab)
        self._comparator_tab.destroy()
        self._comparator_tab = None

    def refresh_stored(self) -> None:
        if not hasattr(self, "stored_tree"):
            return
        for item in self.stored_tree.get_children():
            self.stored_tree.delete(item)
        query_started = time.perf_counter()
        audits = self.storage.list_audits()
        query_finished = time.perf_counter()
        if not getattr(self, "_startup_complete", True):
            logger.info(
                "STARTUP_INTERFACE etapa=consultar_banco_auditorias duracao=%.3fs acumulado=%.3fs",
                query_finished - query_started,
                query_finished - self._startup_started,
            )
            self._startup_last = query_finished
        for audit in audits:
            self.stored_tree.insert("", "end", iid=str(audit.id), values=(
                audit.name, audit.path or "—", audit.unique_id,
                audit.last_version or "—", audit.processed_versions, audit.changes,
                audit.last_audit or "—", audit.checkpoint or "—"))
        if not getattr(self, "_startup_complete", True):
            self._startup_log("popular_auditorias_armazenadas")

    def _selected_stored(self) -> tuple[int, str]:
        selection = self.stored_tree.selection()
        if not selection:
            raise ValueError("Selecione uma auditoria armazenada.")
        item = selection[0]
        return int(item), str(self.stored_tree.item(item, "values")[0])

    def backup_selected(self) -> None:
        try:
            spreadsheet_id, _ = self._selected_stored()
        except ValueError as error:
            messagebox.showinfo("Auditorias armazenadas", str(error))
            return
        self._start_work("Criando backup individual...", lambda: self.storage.backup_individual(spreadsheet_id), self._storage_finished)

    def backup_complete(self) -> None:
        self._start_work("Criando backup completo...", self.storage.backup_full, self._storage_finished)

    def restore_individual(self) -> None:
        path = filedialog.askopenfilename(title="Selecionar backup individual", filetypes=(("Backup SQLite", "*.sqlite3"),))
        if not path:
            return
        self._start_work(
            "Validando e restaurando backup individual...",
            lambda: self._try_restore_individual(path),
            self._individual_restore_finished,
        )

    def _try_restore_individual(self, path: str) -> tuple[str, str]:
        try:
            self.storage.restore_individual(path)
        except RestoreConflictError:
            return "conflict", path
        return "restored", path

    def _individual_restore_finished(self, result: tuple[str, str]) -> None:
        state, path = result
        if state == "conflict":
            if not messagebox.askyesno("Conflito de restauração", "Já existe auditoria para esta planilha. Substituir integralmente pelos dados do backup? Nenhum merge será realizado."):
                self.status.set("Restauração cancelada; dados locais preservados.")
                return
            self._start_work("Restaurando backup individual...", lambda: self.storage.restore_individual(path, replace=True), self._storage_finished)
            return
        self._storage_finished(Path(path))

    def delete_selected(self) -> None:
        try:
            spreadsheet_id, name = self._selected_stored()
        except ValueError as error:
            messagebox.showinfo("Auditorias armazenadas", str(error))
            return
        if not messagebox.askyesno("Excluir auditoria local", f"Excluir permanentemente a auditoria local de '{name}'?\n\nO arquivo no SharePoint não será alterado."):
            return
        def operation() -> None:
            report = self.report_artifacts.locate(spreadsheet_id)
            self.report_artifacts.delete_with_database(
                [report] if report else [],
                lambda: self.storage.delete_individual(spreadsheet_id),
            )
        self._start_work("Excluindo auditoria local...", operation, self._storage_finished)

    def open_stored_report(self) -> None:
        try:
            spreadsheet_id, _ = self._selected_stored()
            report = self.report_artifacts.locate(spreadsheet_id)
        except ValueError as error:
            messagebox.showinfo("Auditorias armazenadas", str(error))
            return
        if report is None:
            messagebox.showinfo(
                "Auditorias armazenadas",
                "Não existe relatório gerado para esta auditoria.",
            )
            return
        try:
            self._open_file(report)
        except OSError as error:
            logger.exception("Falha ao abrir relatório armazenado arquivo=%s", report)
            messagebox.showerror("Auditorias armazenadas", f"Não foi possível abrir o relatório: {error}")

    def restore_complete(self) -> None:
        path = filedialog.askopenfilename(title="Selecionar backup completo", filetypes=(("Backup SQLite", "*.sqlite3"),))
        if not path or not messagebox.askyesno("Restaurar banco completo", "Validar e substituir TODO o estado local pelo backup selecionado? Um backup de segurança será criado antes da substituição."):
            return
        self._start_work("Restaurando backup completo...", lambda: self.storage.restore_full(path), self._storage_finished)

    def delete_all_audits(self) -> None:
        if not messagebox.askyesno("Excluir todas as auditorias", "Esta operação excluirá permanentemente todas as auditorias armazenadas localmente. Os arquivos do SharePoint não serão alterados.\n\nDeseja continuar?"):
            return
        identities = [self.report_artifacts.identity(audit.id) for audit in self.storage.list_audits()]
        reports = self.report_artifacts.controlled_paths(identities)
        backup = messagebox.askyesno("Backup de proteção", "Fazer backup completo antes de excluir?")
        def operation() -> object:
            backup_path = self.storage.backup_full() if backup else None
            self.report_artifacts.delete_with_database(reports, self.storage.delete_all)
            return backup_path
        self._start_work("Excluindo todas as auditorias locais...", operation, self._storage_finished)

    def _storage_finished(self, result: object) -> None:
        self.refresh_stored()
        self.status.set(f"Operação local concluída{f': {result}' if isinstance(result, Path) else '.'}")

    def _configured_values(self) -> tuple[str, tuple[str, ...]]:
        site_url = self.site_url.get().strip().rstrip("/")
        scopes = tuple(
            dict.fromkeys(
                part.strip()
                for part in self.scope_paths.get().split(";")
                if part.strip()
            )
        )
        return site_url, scopes

    def _folders_loaded(self, folders: list[tuple[str, str]]) -> None:
        self.folder_paths = [path for _name, path in folders]
        self.folder_selector["values"] = [name for name, _path in folders]
        if folders:
            self.folder_selector.current(0)

    def copy_folder_to_scope(self) -> None:
        index = self.folder_selector.current()
        if index < 0 or index >= len(self.folder_paths):
            self.status.set("Selecione uma pasta.")
            return
        scope = self.folder_paths[index]
        self.scope_paths.set(scope)
        if self.source is not None:
            setter = getattr(self.source, "set_scope_paths", None)
            if setter is not None:
                setter((scope,))
        self.status.set("Escopo atualizado. Clique em Atualizar lista.")

    def connect(self) -> None:
        if self.connect_source is None:
            self.status.set("Conexão SharePoint não está disponível.")
            return
        connect_source = self.connect_source
        site_url, scopes = self._configured_values()
        try:
            if self.save_configuration is not None:
                self.save_configuration(site_url, scopes)
        except (OSError, ValueError) as error:
            self.status.set(f"Configuração inválida: {error}")
            return
        self._start_work(
            "Conectando ao SharePoint... Conclua o login/MFA no Edge.",
            lambda: connect_source(site_url, scopes),
            self._connected,
        )

    def _connected(self, source: VersionSource) -> None:
        previous = self.source
        self.source = source
        self._version_cache.clear()
        if previous is not None and previous is not source:
            try:
                previous.close()  # type: ignore[attr-defined]
            except (AttributeError, RuntimeError):
                logger.warning(
                    "Falha ao fechar sessão SharePoint anterior", exc_info=True
                )
        self.status.set("Conectado ao SharePoint. Clique em Atualizar lista.")
        self._set_action_state()

    def refresh(self) -> None:
        if self.source is None:
            self.status.set("Conecte ao SharePoint antes de atualizar a lista.")
            return
        source = self.source
        set_status_callback = getattr(source, "set_status_callback", None)
        if callable(set_status_callback):
            set_status_callback(self._report_updates.put)
        self._version_cache.clear()
        self._start_work(
            "Consultando planilhas no SharePoint...",
            lambda: list(source.list_spreadsheets()),
            self._refresh_finished,
        )

    def _refresh_finished(self, spreadsheets: list[SpreadsheetInfo]) -> None:
        self.spreadsheets = spreadsheets
        self.selector["values"] = [
            f"{item.name} — {item.path or item.drive_item_id}"
            for item in self.spreadsheets
        ]
        if self.spreadsheets:
            self.selector.current(0)
            self.show_status()
        else:
            self.status.set(
                "Nenhuma planilha .xlsx encontrada nos escopos configurados."
            )

    def _selected(self) -> SpreadsheetInfo:
        index = self.selector.current()
        if index < 0:
            raise ValueError("Selecione uma planilha")
        return self.spreadsheets[index]

    @staticmethod
    def _version_cache_key(spreadsheet: SpreadsheetInfo) -> str:
        return f"{spreadsheet.site_id}|{spreadsheet.drive_id}|{spreadsheet.drive_item_id}"

    def _cached_versions(
        self, spreadsheet: SpreadsheetInfo
    ) -> tuple[VersionInfo, ...] | None:
        key = self._version_cache_key(spreadsheet)
        cached = self._version_cache.get(key)
        if cached is None:
            return None
        created_at, versions = cached
        if time.monotonic() - created_at > self._version_cache_ttl:
            self._version_cache.pop(key, None)
            return None
        return versions

    def _store_versions(
        self, spreadsheet: SpreadsheetInfo, versions: tuple[VersionInfo, ...]
    ) -> None:
        self._version_cache[self._version_cache_key(spreadsheet)] = (
            time.monotonic(),
            versions,
        )

    def _database_row(self, spreadsheet: SpreadsheetInfo):
        return self.database.connection.execute(
            """SELECT p.id, c.versao_id, c.versao_numero FROM planilha p
               LEFT JOIN checkpoint c ON c.planilha_id = p.id
               WHERE p.site_id=? AND p.drive_id=? AND p.drive_item_id=?""",
            (spreadsheet.site_id, spreadsheet.drive_id, spreadsheet.drive_item_id),
        ).fetchone()

    def show_status(self) -> None:
        try:
            spreadsheet = self._selected()
        except ValueError as error:
            self.status.set(str(error))
            return

        if self._cached_versions(spreadsheet) is None:
            self.status.set("Carregando catálogo local...")

        self._start_work(
            "Verificando novas versões no SharePoint...",
            lambda: self._spreadsheet_status(spreadsheet),
            self._show_status_finished,
        )

    def _spreadsheet_status(self, spreadsheet: SpreadsheetInfo):
        if self.source is None:
            raise RuntimeError("Conecte ao SharePoint antes de consultar versões.")
        row = self._database_row(spreadsheet)
        checkpoint = row["versao_numero"] if row else None
        cached = self._cached_versions(spreadsheet)
        if cached is None:
            catalog = getattr(self, "version_catalog", None)
            if catalog is None:
                catalog = VersionCatalog(self.database)
                self.version_catalog = catalog
            local = catalog.load(spreadsheet)
            if local is not None:
                self._report_updates.put(
                    f"{len(local):,} versões carregadas localmente. Verificando novas versões..."
                )
            result = catalog.sync(
                spreadsheet, self.source, self._version_scan_updates.put
            )
            versions = result.versions
            if result.source == "local_only":
                self._report_updates.put("Catálogo atualizado — nenhuma versão nova.")
            elif result.source == "delta":
                self._report_updates.put(f"{result.new_versions} novas versões encontradas.")
            self._store_versions(spreadsheet, versions)
        else:
            versions = cached
            logger.info(
                "Lista de versões reutilizada do cache para status planilha=%s total=%d",
                spreadsheet.name,
                len(versions),
            )
        latest = versions[-1].number if versions else "—"
        ids = [version.id for version in versions]
        pending = max(len(versions) - 1, 0)
        if row and checkpoint:
            checkpoint_row = self.database.connection.execute(
                "SELECT versao_id FROM checkpoint WHERE planilha_id=?", (row["id"],)
            ).fetchone()
            if checkpoint_row and checkpoint_row["versao_id"] in ids:
                pending = max(
                    len(versions) - ids.index(checkpoint_row["versao_id"]) - 1, 0
                )
        return checkpoint, latest, pending, len(versions)

    def _show_status_finished(
        self, result: tuple[str | None, str, int, int]
    ) -> None:
        checkpoint, latest, pending, total_versions = result
        self._end_version_scan(total_versions=total_versions)
        self._latest_available = latest
        self._set_audit_details(checkpoint, pending)
        self.audit_button.configure(
            text="Continuar auditoria" if checkpoint else "Auditar histórico"
        )
        self.status.set("Pronto.")

    def audit(self) -> None:
        from app.audit_service import AuditService

        try:
            spreadsheet = self._selected()
        except ValueError as error:
            self.status.set(str(error))
            return
        if self.source is None:
            self.status.set("Conecte ao SharePoint antes de auditar.")
            return
        source = self.source
        self._pause_event.clear()
        self._stop_event.clear()
        self._audit_paused = False
        self._audit_active = True
        set_status_callback = getattr(source, "set_status_callback", None)
        if callable(set_status_callback):
            set_status_callback(self._report_updates.put)
        cached_versions = self._cached_versions(spreadsheet)
        self.progress_bar.stop()
        self.progress_bar.configure(mode="determinate")
        self._audit_started_at = time.monotonic()
        self._progress_completed = 0
        self._progress_total = 0
        self.progress_value.set(0)
        self.progress_text.set("0 / 0 (0%) | Tempo total: 00:00")
        self._current_version_started_at = None
        self._current_version = "—"
        self._smooth_version_progress = SmoothVersionProgress()
        self._version_last_animation = 0.0
        self.version_progress_value.set(0)
        self.version_stage_text.set("Preparando auditoria...")
        self._update_version_timing()

        def report_progress(completed: int, total: int) -> None:
            self._progress_updates.put((completed, total))

        def report_checkpoint(version: str, completed: int, pending: int) -> None:
            self._checkpoint_updates.put((version, completed, pending))

        def report_version_progress(event: object) -> None:
            self._version_progress_updates.put(event)

        def report_control(state: str, checkpoint: str | None) -> None:
            self._control_updates.put((state, checkpoint))

        self._start_work(
            "Auditoria em andamento...",
            lambda: AuditService(
                self.database,
                source,
                progress_callback=report_progress,
                checkpoint_callback=report_checkpoint,
                version_progress_callback=report_version_progress,
                pause_event=self._pause_event,
                stop_event=self._stop_event,
                control_callback=report_control,
            ).audit(spreadsheet, versions=cached_versions),
            self._audit_finished,
        )

    def _audit_finished(self, result: AuditResult) -> None:
        self._audit_active = False
        self._audit_paused = False
        self._pause_event.clear()
        self._stop_event.clear()
        self._set_action_state()
        if result.status is AuditExecutionStatus.FAILED:
            self._finish_progress(failed=True)
        elif result.status is AuditExecutionStatus.STOPPED:
            self._finish_progress(failed=True)
            self.status.set(
                "Auditoria interrompida pelo usuário no checkpoint "
                f"{result.final_version or 'nenhum'}."
            )
            self.audit_button.configure(
                text="Continuar auditoria" if result.final_version else "Auditar histórico"
            )
            self.refresh_stored()
            return
        else:
            self._finish_progress()
        self.status.set(
            f"{result.status.value}: {result.processed_versions} versões, "
            f"{result.changes} alterações."
        )
        self.show_status()

    def toggle_pause(self) -> None:
        if not self._audit_active:
            return
        if self._audit_paused or self._pause_event.is_set():
            self._pause_event.clear()
            self._audit_paused = False
            self.pause_button.configure(text="Pausar")
            self.status.set("Auditoria retomada.")
        else:
            self._pause_event.set()
            self.status.set(
                "Solicitação de pausa recebida. Finalizando a versão atual..."
            )

    def stop_audit(self) -> None:
        if not self._audit_active:
            return
        self._stop_event.set()
        # Desbloqueia imediatamente uma worker que já esteja aguardando em pausa.
        self._pause_event.clear()
        self.status.set(
            "Solicitação de parada recebida. Finalizando a versão atual..."
        )
        self.pause_button.configure(state="disabled")
        self.stop_button.configure(state="disabled")

    def generate_report(self) -> None:
        from app.report_service import ReportService

        try:
            spreadsheet = self._selected()
        except ValueError as error:
            self.status.set(str(error))
            return
        row = self._database_row(spreadsheet)
        if row is None:
            messagebox.showinfo(
                "Relatório", "Audite a planilha antes de gerar o relatório."
            )
            return
        self._start_work(
            "Gerando relatório...",
            lambda: ReportService(
                self.database.connection, self.reports_directory,
                progress_callback=self._report_updates.put,
            ).generate(row["id"]),
            self._report_finished,
        )

    def _report_finished(self, report: Path) -> None:
        self.last_report = report
        self.status.set(f"Relatório gerado: {self.last_report}")

    def _start_work(
        self,
        message: str,
        operation: Callable[[], object],
        finished: Callable,
    ) -> None:
        if self._busy:
            self.status.set("Aguarde a operação atual terminar.")
            return
        self._busy = True
        self.status.set(message)
        self._set_action_state()

        def worker() -> None:
            try:
                result = operation()
            except Exception as error:
                logger.warning(
                    "Operação da interface falhou tipo_erro=%s erro=%s",
                    type(error).__name__,
                    error,
                    exc_info=True,
                )
                self._work_results.put((False, error))
            else:
                self._work_results.put((True, result))

        # Tk, including ``after``, is only accessed by the main thread.  The
        # worker communicates exclusively through this queue.
        self.after(50, self._poll_work_result, finished)
        threading.Thread(target=worker, daemon=True).start()

    def _poll_work_result(self, finished: Callable) -> None:
        self._poll_progress_updates()
        self._poll_checkpoint_updates()
        self._poll_version_progress_updates()
        self._poll_version_scan_updates()
        self._poll_report_updates()
        self._poll_control_updates()
        try:
            succeeded, result = self._work_results.get_nowait()
        except queue.Empty:
            self.after(50, self._poll_work_result, finished)
            return
        if succeeded:
            self._work_finished(finished, result)
        else:
            assert isinstance(result, Exception)
            self._work_failed(result)

    def _work_finished(self, finished: Callable, result: object) -> None:
        self._busy = False
        self._set_action_state()
        finished(result)

    def _work_failed(self, error: Exception) -> None:
        self._busy = False
        self._audit_active = False
        self._audit_paused = False
        pause_event = getattr(self, "_pause_event", None)
        stop_event = getattr(self, "_stop_event", None)
        if pause_event is not None:
            pause_event.clear()
        if stop_event is not None:
            stop_event.clear()
        self._set_action_state()
        if self._version_scan_active:
            self._end_version_scan(failed=True)
        if getattr(self, "_audit_started_at", None) is not None:
            elapsed = time.monotonic() - self._audit_started_at
            self.progress_text.set(
                f"Auditoria interrompida | Tempo total: {self._format_duration(elapsed)}"
            )
            self._audit_started_at = None
        self.status.set(f"Falha na operação: {error}")

    def _poll_control_updates(self) -> None:
        if not hasattr(self, "_control_updates"):
            return
        while True:
            try:
                state, checkpoint = self._control_updates.get_nowait()
            except queue.Empty:
                break
            if state == "paused":
                self._audit_paused = True
                self.pause_button.configure(text="Continuar")
                self.status.set(
                    f"Auditoria pausada no checkpoint {checkpoint or 'nenhum'}."
                )
            elif state == "resumed":
                self._audit_paused = False
                self.pause_button.configure(text="Pausar")
                self.status.set("Auditoria retomada.")

    def _begin_version_scan(self, checkpoint_label: str | None = None) -> None:
        """Mostra atividade contínua enquanto o total de versões ainda é desconhecido."""
        self._version_scan_started_at = time.monotonic()
        self._version_scan_active = True
        self._version_scan_checkpoint_label = checkpoint_label
        while True:
            try:
                self._version_scan_updates.get_nowait()
            except queue.Empty:
                break
        self.progress_bar.stop()
        self.progress_bar.configure(mode="indeterminate")
        self.progress_value.set(0)
        self.progress_bar.start(12)
        prefix = (
            f"Buscando versões após o checkpoint {checkpoint_label}..."
            if checkpoint_label
            else "Buscando versões no SharePoint..."
        )
        self.progress_text.set(
            f"{prefix} Versões encontradas: 0 | "
            "Tempo: 00:00"
        )

    def _poll_version_scan_updates(self) -> None:
        if not self._version_scan_active:
            return
        latest_count: int | None = None
        while True:
            try:
                latest_count = self._version_scan_updates.get_nowait()
            except queue.Empty:
                break
        if latest_count is None:
            return
        elapsed = (
            time.monotonic() - self._version_scan_started_at
            if self._version_scan_started_at is not None
            else 0.0
        )
        prefix = (
            f"Buscando versões após o checkpoint "
            f"{self._version_scan_checkpoint_label}... "
            if getattr(self, "_version_scan_checkpoint_label", None)
            else "Buscando versões no SharePoint... "
        )
        self.progress_text.set(
            prefix + f"Versões encontradas: {latest_count:,} | "
            f"Tempo: {self._format_duration(elapsed)}"
        )

    def _end_version_scan(
        self, *, total_versions: int | None = None, failed: bool = False
    ) -> None:
        if not self._version_scan_active:
            return
        elapsed = (
            time.monotonic() - self._version_scan_started_at
            if self._version_scan_started_at is not None
            else 0.0
        )
        self.progress_bar.stop()
        self.progress_bar.configure(mode="determinate")
        self.progress_value.set(0 if failed else 100)
        if failed:
            self.progress_text.set(
                "Carregamento do histórico interrompido | "
                f"Tempo: {self._format_duration(elapsed)}"
            )
        else:
            total_text = (
                f"{total_versions:,} versões encontradas"
                if total_versions is not None
                else "histórico carregado"
            )
            self.progress_text.set(
                f"Histórico do SharePoint: {total_text} | "
                f"Tempo: {self._format_duration(elapsed)}"
            )
        self._version_scan_active = False
        self._version_scan_started_at = None
        self._version_scan_checkpoint_label = None

    def _poll_progress_updates(self) -> None:
        if not hasattr(self, "_progress_updates"):
            return
        while True:
            try:
                completed, total = self._progress_updates.get_nowait()
            except queue.Empty:
                break
            self._update_progress(completed, total)
        if getattr(self, "_audit_started_at", None) is not None:
            self._update_progress(self._progress_completed, self._progress_total)

    def _poll_version_progress_updates(self) -> None:
        """Aplica eventos da worker exclusivamente pela thread principal do Tk."""
        if not hasattr(self, "_version_progress_updates"):
            return
        while True:
            try:
                event = self._version_progress_updates.get_nowait()
            except queue.Empty:
                break
            now = time.monotonic()
            occurred_at = getattr(event, "occurred_at", now)
            if event.percent == 0:
                self._current_version = event.version
                self._current_version_started_at = occurred_at
                self.current_version_frame.configure(text=f"Versão atual: {event.version}")
            value = self._smooth_version_progress.observe(
                event.version, event.percent, occurred_at
            )
            self.version_progress_value.set(value)
            self.version_stage_text.set(event.stage)
            if event.percent == 100 and self._current_version_started_at is not None:
                self._current_version_started_at = None
            self._update_version_timing(now)
        now = time.monotonic()
        if (
            self._current_version_started_at is not None
            and now - self._version_last_animation >= self._version_animation_interval
        ):
            self.version_progress_value.set(self._smooth_version_progress.estimate(now))
            self._version_last_animation = now
        if self._current_version_started_at is not None and now - self._version_last_clock_update >= 0.5:
            self._update_version_timing(now)

    def _update_version_timing(self, now: float | None = None) -> None:
        now = time.monotonic() if now is None else now
        self._version_last_clock_update = now
        elapsed = (
            now - self._current_version_started_at
            if self._current_version_started_at is not None
            else 0.0
        )
        average = self._smooth_version_progress.total_average()
        if average is not None:
            remaining = average * max(self._progress_total - self._progress_completed, 0)
            average_text = self._format_duration(average)
            eta_text = self._format_duration(remaining)
        else:
            average_text = eta_text = "calculando"
        self.version_timing_text.set(
            f"Tempo da versão: {self._format_duration(elapsed)} | "
            f"Média recente: {average_text} | Estimativa restante: {eta_text}"
        )

    def _poll_checkpoint_updates(self) -> None:
        """Transfere checkpoints confirmados da worker para as variáveis Tk."""
        if not hasattr(self, "_checkpoint_updates"):
            return
        while True:
            try:
                version, _completed, pending = self._checkpoint_updates.get_nowait()
            except queue.Empty:
                break
            self._set_audit_details(version, pending)

    def _set_audit_details(self, checkpoint: str | None, pending: int) -> None:
        self.details.set(
            f"Última auditada: {checkpoint or '—'} | "
            f"Última disponível: {self._latest_available} | Pendentes: {pending}"
        )

    def _poll_report_updates(self) -> None:
        """Transfere para o Tk, na main thread, somente a mensagem mais recente."""
        if not hasattr(self, "_report_updates"):
            return
        latest: str | None = None
        while True:
            try:
                latest = self._report_updates.get_nowait()
            except queue.Empty:
                break
        if latest is not None:
            self.status.set(latest)
            if getattr(self, "_current_version_started_at", None) is not None:
                self.version_stage_text.set(latest)

    def _update_progress(self, completed: int, total: int) -> None:
        if self._audit_started_at is None:
            return
        self._progress_completed = completed
        self._progress_total = total
        elapsed = time.monotonic() - self._audit_started_at
        percent = 0 if total <= 0 else completed / total * 100
        self.progress_value.set(percent)
        if completed > 0 and completed < total:
            remaining = elapsed / completed * (total - completed)
            estimate = self._format_duration(remaining)
        elif total > 0 and completed >= total:
            estimate = "00:00"
        else:
            estimate = "calculando"
        self.progress_text.set(
            f"{completed} / {total} ({percent:.0f}%) | "
            f"Estimativa: {estimate} | Tempo total: {self._format_duration(elapsed)}"
        )
        if hasattr(self, "version_timing_text"):
            self._update_version_timing()

    def _finish_progress(self, *, failed: bool = False) -> None:
        if self._audit_started_at is None:
            return
        elapsed = time.monotonic() - self._audit_started_at
        if failed:
            self.progress_text.set(
                "Auditoria interrompida | "
                f"Tempo total: {self._format_duration(elapsed)}"
            )
        else:
            self.progress_value.set(100)
            self.progress_text.set(
                f"Progresso da auditoria: concluída (100%) | "
                f"Tempo total: {self._format_duration(elapsed)}"
            )
        self._audit_started_at = None

    @staticmethod
    def _format_duration(seconds: float) -> str:
        total_seconds = max(0, round(seconds))
        hours, remainder = divmod(total_seconds, 3600)
        minutes, seconds = divmod(remainder, 60)
        if hours:
            return f"{hours:02d}:{minutes:02d}:{seconds:02d}"
        return f"{minutes:02d}:{seconds:02d}"

    def _set_action_state(self) -> None:
        state = "disabled" if self._busy else "normal"
        connected_state = state if self.source is not None else "disabled"
        for name in ("connect_button",):
            if hasattr(self, name):
                getattr(self, name).configure(state=state)
        if hasattr(self, "copy_folder_button"):
            self.copy_folder_button.configure(state=connected_state)
        for name in ("refresh_button", "audit_button", "report_button"):
            if hasattr(self, name):
                getattr(self, name).configure(state=connected_state)
        if hasattr(self, "pause_button"):
            self.pause_button.configure(
                text="Continuar" if self._audit_paused else "Pausar",
                state="normal" if self._audit_active and not self._stop_event.is_set() else "disabled",
            )
        if hasattr(self, "stop_button"):
            self.stop_button.configure(
                state="normal" if self._audit_active and not self._stop_event.is_set() else "disabled"
            )
        for button in getattr(self, "storage_buttons", ()):
            button.configure(state=state)
        self._stored_selection_changed()

    def close_source(self) -> None:
        if self.source is not None:
            close = getattr(self.source, "close", None)
            if close is not None:
                close()
            self.source = None

    def request_close(self) -> None:
        """Fecha a janela somente quando não há operação em andamento.

        As operações usam uma thread daemon e compartilham a conexão SQLite e a
        fonte SharePoint com a interface. Destruir a janela durante esse trabalho
        faria o bloco ``finally`` do processo fechar esses recursos enquanto a
        thread ainda os utiliza. Sem cancelamento cooperativo, a alternativa
        segura é manter a aplicação aberta até a unidade atual terminar.
        """
        if self._busy:
            messagebox.showwarning(
                "Operação em andamento",
                "Aguarde a operação atual terminar antes de fechar a aplicação. "
                "Isso preserva o checkpoint e os arquivos temporários.",
            )
            return
        self.winfo_toplevel().destroy()

    def open_report(self) -> None:
        if self.last_report is None or not self.last_report.exists():
            messagebox.showinfo("Relatório", "Gere o relatório primeiro.")
            return
        self._open_file(self.last_report)

    @staticmethod
    def _open_file(path: Path) -> None:
        if sys.platform == "win32":
            import os

            os.startfile(path)  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.Popen(("open", str(path)))
        else:
            subprocess.Popen(("xdg-open", str(path)))
