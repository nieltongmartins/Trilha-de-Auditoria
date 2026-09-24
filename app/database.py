"""Criação e gerenciamento do banco SQLite oficial da auditoria."""

from __future__ import annotations

import sqlite3
from pathlib import Path
from types import TracebackType

from app.exceptions import DatabaseNotConnectedError


SCHEMA = """
CREATE TABLE IF NOT EXISTS planilha (
    id INTEGER PRIMARY KEY,
    drive_item_id TEXT NOT NULL,
    nome_atual TEXT NOT NULL,
    site_id TEXT NOT NULL,
    drive_id TEXT NOT NULL,
    caminho_sharepoint TEXT,
    data_cadastro TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
    data_atualizacao TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
    CONSTRAINT uq_planilha_identidade UNIQUE (site_id, drive_id, drive_item_id)
);

CREATE TABLE IF NOT EXISTS checkpoint (
    id INTEGER PRIMARY KEY,
    planilha_id INTEGER NOT NULL,
    versao_id TEXT NOT NULL,
    versao_numero TEXT NOT NULL,
    data_hora_versao TEXT,
    data_atualizacao TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
    CONSTRAINT uq_checkpoint_planilha UNIQUE (planilha_id),
    CONSTRAINT fk_checkpoint_planilha FOREIGN KEY (planilha_id)
        REFERENCES planilha (id) ON DELETE RESTRICT
);

CREATE TABLE IF NOT EXISTS execucao_auditoria (
    id INTEGER PRIMARY KEY,
    codigo_execucao TEXT NOT NULL UNIQUE,
    planilha_id INTEGER NOT NULL,
    inicio TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
    fim TEXT,
    checkpoint_inicial TEXT,
    versao_final TEXT,
    versoes_processadas INTEGER NOT NULL DEFAULT 0 CHECK (versoes_processadas >= 0),
    alteracoes_encontradas INTEGER NOT NULL DEFAULT 0 CHECK (alteracoes_encontradas >= 0),
    status TEXT NOT NULL CHECK (
        status IN ('EM_EXECUCAO', 'CONCLUIDA', 'CONCLUIDA_SEM_NOVIDADES', 'FALHA')
    ),
    mensagem TEXT,
    CONSTRAINT fk_execucao_planilha FOREIGN KEY (planilha_id)
        REFERENCES planilha (id) ON DELETE RESTRICT
);

CREATE TABLE IF NOT EXISTS versao_processada (
    id INTEGER PRIMARY KEY,
    planilha_id INTEGER NOT NULL,
    versao_anterior_id TEXT NOT NULL,
    versao_anterior_numero TEXT NOT NULL,
    versao_atual_id TEXT NOT NULL,
    versao_atual_numero TEXT NOT NULL,
    data_hora_versao TEXT,
    autor TEXT,
    autor_email TEXT,
    autor_login TEXT,
    comentario TEXT,
    tamanho INTEGER CHECK (tamanho IS NULL OR tamanho >= 0),
    url_origem TEXT,
    versao_atual INTEGER NOT NULL DEFAULT 0 CHECK (versao_atual IN (0, 1)),
    quantidade_alteracoes INTEGER NOT NULL DEFAULT 0 CHECK (quantidade_alteracoes >= 0),
    status TEXT NOT NULL CHECK (status IN ('PROCESSADA', 'SEM_ALTERACOES', 'ERRO')),
    hash_origem TEXT,
    data_processamento TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
    execucao_id INTEGER NOT NULL,
    CONSTRAINT uq_versao_comparacao UNIQUE (
        planilha_id, versao_anterior_id, versao_atual_id
    ),
    CONSTRAINT fk_versao_planilha FOREIGN KEY (planilha_id)
        REFERENCES planilha (id) ON DELETE RESTRICT,
    CONSTRAINT fk_versao_execucao FOREIGN KEY (execucao_id)
        REFERENCES execucao_auditoria (id) ON DELETE RESTRICT
);

CREATE TABLE IF NOT EXISTS alteracao (
    id INTEGER PRIMARY KEY,
    versao_processada_id INTEGER NOT NULL,
    planilha_id INTEGER NOT NULL,
    tipo TEXT NOT NULL CHECK (tipo IN ('ADD', 'DEL', 'MOD')),
    aba TEXT NOT NULL,
    endereco TEXT NOT NULL,
    valor_anterior TEXT,
    valor_novo TEXT,
    CONSTRAINT uq_alteracao_celula UNIQUE (versao_processada_id, aba, endereco),
    CONSTRAINT fk_alteracao_versao FOREIGN KEY (versao_processada_id)
        REFERENCES versao_processada (id) ON DELETE RESTRICT,
    CONSTRAINT fk_alteracao_planilha FOREIGN KEY (planilha_id)
        REFERENCES planilha (id) ON DELETE RESTRICT
);

CREATE TABLE IF NOT EXISTS erro_processamento (
    id INTEGER PRIMARY KEY,
    execucao_id INTEGER NOT NULL,
    planilha_id INTEGER NOT NULL,
    versao_anterior TEXT,
    versao_atual TEXT,
    tipo_erro TEXT NOT NULL,
    mensagem TEXT NOT NULL,
    data_hora TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
    CONSTRAINT fk_erro_execucao FOREIGN KEY (execucao_id)
        REFERENCES execucao_auditoria (id) ON DELETE RESTRICT,
    CONSTRAINT fk_erro_planilha FOREIGN KEY (planilha_id)
        REFERENCES planilha (id) ON DELETE RESTRICT
);

CREATE TABLE IF NOT EXISTS version_catalog (
    id INTEGER PRIMARY KEY,
    workbook_identity TEXT NOT NULL,
    technical_version_id TEXT NOT NULL,
    version_label TEXT NOT NULL,
    created_at_sharepoint TEXT,
    is_current_snapshot INTEGER NOT NULL DEFAULT 0
        CHECK (is_current_snapshot IN (0, 1)),
    discovered_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
    last_verified_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
    author TEXT,
    author_email TEXT,
    author_login TEXT,
    comment TEXT,
    size INTEGER CHECK (size IS NULL OR size >= 0),
    source_url TEXT,
    CONSTRAINT uq_version_catalog_identity
        UNIQUE (workbook_identity, technical_version_id)
);

CREATE TABLE IF NOT EXISTS version_catalog_state (
    workbook_identity TEXT PRIMARY KEY,
    last_historical_id TEXT,
    last_historical_label TEXT,
    last_known_current_id TEXT NOT NULL,
    last_known_current_label TEXT NOT NULL,
    catalog_count INTEGER NOT NULL CHECK (catalog_count >= 0),
    last_sync_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
    catalog_status TEXT NOT NULL CHECK (
        catalog_status IN ('VALID', 'NEEDS_RECONCILIATION', 'INVALID')
    )
);

CREATE INDEX IF NOT EXISTS idx_versao_planilha
    ON versao_processada (planilha_id, data_processamento);
CREATE INDEX IF NOT EXISTS idx_alteracao_planilha
    ON alteracao (planilha_id);
CREATE INDEX IF NOT EXISTS idx_execucao_planilha
    ON execucao_auditoria (planilha_id, inicio);
CREATE INDEX IF NOT EXISTS idx_erro_execucao
    ON erro_processamento (execucao_id);
CREATE INDEX IF NOT EXISTS idx_version_catalog_order
    ON version_catalog (workbook_identity, CAST(technical_version_id AS INTEGER));
"""

# Colunas acrescentadas ao modelo depois da criação dos primeiros bancos F1.
# CREATE TABLE IF NOT EXISTS não evolui uma tabela que já existe, portanto cada
# acréscimo precisa permanecer registrado como uma migração explícita.
SCHEMA_VERSION = 3
VERSION_PROCESSED_MIGRATIONS = {
    "autor_email": "TEXT",
    "autor_login": "TEXT",
    "url_origem": "TEXT",
    "versao_atual": "INTEGER NOT NULL DEFAULT 0 CHECK (versao_atual IN (0, 1))",
    "hash_origem": "TEXT",
}


class Database:
    """Controla uma conexão SQLite com integridade referencial habilitada."""

    def __init__(self, path: str | Path) -> None:
        self.path = Path(path)
        self._connection: sqlite3.Connection | None = None

    @property
    def connection(self) -> sqlite3.Connection:
        if self._connection is None:
            raise DatabaseNotConnectedError("A conexão com o banco não está aberta.")
        return self._connection

    @property
    def is_connected(self) -> bool:
        return self._connection is not None

    def connect(self) -> sqlite3.Connection:
        """Abre uma conexão e habilita constraints de chave estrangeira."""
        if self._connection is None:
            self.path.parent.mkdir(parents=True, exist_ok=True)
            # A interface executa uma auditoria por vez em uma thread de trabalho
            # para permanecer responsiva; o acesso continua serializado pela tela.
            self._connection = sqlite3.connect(self.path, check_same_thread=False)
            self._connection.row_factory = sqlite3.Row
            self._connection.execute("PRAGMA foreign_keys = ON")
        return self._connection

    def initialize(self) -> None:
        """Cria ou migra o esquema de modo idempotente, preservando os dados."""
        connection = self.connect()
        with connection:
            connection.executescript(SCHEMA)
            self._migrate_schema(connection)

    @staticmethod
    def _migrate_schema(connection: sqlite3.Connection) -> None:
        """Aplica evoluções aditivas ausentes no banco legado da F1."""
        connection.execute("BEGIN IMMEDIATE")
        columns = {
            row["name"]
            for row in connection.execute("PRAGMA table_info(versao_processada)")
        }
        for name, definition in VERSION_PROCESSED_MIGRATIONS.items():
            if name not in columns:
                connection.execute(
                    f'ALTER TABLE versao_processada ADD COLUMN "{name}" {definition}'
                )
        connection.execute(f"PRAGMA user_version = {SCHEMA_VERSION}")

    def close(self) -> None:
        """Encerra a conexão aberta."""
        if self._connection is not None:
            self._connection.close()
            self._connection = None

    def __enter__(self) -> "Database":
        self.connect()
        return self

    def __exit__(
        self,
        exc_type: type[BaseException] | None,
        exc_value: BaseException | None,
        traceback: TracebackType | None,
    ) -> None:
        self.close()
