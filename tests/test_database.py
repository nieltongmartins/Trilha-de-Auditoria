import sqlite3
from pathlib import Path

import pytest

from app.database import Database, SCHEMA_VERSION, VERSION_PROCESSED_MIGRATIONS
from app.exceptions import DatabaseNotConnectedError


EXPECTED_TABLES = {
    "planilha",
    "checkpoint",
    "versao_processada",
    "alteracao",
    "execucao_auditoria",
    "erro_processamento",
}


@pytest.fixture
def database(tmp_path: Path) -> Database:
    with Database(tmp_path / "database" / "audit.db") as instance:
        instance.initialize()
        yield instance


def insert_spreadsheet(connection: sqlite3.Connection) -> int:
    cursor = connection.execute(
        """
        INSERT INTO planilha (drive_item_id, nome_atual, site_id, drive_id)
        VALUES (?, ?, ?, ?)
        """,
        ("item-1", "CQL028.xlsx", "site-1", "drive-1"),
    )
    connection.commit()
    assert cursor.lastrowid is not None
    return cursor.lastrowid


def test_initialize_creates_database_and_all_official_tables(tmp_path: Path) -> None:
    path = tmp_path / "nested" / "audit.db"

    with Database(path) as database:
        database.initialize()
        tables = {
            row[0]
            for row in database.connection.execute(
                "SELECT name FROM sqlite_master WHERE type = 'table'"
            )
        }

    assert path.is_file()
    assert EXPECTED_TABLES <= tables


def test_repeated_initialization_preserves_existing_data(database: Database) -> None:
    spreadsheet_id = insert_spreadsheet(database.connection)

    database.initialize()

    row = database.connection.execute(
        "SELECT nome_atual FROM planilha WHERE id = ?", (spreadsheet_id,)
    ).fetchone()
    assert row["nome_atual"] == "CQL028.xlsx"


def create_legacy_database(path: Path) -> None:
    """Cria a tabela no formato anterior às quatro colunas introduzidas na F4."""
    with sqlite3.connect(path) as connection:
        connection.executescript(
            """
            CREATE TABLE versao_processada (
                id INTEGER PRIMARY KEY,
                planilha_id INTEGER NOT NULL,
                versao_anterior_id TEXT NOT NULL,
                versao_anterior_numero TEXT NOT NULL,
                versao_atual_id TEXT NOT NULL,
                versao_atual_numero TEXT NOT NULL,
                data_hora_versao TEXT,
                autor TEXT,
                comentario TEXT,
                tamanho INTEGER CHECK (tamanho IS NULL OR tamanho >= 0),
                quantidade_alteracoes INTEGER NOT NULL DEFAULT 0,
                status TEXT NOT NULL,
                data_processamento TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                execucao_id INTEGER NOT NULL,
                UNIQUE (planilha_id, versao_anterior_id, versao_atual_id)
            );
            INSERT INTO versao_processada (
                id, planilha_id, versao_anterior_id, versao_anterior_numero,
                versao_atual_id, versao_atual_numero, autor, comentario,
                tamanho, status, execucao_id
            ) VALUES (
                7, 2, '512', '1.0', '513', '1.1', 'Autor legado',
                'dado preservado', 19440, 'PROCESSADA', 11
            );
            """
        )


def version_columns(connection: sqlite3.Connection) -> dict[str, tuple[object, ...]]:
    return {
        row["name"]: (row["type"], row["notnull"], row["dflt_value"], row["pk"])
        for row in connection.execute("PRAGMA table_info(versao_processada)")
    }


def test_initialize_migrates_legacy_database_and_preserves_data(tmp_path: Path) -> None:
    path = tmp_path / "legacy.db"
    create_legacy_database(path)

    with Database(path) as database:
        database.initialize()
        columns = version_columns(database.connection)
        row = database.connection.execute(
            "SELECT * FROM versao_processada WHERE id = 7"
        ).fetchone()

        assert VERSION_PROCESSED_MIGRATIONS.keys() <= columns.keys()
        assert row["comentario"] == "dado preservado"
        assert row["autor_email"] is None
        assert row["autor_login"] is None
        assert row["url_origem"] is None
        assert row["versao_atual"] == 0
        assert row["hash_origem"] is None
        assert database.connection.execute("PRAGMA user_version").fetchone()[0] == 3


def test_schema_migration_is_idempotent(tmp_path: Path) -> None:
    path = tmp_path / "legacy.db"
    create_legacy_database(path)

    with Database(path) as database:
        database.initialize()
        columns_after_first_run = version_columns(database.connection)
        database.initialize()

        assert version_columns(database.connection) == columns_after_first_run
        assert (
            database.connection.execute(
                "SELECT COUNT(*) FROM versao_processada WHERE id = 7"
            ).fetchone()[0]
            == 1
        )


def test_new_and_migrated_databases_have_equivalent_required_fields(
    tmp_path: Path,
) -> None:
    legacy_path = tmp_path / "legacy.db"
    new_path = tmp_path / "new.db"
    create_legacy_database(legacy_path)

    with Database(legacy_path) as migrated, Database(new_path) as new:
        migrated.initialize()
        new.initialize()
        migrated_columns = version_columns(migrated.connection)
        new_columns = version_columns(new.connection)

        assert migrated_columns.keys() == new_columns.keys()
        for name in VERSION_PROCESSED_MIGRATIONS:
            assert migrated_columns[name] == new_columns[name]
        assert migrated.connection.execute("PRAGMA user_version").fetchone()[0] == (
            SCHEMA_VERSION
        )
        assert new.connection.execute("PRAGMA user_version").fetchone()[0] == (
            SCHEMA_VERSION
        )


def test_identity_and_checkpoint_constraints_prevent_duplicates(
    database: Database,
) -> None:
    spreadsheet_id = insert_spreadsheet(database.connection)

    with pytest.raises(sqlite3.IntegrityError):
        database.connection.execute(
            """
            INSERT INTO planilha (drive_item_id, nome_atual, site_id, drive_id)
            VALUES ('item-1', 'Renomeada.xlsx', 'site-1', 'drive-1')
            """
        )
    database.connection.rollback()

    database.connection.execute(
        """
        INSERT INTO checkpoint (planilha_id, versao_id, versao_numero)
        VALUES (?, 'version-1', '0.99')
        """,
        (spreadsheet_id,),
    )
    database.connection.commit()

    with pytest.raises(sqlite3.IntegrityError):
        database.connection.execute(
            """
            INSERT INTO checkpoint (planilha_id, versao_id, versao_numero)
            VALUES (?, 'version-2', '1.00')
            """,
            (spreadsheet_id,),
        )


def test_foreign_keys_and_controlled_values_are_enforced(database: Database) -> None:
    with pytest.raises(sqlite3.IntegrityError):
        database.connection.execute(
            """
            INSERT INTO checkpoint (planilha_id, versao_id, versao_numero)
            VALUES (999, 'version-1', '1.0')
            """
        )
    database.connection.rollback()

    spreadsheet_id = insert_spreadsheet(database.connection)
    with pytest.raises(sqlite3.IntegrityError):
        database.connection.execute(
            """
            INSERT INTO execucao_auditoria
                (codigo_execucao, planilha_id, status)
            VALUES ('AUD-1', ?, 'STATUS_INVALIDO')
            """,
            (spreadsheet_id,),
        )


def test_processed_comparison_and_change_cannot_be_duplicated(
    database: Database,
) -> None:
    spreadsheet_id = insert_spreadsheet(database.connection)
    execution_id = database.connection.execute(
        """
        INSERT INTO execucao_auditoria (codigo_execucao, planilha_id, status)
        VALUES ('AUD-1', ?, 'EM_EXECUCAO')
        """,
        (spreadsheet_id,),
    ).lastrowid
    version_values = (
        spreadsheet_id,
        "v099",
        "0.99",
        "v100",
        "1.00",
        execution_id,
    )
    version_id = database.connection.execute(
        """
        INSERT INTO versao_processada (
            planilha_id, versao_anterior_id, versao_anterior_numero,
            versao_atual_id, versao_atual_numero, status, execucao_id
        ) VALUES (?, ?, ?, ?, ?, 'PROCESSADA', ?)
        """,
        version_values,
    ).lastrowid
    assert version_id is not None
    database.connection.commit()

    with pytest.raises(sqlite3.IntegrityError):
        database.connection.execute(
            """
            INSERT INTO versao_processada (
                planilha_id, versao_anterior_id, versao_anterior_numero,
                versao_atual_id, versao_atual_numero, status, execucao_id
            ) VALUES (?, ?, ?, ?, ?, 'PROCESSADA', ?)
            """,
            version_values,
        )
    database.connection.rollback()

    change_values = (version_id, spreadsheet_id, "Plan1", "A1")
    database.connection.execute(
        """
        INSERT INTO alteracao (
            versao_processada_id, planilha_id, tipo, aba, endereco, valor_novo
        ) VALUES (?, ?, 'ADD', ?, ?, 'novo')
        """,
        change_values,
    )
    database.connection.commit()
    with pytest.raises(sqlite3.IntegrityError):
        database.connection.execute(
            """
            INSERT INTO alteracao (
                versao_processada_id, planilha_id, tipo, aba, endereco, valor_novo
            ) VALUES (?, ?, 'ADD', ?, ?, 'duplicado')
            """,
            change_values,
        )


def test_context_manager_closes_connection(tmp_path: Path) -> None:
    database = Database(tmp_path / "audit.db")

    with database:
        database.initialize()
        assert database.is_connected

    assert not database.is_connected
    with pytest.raises(DatabaseNotConnectedError):
        _ = database.connection
