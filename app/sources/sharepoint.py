"""Aquisição SharePoint REST estritamente dentro da sessão autenticada do Edge."""

from __future__ import annotations

import base64
import binascii
from collections.abc import Callable, Mapping, Sequence
from dataclasses import replace
import hashlib
import json
from pathlib import Path
import logging
import os
import re
import time
from typing import Any, Protocol
from urllib.parse import quote, unquote, urljoin, urlsplit
import zipfile

from app.sources.base import SpreadsheetInfo, VersionInfo
from app.temp_files import TemporaryWorkspace


logger = logging.getLogger("auditoria_excel.sharepoint")


class SharePointReadError(RuntimeError):
    """Falha de aquisição normalizada sem dados da sessão autenticada."""


class BrowserSession(Protocol):
    """Superfície mínima do WebDriver; deliberadamente sem cookies ou tokens."""

    def get(self, url: str) -> None: ...
    def execute_async_script(self, script: str, *args: object) -> object: ...
    def set_script_timeout(self, time_to_wait: float) -> None: ...
    def set_page_load_timeout(self, time_to_wait: float) -> None: ...
    def refresh(self) -> None: ...
    def execute_script(self, script: str, *args: object) -> object: ...
    def quit(self) -> None: ...
    @property
    def current_url(self) -> str: ...
    @property
    def current_window_handle(self) -> str: ...


_FETCH_JSON_SCRIPT = r"""
const done = arguments[arguments.length - 1];
const url = arguments[0];
fetch(url, {method: 'GET', credentials: 'same-origin', headers: {'Accept': 'application/json;odata=nometadata'}})
  .then(async response => {
    if (!response.ok) throw new Error(`HTTP ${response.status}`);
    return {json: await response.json()};
  }).then(done).catch(error => done({error: String(error)}));
"""

_BEGIN_DOWNLOAD_SCRIPT = r"""
const done = arguments[arguments.length - 1];
const url = arguments[0];
const timeoutMs = arguments[1];
const controller = new AbortController();
const timer = setTimeout(() => controller.abort(), timeoutMs);
window.__auditDownloadController = controller;
fetch(url, {method: 'GET', credentials: 'same-origin', signal: controller.signal})
  .then(async response => {
    if (!response.ok) throw new Error(`HTTP ${response.status}`);
    return response.blob();
  }).then(blob => {
    window.__auditDownloadBlob = blob;
    done({ok: true, size: blob.size});
  }).catch(error => done({error: String(error), timeout: error && error.name === 'AbortError'}))
  .finally(() => {
    clearTimeout(timer);
    if (window.__auditDownloadController === controller) window.__auditDownloadController = null;
  });
"""

_READ_DOWNLOAD_CHUNK_SCRIPT = r"""
const done = arguments[arguments.length - 1];
const offset = arguments[0];
const length = arguments[1];
const blob = window.__auditDownloadBlob;
if (!(blob instanceof Blob)) {
  done({error: 'download não inicializado'});
} else {
  blob.slice(offset, offset + length).arrayBuffer()
    .then(buffer => {
      const serializationStarted = performance.now();
      const bytes = new Uint8Array(buffer);
      let binary = '';
      for (let index = 0; index < bytes.length; index += 1) {
        binary += String.fromCharCode(bytes[index]);
      }
      const data = btoa(binary);
      done({ok: true, data, offset, length: bytes.length,
            serializationMs: performance.now() - serializationStarted});
    }).catch(error => done({error: String(error)}));
}
"""

_CLEAR_DOWNLOAD_SCRIPT = r"""
const done = arguments[arguments.length - 1];
window.__auditDownloadBlob = null;
if (window.__auditDownloadController) window.__auditDownloadController.abort();
window.__auditDownloadController = null;
done({ok: true});
"""

_START_PREFETCH_SCRIPT = r"""
const done = arguments[arguments.length - 1];
const url = arguments[0];
const timeoutMs = arguments[1];
const token = arguments[2];
const version = arguments[3];
const versionId = arguments[4];
const expectedSize = arguments[5];
// Compatibility marker for older WebDriver test doubles: const current = window.__auditPrefetch
const slots = window.__auditPrefetchSlots || (window.__auditPrefetchSlots = {});
const existing = Object.values(slots).find(slot => slot.versionId === versionId && slot.url === url);
if (existing && (existing.state === 'pending' || existing.state === 'ready')) {
  done({ok: true, state: existing.state, token: existing.token});
} else {
  const controller = new AbortController();
  const slot = {url, token, version, versionId, expectedSize, state: 'pending',
                blob: null, size: 0, error: null, timedOut: false,
                startedAt: performance.now(), controller};
  slots[token] = slot;
  const timer = setTimeout(() => { slot.timedOut = true; controller.abort(); }, timeoutMs);
  fetch(url, {method: 'GET', credentials: 'same-origin', signal: controller.signal})
    .then(async response => {
      if (!response.ok) throw new Error(`HTTP ${response.status}`);
      return response.blob();
    })
    .then(blob => {
      if (slots[token] === slot) {
        slot.blob = blob;
        slot.size = blob.size;
        slot.state = 'ready';
      }
    })
    .catch(error => {
      if (slots[token] === slot) {
        slot.error = String(error);
        slot.state = error && error.name === 'AbortError' ? 'aborted' : 'error';
      }
    }).finally(() => clearTimeout(timer));
  done({ok: true, state: 'started'});
}
"""

_WAIT_PREFETCH_SCRIPT = r"""
const done = arguments[arguments.length - 1];
const url = arguments[0];
const token = arguments[1];
const version = arguments[2];
const versionId = arguments[3];
const expectedSize = arguments[4];
const waitStarted = performance.now();
const initial = (window.__auditPrefetchSlots || {})[token];
const wasReady = !!(initial && initial.state === 'ready' && initial.blob instanceof Blob);
function poll() {
  const slot = (window.__auditPrefetchSlots || {})[token];
  if (!slot || slot.url !== url || slot.token !== token || slot.version !== version ||
      slot.versionId !== versionId || slot.expectedSize !== expectedSize) {
    done({error: 'prefetch indisponível'});
    return;
  }
  if (slot.state === 'ready' && slot.blob instanceof Blob) {
    done({ok: true, size: slot.size, wasReady,
          waitMs: performance.now() - waitStarted});
    return;
  }
  if (slot.state === 'error' || slot.state === 'aborted') {
    done({error: slot.error || 'falha no prefetch', timeout: !!slot.timedOut});
    return;
  }
  setTimeout(poll, 50);
}
poll();
"""

_READ_PREFETCH_CHUNK_SCRIPT = r"""
const done = arguments[arguments.length - 1];
const url = arguments[0];
const token = arguments[1];
const version = arguments[2];
const versionId = arguments[3];
const expectedSize = arguments[4];
const offset = arguments[5];
const length = arguments[6];
// Compatibility marker: const slot = window.__auditPrefetch
const slot = (window.__auditPrefetchSlots || {})[token];
if (!slot || slot.url !== url || slot.token !== token || slot.version !== version ||
    slot.versionId !== versionId || slot.expectedSize !== expectedSize ||
    slot.state !== 'ready' || (expectedSize !== null && slot.size !== expectedSize) ||
    !(slot.blob instanceof Blob)) {
  done({error: 'prefetch não inicializado'});
} else {
  slot.blob.slice(offset, offset + length).arrayBuffer()
    .then(buffer => {
      const serializationStarted = performance.now();
      const bytes = new Uint8Array(buffer);
      let binary = '';
      for (let index = 0; index < bytes.length; index += 1) {
        binary += String.fromCharCode(bytes[index]);
      }
      const data = btoa(binary);
      done({ok: true, data, offset, length: bytes.length,
            serializationMs: performance.now() - serializationStarted});
    }).catch(error => done({error: String(error)}));
}
"""

_CLEAR_PREFETCH_SCRIPT = r"""
const done = arguments[arguments.length - 1];
const token = arguments[0];
// Compatibility marker: window.__auditPrefetch = null
const slots = window.__auditPrefetchSlots || {};
const tokens = token ? [token] : Object.keys(slots);
for (const key of tokens) {
  const slot = slots[key];
  if (slot && slot.controller) slot.controller.abort();
  if (slot) { slot.blob = null; slot.controller = null; slot.error = null; delete slots[key]; }
}
done({ok: true});
"""
_DOWNLOAD_CHUNK_SIZE = 512 * 1024
_VERSION_PAGE_SIZE = 1000
_VERSION_PAGE_RETRIES = 3
FETCH_OPERATION_TIMEOUT_SECONDS = 60
PREFETCH_RETRIES = 1
PREFETCH_BUFFER_SIZE = 2
NORMAL_DOWNLOAD_RETRIES = 1
EDGE_RECOVERY_TIMEOUT_SECONDS = 60


def _odata_value(payload: Mapping[str, Any]) -> object:
    value: object = payload.get("value")
    if value is None and isinstance(payload.get("d"), Mapping):
        value = payload["d"].get("results", payload["d"])
    return value


def _odata_results(payload: Mapping[str, Any]) -> list[Mapping[str, Any]]:
    value = _odata_value(payload)
    if not isinstance(value, list):
        raise SharePointReadError("SharePoint REST retornou coleção inválida")
    return [item for item in value if isinstance(item, Mapping)]


def _odata_object(payload: Mapping[str, Any]) -> Mapping[str, Any]:
    # With ``odata=nometadata`` SharePoint returns a single entity as the
    # top-level JSON object.  Collections still use ``value`` and older OData
    # modes can use ``d``.  Keep all three documented wire shapes separate so
    # a direct entity is not mistaken for a malformed collection response.
    if "value" in payload:
        value = payload["value"]
    elif isinstance(payload.get("d"), Mapping):
        value = payload["d"]
    else:
        value = payload
    if not isinstance(value, Mapping):
        raise SharePointReadError("SharePoint REST retornou objeto inválido")
    return value


def _escape_odata_path(path: str) -> str:
    return quote(path.replace("'", "''"), safe="/'()$=,:?&")


def _normalize_scope_path(scope_path: str, site_path: str) -> str:
    """Converte um caminho relativo ao site em server-relative URL."""
    scope = "/" + unquote(scope_path).strip("/")
    site = "/" + unquote(site_path).strip("/") if site_path.strip("/") else ""
    if site and scope.casefold() != site.casefold() and not scope.casefold().startswith(
        site.casefold() + "/"
    ):
        return f"{site}/{scope.lstrip('/')}"
    return scope


class BrowserSharePointSource:
    """Descobre, enumera e baixa XLSX por GET read-only no próprio Edge."""

    REST_CONTEXT = "sharepoint-rest"

    @property
    def prefetch_buffer_size(self) -> int:
        """Quantidade de versões futuras admitidas pelo buffer do Edge."""
        return PREFETCH_BUFFER_SIZE

    def __init__(
        self,
        site_url: str,
        scope_paths: Sequence[str],
        browser: BrowserSession,
        *,
        temp_directory: str | Path = "data/temp",
        owns_browser: bool = False,
        status_callback: Callable[[str], None] | None = None,
    ) -> None:
        parsed = urlsplit(site_url)
        if parsed.scheme != "https" or not parsed.netloc:
            raise ValueError("A URL do site SharePoint deve usar HTTPS")
        self.site_url = site_url.rstrip("/")
        self._origin = f"{parsed.scheme}://{parsed.netloc}"
        self._site_path = parsed.path
        scopes = tuple(
            dict.fromkeys(
                _normalize_scope_path(path, self._site_path)
                for path in scope_paths
                if path.strip("/")
            )
        )
        if not scopes:
            raise ValueError("Ao menos um escopo SharePoint deve ser configurado")
        self.scope_paths = scopes
        self._browser = browser
        self._owns_browser = owns_browser
        self._workspace = TemporaryWorkspace(temp_directory)
        self._prefetch_slots: dict[str, dict[str, object]] = {}
        self._status_callback = status_callback
        self._incremental_digests: dict[Path, str] = {}

    @classmethod
    def open_edge(
        cls,
        site_url: str,
        scope_paths: Sequence[str],
        *,
        temp_directory: str | Path = "data/temp",
    ) -> "BrowserSharePointSource":
        """Abre Edge visível; login/MFA continuam inteiramente sob controle do usuário."""
        from selenium import webdriver

        root = Path(temp_directory)
        root.mkdir(parents=True, exist_ok=True)
        options = webdriver.EdgeOptions()
        logger.info(
            "Abrindo Edge para autenticação manual SharePoint site=%s escopos=%d modo=read-only",
            site_url,
            len(scope_paths),
        )
        browser = webdriver.Edge(options=options)
        # Há dois timeouts distintos no Selenium:
        # 1) o timeout do JavaScript assíncrono executado no Edge;
        # 2) o timeout HTTP usado pelo Python para aguardar o WebDriver local.
        # A enumeração de dezenas de milhares de versões pode ultrapassar os
        # 120 s padrão do segundo limite, mesmo com script_timeout=600.
        browser.set_script_timeout(600)
        # Um reload de recuperação também precisa ser limitado. Ele ocorre na
        # mesma thread e no mesmo WebDriver, sem criar acesso concorrente.
        browser.set_page_load_timeout(EDGE_RECOVERY_TIMEOUT_SECONDS)
        command_executor = getattr(browser, "command_executor", None)
        if command_executor is not None:
            set_timeout = getattr(command_executor, "set_timeout", None)
            if callable(set_timeout):
                set_timeout(600)
            else:
                client_config = getattr(command_executor, "_client_config", None)
                if client_config is not None:
                    client_config.timeout = 600
        logger.info("Timeouts Selenium configurados script=600s webdriver_http=600s")
        try:
            source = cls(
                site_url,
                scope_paths,
                browser,
                temp_directory=root,
                owns_browser=True,
            )
        except Exception:
            browser.quit()
            raise
        browser.get(site_url)
        logger.info(
            "Edge aberto; aguardando autenticação manual do usuário site=%s",
            site_url,
        )
        return source

    def wait_until_authenticated(self, timeout: float = 600) -> None:
        """Espera deterministicamente a sessão alcançar e ler o site configurado."""
        from selenium.common.exceptions import TimeoutException
        from selenium.webdriver.support.ui import WebDriverWait

        logger.info("Aguardando autenticação manual SharePoint site=%s", self.site_url)

        def authenticated(_browser: BrowserSession) -> bool:
            current = urlsplit(_browser.current_url)
            site = urlsplit(self.site_url)
            if current.scheme != site.scheme or current.netloc != site.netloc:
                return False
            try:
                payload = self._json("web?$select=Id")
                entity = _odata_object(payload)
                return isinstance(entity.get("Id"), str)
            except SharePointReadError:
                return False

        try:
            WebDriverWait(self._browser, timeout, poll_frequency=0.5).until(
                authenticated
            )
        except TimeoutException as error:
            raise SharePointReadError(
                "Tempo esgotado aguardando autenticação manual no SharePoint"
            ) from error
        logger.info("Autenticação SharePoint detectada e site validado")

    def close(self) -> None:
        """Fecha a sessão sem transformar arquivo temporário bloqueado em falha fatal."""
        try:
            self.cancel_prefetch()
        except Exception:
            logger.debug("Falha ao cancelar prefetch durante encerramento", exc_info=True)

        if self._owns_browser:
            try:
                self._browser.quit()
            except Exception:
                logger.debug("Falha ao encerrar Edge durante encerramento", exc_info=True)

        # No Windows, antivírus/indexador ou uma thread que ainda esteja terminando
        # pode manter o XLSX aberto por alguns instantes. Fazemos tentativas curtas
        # e, se continuar bloqueado, deixamos a pasta temporária para a próxima
        # limpeza em vez de derrubar a aplicação com WinError 32.
        last_error: Exception | None = None
        for attempt in range(5):
            try:
                self._workspace.close()
                return
            except PermissionError as error:
                last_error = error
                time.sleep(0.2 * (attempt + 1))
            except FileNotFoundError:
                return
        if last_error is not None:
            logger.warning(
                "Não foi possível remover imediatamente o diretório temporário; "
                "ele permanecerá para limpeza posterior erro=%s",
                last_error,
            )

    def __enter__(self) -> "BrowserSharePointSource":
        return self

    def __exit__(self, *_: object) -> None:
        self.close()

    def _endpoint(self, relative: str) -> str:
        url = f"{self.site_url}/_api/{relative.lstrip('/')}"
        parsed = urlsplit(url)
        if (
            f"{parsed.scheme}://{parsed.netloc}" != self._origin
            or "/_api/" not in parsed.path
        ):
            raise SharePointReadError("Endpoint fora do site ou da API REST permitida")
        return url

    def _json(self, relative: str) -> Mapping[str, Any]:
        return self._json_url(self._endpoint(relative))

    def _json_url(self, url: str) -> Mapping[str, Any]:
        """Lê um endpoint REST absoluto, sempre limitado ao mesmo site SharePoint."""
        parsed = urlsplit(url)
        if (
            f"{parsed.scheme}://{parsed.netloc}" != self._origin
            or "/_api/" not in parsed.path
        ):
            raise SharePointReadError(
                "Somente endpoints GET same-origin da API REST são permitidos"
            )
        result = self._browser.execute_async_script(_FETCH_JSON_SCRIPT, url)
        if not isinstance(result, Mapping) or result.get("error"):
            detail = (
                result.get("error")
                if isinstance(result, Mapping)
                else "resposta inválida"
            )
            raise SharePointReadError(f"Falha de leitura SharePoint REST: {detail}")
        payload = result.get("json")
        if not isinstance(payload, Mapping):
            raise SharePointReadError("SharePoint REST retornou JSON inválido")
        return payload

    def _json_url_with_retry(
        self, url: str, *, page: int
    ) -> tuple[Mapping[str, Any], int, float]:
        """Repete somente a página que falhou, sem reiniciar toda a enumeração."""
        last_error: Exception | None = None
        for attempt in range(1, _VERSION_PAGE_RETRIES + 1):
            try:
                started = time.perf_counter()
                payload = self._json_url(url)
                return payload, attempt - 1, time.perf_counter() - started
            except Exception as error:
                last_error = error
                if attempt >= _VERSION_PAGE_RETRIES:
                    break
                delay = float(2 ** (attempt - 1))
                logger.warning(
                    "Falha temporária ao enumerar versões pagina=%d tentativa=%d/%d "
                    "aguardando=%.0fs erro=%s",
                    page,
                    attempt,
                    _VERSION_PAGE_RETRIES,
                    delay,
                    error,
                )
                time.sleep(delay)
        assert last_error is not None
        raise last_error

    @staticmethod
    def _next_page_url(payload: Mapping[str, Any]) -> str | None:
        next_link = (
            payload.get("@odata.nextLink")
            or payload.get("odata.nextLink")
            or payload.get("__next")
        )
        if next_link is None and isinstance(payload.get("d"), Mapping):
            next_link = payload["d"].get("__next")
        return next_link if isinstance(next_link, str) and next_link else None

    def list_spreadsheets(self) -> tuple[SpreadsheetInfo, ...]:
        logger.info(
            "Descoberta de planilhas iniciada site=%s escopos=%d modo=read-only",
            self.site_url,
            len(self.scope_paths),
        )
        found: dict[str, SpreadsheetInfo] = {}
        pending = list(self.scope_paths)
        visited: set[str] = set()
        while pending:
            folder = pending.pop(0)
            if folder in visited:
                continue
            visited.add(folder)
            encoded = _escape_odata_path(folder)
            files = _odata_results(
                self._json(
                    f"web/GetFolderByServerRelativeUrl('{encoded}')/Files"
                    "?$select=Name,ServerRelativeUrl,UniqueId,Length"
                )
            )
            for item in files:
                name, path, unique_id = (
                    item.get("Name"),
                    item.get("ServerRelativeUrl"),
                    item.get("UniqueId"),
                )
                if (
                    not isinstance(name, str)
                    or not isinstance(path, str)
                    or not isinstance(unique_id, str)
                ):
                    raise SharePointReadError(
                        "Arquivo retornado sem identidade REST comprovável"
                    )
                if name.lower().endswith(".xlsx"):
                    found[unique_id.lower()] = SpreadsheetInfo(
                        site_id=self.site_url,
                        drive_id=self.REST_CONTEXT,
                        drive_item_id=unique_id,
                        name=name,
                        path=path,
                        folder=path.rsplit("/", 1)[0],
                    )
            folders = _odata_results(
                self._json(
                    f"web/GetFolderByServerRelativeUrl('{encoded}')/Folders?$select=Name,ServerRelativeUrl"
                )
            )
            pending.extend(
                path
                for item in folders
                if isinstance((path := item.get("ServerRelativeUrl")), str)
                and item.get("Name") != "Forms"
            )
        result = tuple(sorted(found.values(), key=lambda item: item.path or ""))
        logger.info(
            "Descoberta de planilhas concluída site=%s pastas=%d planilhas=%d",
            self.site_url,
            len(visited),
            len(result),
        )
        return result

    def list_folders(self) -> tuple[tuple[str, str], ...]:
        """Lista recursivamente as subpastas dos escopos para seleção na interface."""
        found: dict[str, str] = {}
        pending = list(self.scope_paths)
        visited: set[str] = set()
        while pending:
            folder = pending.pop(0)
            if folder in visited:
                continue
            visited.add(folder)
            encoded = _escape_odata_path(folder)
            payload = self._json(
                f"web/GetFolderByServerRelativeUrl('{encoded}')/Folders"
                "?$select=Name,ServerRelativeUrl"
            )
            for item in _odata_results(payload):
                name, path = item.get("Name"), item.get("ServerRelativeUrl")
                if (
                    isinstance(name, str)
                    and isinstance(path, str)
                    and name != "Forms"
                ):
                    found[path] = name
                    pending.append(path)
        return tuple(
            (name, path)
            for path, name in sorted(found.items(), key=lambda item: item[1].casefold())
        )

    def set_scope_paths(self, scope_paths: Sequence[str]) -> None:
        """Atualiza os escopos de leitura sem recriar a sessão autenticada."""
        scopes = tuple(
            dict.fromkeys(
                _normalize_scope_path(path, self._site_path)
                for path in scope_paths
                if path.strip("/")
            )
        )
        if not scopes:
            raise ValueError("Ao menos um escopo SharePoint deve ser configurado")
        self.scope_paths = scopes

    def _file_metadata(self, spreadsheet: SpreadsheetInfo) -> Mapping[str, Any]:
        encoded = _escape_odata_path(spreadsheet.path or "")
        return _odata_object(
            self._json(
                f"web/GetFileByServerRelativeUrl('{encoded}')"
                "?$select=Name,ServerRelativeUrl,UniqueId,UIVersion,UIVersionLabel,"
                "TimeLastModified,Length,ModifiedBy/Title,ModifiedBy/Email,ModifiedBy/LoginName"
                "&$expand=ModifiedBy"
            )
        )

    def get_current_version(self, spreadsheet: SpreadsheetInfo) -> VersionInfo:
        """Obtém somente o watermark atual, sem acessar ``File/Versions``."""
        self._validate_spreadsheet(spreadsheet)
        metadata = self._file_metadata(spreadsheet)
        unique_id = metadata.get("UniqueId")
        if not isinstance(unique_id, str) or unique_id.lower() != spreadsheet.drive_item_id.lower():
            raise SharePointReadError("UniqueId atual diverge da identidade da planilha")
        version_id = metadata.get("UIVersion")
        label = metadata.get("UIVersionLabel")
        if not isinstance(version_id, int) or not isinstance(label, str):
            raise SharePointReadError("Arquivo atual sem UIVersion/UIVersionLabel")
        author = metadata.get("ModifiedBy")
        return VersionInfo(
            id=str(version_id), number=label,
            modified_at=metadata.get("TimeLastModified") if isinstance(metadata.get("TimeLastModified"), str) else None,
            author=author.get("Title") if isinstance(author, Mapping) and isinstance(author.get("Title"), str) else None,
            author_email=author.get("Email") if isinstance(author, Mapping) and isinstance(author.get("Email"), str) else None,
            author_login=author.get("LoginName") if isinstance(author, Mapping) and isinstance(author.get("LoginName"), str) else None,
            size=self._size(metadata), source_url=spreadsheet.path, is_current=True,
        )

    def list_versions_delta(
        self,
        spreadsheet: SpreadsheetInfo,
        anchor_id: str,
        anchor_label: str,
        progress_callback: Callable[[int], None] | None = None,
    ) -> tuple[VersionInfo, ...]:
        """Enumera uma cauda inclusiva estrita, ancorada no catálogo local."""
        try:
            technical_id = int(anchor_id)
        except (TypeError, ValueError) as error:
            raise SharePointReadError("anchor técnico inválido") from error
        # Não usa o fallback interno de list_versions: o catálogo decide de
        # forma explícita se uma inconsistência justifica reconstrução completa.
        return self._list_versions(
            spreadsheet, progress_callback,
            checkpoint=(technical_id, anchor_label),
        )

    def list_versions(
        self,
        spreadsheet: SpreadsheetInfo,
        progress_callback: Callable[[int], None] | None = None,
        *,
        checkpoint_id: str | None = None,
        checkpoint_label: str | None = None,
    ) -> tuple[VersionInfo, ...]:
        """Enumera versões, preferindo uma consulta incremental comprovável.

        ``checkpoint_id`` é a única chave usada na consulta. O label serve
        exclusivamente para validar a fronteira devolvida pelo servidor. Caso
        o tenant rejeite/ignore ``$filter`` ou ``$orderby``, ou a fronteira não
        possa ser provada, a enumeração completa é repetida automaticamente.
        """
        if checkpoint_id is None:
            return self._list_versions(spreadsheet, progress_callback)
        if checkpoint_label is None:
            logger.info(
                "PERF enumeracao_fallback planilha=%s motivo=checkpoint_sem_label",
                spreadsheet.name,
            )
            return self._list_versions(spreadsheet, progress_callback)
        try:
            technical_id = int(checkpoint_id)
        except (TypeError, ValueError):
            logger.info(
                "PERF enumeracao_fallback planilha=%s motivo=checkpoint_id_invalido",
                spreadsheet.name,
            )
            return self._list_versions(spreadsheet, progress_callback)

        try:
            return self._list_versions(
                spreadsheet,
                progress_callback,
                checkpoint=(technical_id, checkpoint_label),
            )
        except SharePointReadError as error:
            # Mudança do watermark nunca deve ser mascarada por uma segunda
            # leitura: a enumeração observada já é potencialmente incompleta.
            if "mudou durante a paginação" in str(error):
                raise
            logger.warning(
                "PERF enumeracao_fallback planilha=%s checkpoint_id=%s motivo=%s",
                spreadsheet.name,
                checkpoint_id,
                str(error).replace("\n", " "),
            )
            logger.warning(
                "ENUMERACAO modo=completa_fallback planilha=%s motivo=%s",
                spreadsheet.name,
                str(error).replace("\n", " "),
            )
            if progress_callback is not None:
                progress_callback(0)
            complete = self._list_versions(
                spreadsheet, progress_callback, fallback=True
            )
            boundary = next(
                (version for version in complete if version.id == str(technical_id)),
                None,
            )
            if boundary is None:
                raise SharePointReadError(
                    f"checkpoint técnico ID {technical_id} ausente também na "
                    "enumeração completa"
                )
            if boundary.number != checkpoint_label:
                raise SharePointReadError(
                    "checkpoint local diverge do histórico completo: "
                    f"ID={technical_id}, salvo={checkpoint_label!r}, "
                    f"retornado={boundary.number!r}"
                )
            return complete

    def _list_versions(
        self,
        spreadsheet: SpreadsheetInfo,
        progress_callback: Callable[[int], None] | None = None,
        *,
        checkpoint: tuple[int, str] | None = None,
        fallback: bool = False,
    ) -> tuple[VersionInfo, ...]:
        """Enumera todo o histórico em páginas, sem aceitar truncamento silencioso.

        O endpoint ``File/Versions`` do SharePoint Online nem sempre devolve um
        ``@odata.nextLink`` quando ``$top`` é usado. Por isso, seguimos o
        ``nextLink`` quando ele existir e, se uma página vier cheia sem link de
        continuação, avançamos ordinalmente com ``$skip``. Se o servidor ignorar
        o ``$skip`` e repetir a mesma página, a operação é abortada em vez de
        entregar uma lista incompleta à auditoria.
        """
        self._validate_spreadsheet(spreadsheet)
        mode = "incremental" if checkpoint is not None else (
            "completa_fallback" if fallback else "completa"
        )
        logger.info(
            "PERF enumeracao_inicio planilha=%s modo=%s checkpoint_id=%s checkpoint_label=%s",
            spreadsheet.name,
            mode,
            checkpoint[0] if checkpoint else None,
            checkpoint[1] if checkpoint else None,
        )
        # Um watermark antes/depois impede aceitar silenciosamente uma lista
        # montada enquanto uma nova versão era publicada no SharePoint.
        initial_metadata = self._file_metadata(spreadsheet)
        initial_watermark = (
            initial_metadata.get("UniqueId"),
            initial_metadata.get("UIVersion"),
            initial_metadata.get("UIVersionLabel"),
        )
        encoded = _escape_odata_path(spreadsheet.path or "")
        base_relative = (
            f"web/GetFileByServerRelativeUrl('{encoded}')/Versions"
            "?$expand=CreatedBy&$select=ID,VersionLabel,Created,CreatedBy/Title,"
            "CreatedBy/Email,CreatedBy/LoginName,CheckInComment,Size,Length,Url,IsCurrentVersion"
        )
        if checkpoint is not None:
            # ID is numeric in File/Versions. Explicit ordering plus an
            # inclusive boundary makes the checkpoint snapshot available for
            # the first pending comparison; gaps in IDs are intentionally fine.
            base_relative += f"&$filter=ID ge {checkpoint[0]}&$orderby=ID asc"

        next_url: str | None = self._endpoint(
            base_relative + f"&$top={_VERSION_PAGE_SIZE}&$skip=0"
        )
        historical: list[VersionInfo] = []
        seen_ids: set[str] = set()
        labels_to_ids: dict[str, str] = {}
        visited_page_urls: set[str] = set()
        ordering_direction = 0
        last_version_id: int | None = None
        page = 0
        skip = 0
        started = time.perf_counter()

        while next_url:
            if next_url in visited_page_urls:
                raise SharePointReadError(
                    "SharePoint repetiu o nextLink; a enumeração pode estar incompleta."
                )
            visited_page_urls.add(next_url)
            page += 1
            payload, retries, webdriver_seconds = self._json_url_with_retry(
                next_url, page=page
            )
            parse_started = time.perf_counter()
            items = _odata_results(payload)
            # This is deliberately an approximation of the browser-to-Python JSON
            # payload, useful for comparing $top candidates without changing its
            # fields or wire representation.
            json_bytes = len(
                json.dumps(payload, ensure_ascii=False, separators=(",", ":")).encode(
                    "utf-8"
                )
            )
            parse_seconds = time.perf_counter() - parse_started

            if not items and self._next_page_url(payload) is not None:
                raise SharePointReadError(
                    "SharePoint retornou página vazia com continuação; "
                    "a enumeração pode estar incompleta."
                )

            for item in items:
                version_id, label = item.get("ID"), item.get("VersionLabel")
                if not isinstance(version_id, int) or not isinstance(label, str):
                    raise SharePointReadError(
                        "Versão REST sem ID técnico ou VersionLabel"
                    )
                version_key = str(version_id)
                if version_key in seen_ids:
                    raise SharePointReadError(
                        f"Versão duplicada na paginação: ID {version_key}."
                    )
                known_id = labels_to_ids.get(label)
                if known_id is not None and known_id != version_key:
                    raise SharePointReadError(
                        "VersionLabel associado a IDs diferentes na paginação: "
                        f"{label}."
                    )

                if last_version_id is not None:
                    step = 1 if version_id > last_version_id else -1
                    if ordering_direction == 0:
                        ordering_direction = step
                    elif step != ordering_direction:
                        raise SharePointReadError(
                            "Página de versões fora de ordem monotônica; "
                            "a enumeração foi rejeitada."
                        )
                last_version_id = version_id
                seen_ids.add(version_key)
                labels_to_ids[label] = version_key
                historical.append(self._historical_version(item, version_id, label))

            if progress_callback is not None:
                try:
                    progress_callback(len(historical))
                except Exception:
                    logger.warning(
                        "Falha ao publicar progresso da enumeração de versões",
                        exc_info=True,
                    )

            next_link = self._next_page_url(payload)
            logger.info(
                "PERF enumeracao_pagina planilha=%s pagina=%d itens=%d "
                "primeiro_id=%s ultimo_id=%s webdriver=%.3fs json_bytes_aprox=%d "
                "parse_python=%.3fs retries=%d nextlink=%s acumulado=%d skip=%d",
                spreadsheet.name,
                page,
                len(items),
                items[0].get("ID") if items else None,
                items[-1].get("ID") if items else None,
                webdriver_seconds,
                json_bytes,
                parse_seconds,
                retries,
                next_link or "-",
                len(historical),
                skip,
            )
            logger.info(
                "PERF enumeracao_progresso planilha=%s modo=%s encontradas=%d paginas=%d",
                spreadsheet.name, mode, len(historical), page,
            )

            if next_link is not None:
                next_url = urljoin(self.site_url + "/", next_link)
                # Mantemos o skip apenas para log/fallback caso o próximo payload
                # deixe de fornecer continuação.
                skip += len(items)
            elif len(items) >= _VERSION_PAGE_SIZE:
                # Alguns tenants/coleções de versões aplicam $top, mas não expõem
                # nextLink. Nesse caso, avançamos com $skip.
                skip += len(items)
                next_url = self._endpoint(
                    base_relative
                    + f"&$top={_VERSION_PAGE_SIZE}&$skip={skip}"
                )
            else:
                next_url = None

        if checkpoint is not None and ordering_direction < 0:
            raise SharePointReadError(
                "servidor não respeitou a ordenação incremental ascendente"
            )
        historical.sort(key=lambda version: int(version.id))
        metadata = self._file_metadata(spreadsheet)
        final_watermark = (
            metadata.get("UniqueId"),
            metadata.get("UIVersion"),
            metadata.get("UIVersionLabel"),
        )
        if initial_watermark != final_watermark:
            raise SharePointReadError(
                "A versão atual mudou durante a paginação; a enumeração "
                "foi rejeitada como potencialmente incompleta."
            )
        if checkpoint is not None:
            checkpoint_id, checkpoint_label = checkpoint
            boundary = next(
                (version for version in historical if int(version.id) == checkpoint_id),
                None,
            )
            if boundary is None:
                raise SharePointReadError(
                    f"checkpoint técnico ID {checkpoint_id} ausente no resultado incremental"
                )
            if boundary.number != checkpoint_label:
                raise SharePointReadError(
                    "checkpoint com conflito ID/VersionLabel: "
                    f"ID={checkpoint_id}, salvo={checkpoint_label!r}, "
                    f"retornado={boundary.number!r}"
                )
            if any(int(version.id) < checkpoint_id for version in historical):
                raise SharePointReadError(
                    "servidor ignorou a fronteira incremental solicitada"
                )
        unique_id = metadata.get("UniqueId")
        if (
            not isinstance(unique_id, str)
            or unique_id.lower() != spreadsheet.drive_item_id.lower()
        ):
            raise SharePointReadError(
                "UniqueId atual diverge da identidade da planilha"
            )
        ui_version, current_label = (
            metadata.get("UIVersion"),
            metadata.get("UIVersionLabel"),
        )
        if not isinstance(ui_version, int) or not isinstance(current_label, str):
            raise SharePointReadError("Arquivo atual sem UIVersion/UIVersionLabel")
        modified_by = metadata.get("ModifiedBy")
        current = VersionInfo(
            id=str(ui_version),
            number=current_label,
            modified_at=metadata.get("TimeLastModified")
            if isinstance(metadata.get("TimeLastModified"), str)
            else None,
            author=modified_by.get("Title")
            if isinstance(modified_by, Mapping)
            and isinstance(modified_by.get("Title"), str)
            else None,
            author_email=modified_by.get("Email")
            if isinstance(modified_by, Mapping)
            and isinstance(modified_by.get("Email"), str)
            else None,
            author_login=modified_by.get("LoginName")
            if isinstance(modified_by, Mapping)
            and isinstance(modified_by.get("LoginName"), str)
            else None,
            size=self._size(metadata),
            source_url=spreadsheet.path,
            is_current=True,
        )
        current_by_id = next((v for v in historical if v.id == current.id), None)
        current_by_label = next(
            (v for v in historical if v.number == current.number), None
        )
        if current_by_id is not None and current_by_id.number != current.number:
            raise SharePointReadError(
                "ID dos metadados atuais associado a VersionLabel diferente no "
                f"histórico: ID={current.id!r}, label_historico="
                f"{current_by_id.number!r}, label_atual={current.number!r}."
            )
        if current_by_label is not None and current_by_label.id != current.id:
            raise SharePointReadError(
                "VersionLabel dos metadados atuais associado a ID diferente no "
                f"histórico: label={current.number!r}, id_historico="
                f"{current_by_label.id!r}, id_atual={current.id!r}."
            )
        marked_current = [version for version in historical if version.is_current]
        for marked in marked_current:
            if marked.id != current.id or marked.number != current.number:
                logger.warning(
                    "IsCurrentVersion histórico diverge dos metadados atuais "
                    "planilha=%s id_historico=%s label_historico=%s "
                    "data_historico=%s id_atual=%s label_atual=%s data_atual=%s",
                    spreadsheet.name,
                    marked.id,
                    marked.number,
                    marked.modified_at,
                    current.id,
                    current.number,
                    current.modified_at,
                )
        if len(marked_current) > 1:
            logger.warning(
                "Múltiplos IsCurrentVersion no histórico planilha=%s quantidade=%d "
                "marcadores=%s atual_autoritativa=(id=%s,label=%s,data=%s)",
                spreadsheet.name,
                len(marked_current),
                [(v.id, v.number, v.modified_at) for v in marked_current],
                current.id,
                current.number,
                current.modified_at,
            )

        # Quando o endpoint histórico inclui a atual, substituímos somente a
        # duplicata exata pelos metadados atuais mais completos. IsCurrentVersion
        # histórico é apenas diagnóstico; não redefine a versão atual.
        combined = [
            replace(version, is_current=False)
            for version in historical
            if version is not current_by_id
        ]
        combined.append(current)
        combined_ids = [version.id for version in combined]
        combined_labels = [version.number for version in combined]
        if len(combined_ids) != len(set(combined_ids)):
            raise SharePointReadError("ID duplicado após reconciliar a versão atual.")
        if len(combined_labels) != len(set(combined_labels)):
            raise SharePointReadError(
                "VersionLabel conflitante após reconciliar a versão atual."
            )
        historical_ids = [int(version.id) for version in combined[:-1]]
        if historical_ids != sorted(historical_ids):
            raise SharePointReadError(
                "Ordem histórica inválida após reconciliar a versão atual."
            )
        logger.info(
            "ENUMERACAO modo=%s planilha=%s checkpoint_id=%s",
            mode, spreadsheet.name, checkpoint[0] if checkpoint else None,
        )
        logger.info(
            "PERF enumeracao_versoes planilha=%s modo=%s checkpoint_id=%s "
            "historicas_recebidas=%d paginas=%d total_segundos=%.3f",
            spreadsheet.name,
            mode,
            checkpoint[0] if checkpoint else None,
            len(historical),
            page,
            time.perf_counter() - started,
        )
        if [version for version in combined if version.is_current] != [current]:
            raise SharePointReadError(
                "Reconciliação não produziu exatamente a versão atual autoritativa."
            )
        logger.info(
            "Versões enumeradas planilha=%s identidade=%s total=%d modo=read-only",
            spreadsheet.name,
            spreadsheet.drive_item_id,
            len(combined),
        )
        return tuple(combined)

    @staticmethod
    def _historical_version(
        item: Mapping[str, Any], version_id: int, label: str
    ) -> VersionInfo:
        author = item.get("CreatedBy")
        return VersionInfo(
            id=str(version_id),
            number=label,
            modified_at=item.get("Created")
            if isinstance(item.get("Created"), str)
            else None,
            author=author.get("Title")
            if isinstance(author, Mapping) and isinstance(author.get("Title"), str)
            else None,
            author_email=author.get("Email")
            if isinstance(author, Mapping) and isinstance(author.get("Email"), str)
            else None,
            author_login=author.get("LoginName")
            if isinstance(author, Mapping) and isinstance(author.get("LoginName"), str)
            else None,
            comment=item.get("CheckInComment")
            if isinstance(item.get("CheckInComment"), str)
            else None,
            size=BrowserSharePointSource._size(item),
            source_url=item.get("Url") if isinstance(item.get("Url"), str) else None,
            is_current=bool(item.get("IsCurrentVersion", False)),
        )

    @staticmethod
    def _size(item: Mapping[str, Any]) -> int | None:
        value = item.get("Size", item.get("Length"))
        return (
            int(value)
            if isinstance(value, (int, str)) and str(value).isdigit()
            else None
        )

    def _version_download_url(
        self, spreadsheet: SpreadsheetInfo, version: VersionInfo
    ) -> str:
        self._validate_spreadsheet(spreadsheet)
        if not re.fullmatch(r"\d+", version.id):
            raise SharePointReadError("ID de versão REST inválido")
        encoded = _escape_odata_path(spreadsheet.path or "")
        relative = f"web/GetFileByServerRelativeUrl('{encoded}')"
        if not version.is_current:
            relative += f"/Versions({version.id})"
        return self._endpoint(relative + "/$value")

    def _historical_url_fallback(self, version: VersionInfo) -> str | None:
        """Normaliza o ``FileVersion.Url`` sem aceitar saída do site autenticado.

        Alguns tenants enumeram corretamente uma versão, mas deixam a rota
        ``Versions(id)/$value`` aguardando indefinidamente. ``FileVersion.Url``
        identifica o mesmo binário histórico e é devolvido junto do ID técnico;
        ele é usado apenas como segunda rota, nunca para localizar o checkpoint.
        """
        if version.is_current or not version.source_url:
            return None
        candidate = urljoin(self.site_url.rstrip("/") + "/", version.source_url)
        parsed = urlsplit(candidate)
        if f"{parsed.scheme}://{parsed.netloc}" != self._origin:
            logger.warning(
                "URL histórica alternativa rejeitada por origem diferente versao=%s",
                version.number,
            )
            return None
        site_prefix = self._site_path.rstrip("/") + "/"
        if self._site_path and not (
            parsed.path.casefold().startswith(site_prefix.casefold())
            or parsed.path.casefold() == self._site_path.casefold()
        ):
            logger.warning(
                "URL histórica alternativa rejeitada fora do site versao=%s path=%s",
                version.number,
                parsed.path,
            )
            return None
        return candidate

    def prefetch_version(
        self, spreadsheet: SpreadsheetInfo, version: VersionInfo
    ) -> bool:
        """Inicia o download da próxima versão no Edge sem bloquear o Python.

        Até dois prefetches ficam ativos. Os ``fetches`` continuam no próprio Edge
        enquanto o Python calcula hash, lê o XLSX e compara a versão atual.
        """
        url = self._version_download_url(spreadsheet, version)
        existing = next(
            (slot for slot in self._prefetch_slots.values()
             if slot["url"] == url and slot["version_id"] == version.id),
            None,
        )
        if existing is not None:
            return True
        if len(self._prefetch_slots) >= PREFETCH_BUFFER_SIZE:
            self._log_prefetch_buffer()
            return False
        started = time.perf_counter()
        token = f"{version.id}:{time.monotonic_ns()}"
        result = self._browser.execute_async_script(
            _START_PREFETCH_SCRIPT,
            url,
            int(FETCH_OPERATION_TIMEOUT_SECONDS * 1000),
            token,
            version.number,
            version.id,
            version.size,
        )
        if not isinstance(result, Mapping) or result.get("error") or not result.get("ok"):
            detail = result.get("error") if isinstance(result, Mapping) else "resposta inválida"
            logger.warning(
                "Prefetch não iniciado planilha=%s versao=%s erro=%s",
                spreadsheet.name,
                version.number,
                detail,
            )
            return False
        actual_token = result.get("token", token)
        if not isinstance(actual_token, str):
            return False
        self._prefetch_slots[actual_token] = {
            "url": url, "started_at": started, "version_id": version.id,
            "version": version.number, "size": version.size,
        }
        logger.info(
            "PERF prefetch_slot_inicio planilha=%s versao=%s estado=%s",
            spreadsheet.name,
            version.number,
            result.get("state", "started"),
        )
        self._log_prefetch_buffer()
        return True

    def _log_prefetch_buffer(self) -> None:
        """Publica limites do buffer; bytes exatos são confirmados ao consumir."""
        bytes_total = sum(
            int(slot["size"]) for slot in self._prefetch_slots.values()
            if isinstance(slot.get("size"), int)
        )
        logger.info(
            "PERF prefetch_buffer ocupados=%d ready=%d pending=%d bytes_total=%d "
            "rss_python=%d rss_edge=%d",
            len(self._prefetch_slots), 0, len(self._prefetch_slots), bytes_total,
            self._rss_bytes("self"), self._edge_rss_bytes(),
        )

    @staticmethod
    def _rss_bytes(pid: int | str) -> int:
        """Obtém RSS corrente pelo procfs; retorna -1 fora de ambientes compatíveis."""
        try:
            fields = Path(f"/proc/{pid}/statm").read_text(encoding="ascii").split()
            return int(fields[1]) * int(os.sysconf("SC_PAGE_SIZE"))
        except (OSError, ValueError, IndexError):
            return -1

    def _edge_rss_bytes(self) -> int:
        """Mede a árvore do driver Edge quando o PID é exposto pelo Selenium."""
        service = getattr(self._browser, "service", None)
        process = getattr(service, "process", None)
        root_pid = getattr(process, "pid", None)
        if not isinstance(root_pid, int):
            return -1
        descendants = {root_pid}
        changed = True
        while changed:
            changed = False
            for status in Path("/proc").glob("[0-9]*/status"):
                try:
                    values = status.read_text(encoding="utf-8").splitlines()
                    pid = int(status.parent.name)
                    ppid = int(
                        next(line for line in values if line.startswith("PPid:")).split()[1]
                    )
                except (OSError, ValueError, StopIteration, IndexError):
                    continue
                if ppid in descendants and pid not in descendants:
                    descendants.add(pid)
                    changed = True
        sizes = [self._rss_bytes(pid) for pid in descendants]
        known = [size for size in sizes if size >= 0]
        return sum(known) if known else -1

    def cancel_prefetch(self, token: str | None = None, *, reason: str = "cancelado") -> None:
        """Descarta o blob pré-baixado, sem alterar qualquer arquivo no SharePoint."""
        discarded = list(self._prefetch_slots.items()) if token is None else [
            (token, self._prefetch_slots[token])
        ] if token in self._prefetch_slots else []
        try:
            self._browser.execute_async_script(_CLEAR_PREFETCH_SCRIPT, token or "")
        except Exception:
            logger.debug("Falha ao limpar prefetch", exc_info=True)
        finally:
            for key, slot in discarded:
                self._prefetch_slots.pop(key, None)
                logger.info("PERF prefetch_slot_descartado versao=%s motivo=%s", slot["version"], reason)
            self._log_prefetch_buffer()

    def _notify_status(self, message: str) -> None:
        if self._status_callback is not None:
            try:
                self._status_callback(message)
            except Exception:
                logger.warning("Falha ao publicar estado do fetch", exc_info=True)

    def set_status_callback(self, callback: Callable[[str], None] | None) -> None:
        """Define o canal thread-safe usado para mensagens operacionais da UI."""
        self._status_callback = callback

    def _abort_attempt(
        self, spreadsheet: SpreadsheetInfo, version: VersionInfo, *, use_prefetch: bool,
        prefetch_token: str | None = None,
    ) -> None:
        cleanup_ok = True
        try:
            if use_prefetch:
                self.cancel_prefetch(prefetch_token, reason="falha")
            else:
                self._browser.execute_async_script(_CLEAR_DOWNLOAD_SCRIPT)
        except Exception:
            cleanup_ok = False
            logger.warning("Falha ao limpar tentativa de fetch", exc_info=True)
        logger.warning(
            "PERF fetch_abortado planilha=%s versao=%s limpeza_ok=%s",
            spreadsheet.name, version.number, str(cleanup_ok).lower(),
        )

    def _recover_edge_session(
        self, spreadsheet: SpreadsheetInfo, version: VersionInfo, *, reason: str
    ) -> None:
        """Recarrega e valida a sessão existente, sem substituir o WebDriver.

        A limpeza anterior ao refresh é deliberadamente best-effort: um contexto
        JavaScript travado pode não responder, mas o reload ainda deve ter a
        oportunidade de descartá-lo. A validação REST posterior comprova tanto a
        autenticação quanto o acesso ao site esperado.
        """
        recovery_started = time.perf_counter()
        logger.warning(
            "PERF edge_recovery_inicio planilha=%s versao=%s motivo=%s",
            spreadsheet.name,
            version.number,
            reason,
        )
        self._notify_status("Edge não respondeu. Recuperando sessão...")
        try:
            self.cancel_prefetch()
            # Também elimina um possível slot do download normal. Falhar aqui
            # não impede o refresh, que destrói todo o contexto JavaScript.
            try:
                self._browser.execute_async_script(_CLEAR_DOWNLOAD_SCRIPT)
            except Exception:
                logger.debug("Contexto JS não respondeu antes do refresh", exc_info=True)

            handle = self._browser.current_window_handle
            if not isinstance(handle, str) or not handle:
                raise SharePointReadError("window handle do Edge inválido")

            self._notify_status("Atualizando página do SharePoint...")
            refresh_started = time.perf_counter()
            logger.warning("PERF edge_refresh_inicio versao=%s", version.number)
            self._browser.refresh()
            logger.warning(
                "PERF edge_refresh_concluido versao=%s duracao=%.3fs",
                version.number,
                time.perf_counter() - refresh_started,
            )

            deadline = time.monotonic() + EDGE_RECOVERY_TIMEOUT_SECONDS
            while True:
                ready_state = self._browser.execute_script("return document.readyState")
                if ready_state in {"interactive", "complete"}:
                    break
                if time.monotonic() >= deadline:
                    raise SharePointReadError("timeout aguardando reload do SharePoint")
                time.sleep(0.1)

            refreshed_handle = self._browser.current_window_handle
            if not isinstance(refreshed_handle, str) or not refreshed_handle:
                raise SharePointReadError("window handle do Edge inválido após reload")
            current_url = self._browser.current_url
            current = urlsplit(current_url)
            expected = urlsplit(self.site_url)
            expected_path = expected.path.rstrip("/")
            if (
                f"{current.scheme}://{current.netloc}" != self._origin
                or (
                    expected_path
                    and current.path.rstrip("/") != expected_path
                    and not current.path.startswith(expected_path + "/")
                )
            ):
                raise SharePointReadError(
                    f"Edge saiu do site SharePoint esperado: {current_url}"
                )
            entity = _odata_object(self._json("web?$select=Id"))
            if not isinstance(entity.get("Id"), str) or not entity["Id"]:
                raise SharePointReadError("sessão SharePoint não autenticada após reload")
            logger.warning(
                "PERF edge_session_validada versao=%s url=%s",
                version.number,
                current_url,
            )
            logger.warning(
                "PERF edge_recovery_concluida versao=%s duracao=%.3fs",
                version.number,
                time.perf_counter() - recovery_started,
            )
            self._notify_status(
                f"Sessão recuperada. Tentando novamente a versão {version.number}..."
            )
        except Exception as error:
            self.cancel_prefetch()
            logger.error(
                "PERF edge_recovery_falhou versao=%s erro=%s",
                version.number,
                error,
            )
            message = (
                "Não foi possível recuperar a sessão do Edge. Auditoria interrompida "
                "com checkpoint preservado. Reconecte/autentique novamente."
            )
            self._notify_status(message)
            raise SharePointReadError(f"{message} Motivo: {error}") from error

    def _download_to_destination(
        self,
        spreadsheet: SpreadsheetInfo,
        version: VersionInfo,
        url: str,
        destination: Path,
        *,
        use_prefetch: bool,
        attempt: int,
        prefetch_token: str | None = None,
    ) -> Path:
        total_started = time.perf_counter()
        prefetch_slot = self._prefetch_slots.get(prefetch_token or "") if use_prefetch else None
        prefetch_started_at = prefetch_slot.get("started_at") if prefetch_slot else None
        prefetch_age = (
            time.perf_counter() - prefetch_started_at
            if prefetch_started_at is not None
            else 0.0
        )
        begin_started = time.perf_counter()
        logger.info(
            "PERF fetch_edge_inicio planilha=%s versao=%s url=%s modo=%s",
            spreadsheet.name,
            version.number,
            url,
            "prefetch" if use_prefetch else "normal",
        )
        if use_prefetch:
            if prefetch_slot is None:
                raise SharePointReadError("Slot de prefetch esperado não existe")
            result = self._browser.execute_async_script(
                _WAIT_PREFETCH_SCRIPT, url, prefetch_token, version.number,
                version.id, version.size,
            )
        else:
            result = self._browser.execute_async_script(
                _BEGIN_DOWNLOAD_SCRIPT,
                url,
                int(FETCH_OPERATION_TIMEOUT_SECONDS * 1000),
            )
        begin_seconds = time.perf_counter() - begin_started

        if (
            not isinstance(result, Mapping)
            or result.get("error")
            or not result.get("ok")
        ):
            detail = result.get("error") if isinstance(result, Mapping) else "resposta inválida"
            timed_out = bool(isinstance(result, Mapping) and result.get("timeout"))
            if timed_out:
                logger.warning(
                    "PERF fetch_timeout planilha=%s versao=%s modo=%s tentativa=%d "
                    "timeout=%.3fs prefetch_idade=%.3fs url=%s",
                    spreadsheet.name, version.number,
                    "prefetch" if use_prefetch else "normal", attempt,
                    FETCH_OPERATION_TIMEOUT_SECONDS, prefetch_age, url,
                )
            self._abort_attempt(
                spreadsheet, version, use_prefetch=use_prefetch,
                prefetch_token=prefetch_token,
            )
            reason = "timeout" if timed_out else "erro"
            raise SharePointReadError(
                f"Falha no download SharePoint REST ({reason}): {detail}"
            )

        temporary = destination.with_suffix(destination.suffix + ".part")
        clear_seconds = 0.0
        size = result.get("size")
        if not isinstance(size, int) or size <= 0:
            if use_prefetch:
                self.cancel_prefetch(prefetch_token, reason="tamanho_invalido")
            else:
                self._browser.execute_async_script(_CLEAR_DOWNLOAD_SCRIPT)
            raise SharePointReadError("Download SharePoint vazio ou com tamanho inválido")
        if version.size is not None and size != version.size:
            if use_prefetch:
                self.cancel_prefetch(prefetch_token, reason="tamanho_divergente")
            else:
                self._browser.execute_async_script(_CLEAR_DOWNLOAD_SCRIPT)
            raise SharePointReadError(
                "Tamanho do blob diverge do metadado da versão: "
                f"esperado={version.size} recebido={size}"
            )
        logger.info(
            "PERF fetch_edge_concluido planilha=%s versao=%s url=%s bytes=%d "
            "fetch=%.3fs prefetch_idade=%.3fs prefetch_pronto=%s espera_residual=%.3fs",
            spreadsheet.name,
            version.number,
            url,
            size,
            begin_seconds,
            prefetch_age,
            bool(result.get("wasReady")) if use_prefetch else False,
            float(result.get("waitMs", begin_seconds * 1000.0)) / 1000.0
            if use_prefetch
            and isinstance(result.get("waitMs", begin_seconds * 1000.0), (int, float))
            else 0.0,
        )
        if use_prefetch:
            logger.info(
                "PERF prefetch_slot_ready versao=%s idade=%.3fs bytes=%d",
                version.number,
                prefetch_age,
                size,
            )

        try:
            chunk_calls = 0
            chunk_transfer_seconds = 0.0
            serialization_seconds = 0.0
            decode_seconds = 0.0
            write_seconds = 0.0
            incremental_digest = hashlib.sha256()
            with temporary.open("wb") as output:
                offset = 0
                while offset < size:
                    requested = min(_DOWNLOAD_CHUNK_SIZE, size - offset)
                    chunk_started = time.perf_counter()
                    if use_prefetch:
                        chunk = self._browser.execute_async_script(
                            _READ_PREFETCH_CHUNK_SCRIPT,
                            url,
                            prefetch_token,
                            version.number,
                            version.id,
                            version.size,
                            offset,
                            requested,
                        )
                    else:
                        chunk = self._browser.execute_async_script(
                            _READ_DOWNLOAD_CHUNK_SCRIPT,
                            offset,
                            requested,
                        )
                    chunk_transfer_seconds += time.perf_counter() - chunk_started
                    chunk_calls += 1
                    if (
                        not isinstance(chunk, Mapping)
                        or chunk.get("error")
                        or not isinstance(chunk.get("data"), str)
                        or not isinstance(chunk.get("length"), int)
                        or chunk["length"] <= 0
                    ):
                        detail = chunk.get("error") if isinstance(chunk, Mapping) else "resposta inválida"
                        raise SharePointReadError(
                            f"Falha ao ler download SharePoint: {detail}"
                        )
                    if chunk.get("offset") != offset:
                        raise SharePointReadError(
                            "Bloco SharePoint fora de ordem: "
                            f"esperado={offset} recebido={chunk.get('offset')}"
                        )
                    if chunk["length"] != requested:
                        raise SharePointReadError(
                            "Tamanho do bloco SharePoint incorreto: "
                            f"esperado={requested} recebido={chunk['length']}"
                        )
                    decode_started = time.perf_counter()
                    try:
                        decoded = base64.b64decode(chunk["data"], validate=True)
                    except (binascii.Error, ValueError) as error:
                        raise SharePointReadError(
                            "Bloco SharePoint contém base64 inválido"
                        ) from error
                    if len(decoded) != chunk["length"]:
                        raise SharePointReadError(
                            "Tamanho do bloco baixado diverge do informado"
                        )
                    decode_seconds += time.perf_counter() - decode_started
                    write_started = time.perf_counter()
                    output.write(decoded)
                    incremental_digest.update(decoded)
                    write_seconds += time.perf_counter() - write_started
                    serialization_ms = chunk.get("serializationMs", 0.0)
                    if isinstance(serialization_ms, (int, float)):
                        serialization_seconds += float(serialization_ms) / 1000.0
                    logger.info(
                        "PERF download_chunk planilha=%s versao=%s url=%s numero=%d "
                        "offset=%d bytes=%d webdriver=%.3fs base64_serializacao=%.3fs",
                        spreadsheet.name,
                        version.number,
                        url,
                        chunk_calls,
                        offset,
                        chunk["length"],
                        time.perf_counter() - chunk_started,
                        float(serialization_ms) / 1000.0
                        if isinstance(serialization_ms, (int, float))
                        else 0.0,
                    )
                    offset += chunk["length"]
            if offset != size or temporary.stat().st_size != size:
                raise SharePointReadError(
                    f"Download SharePoint truncado: esperado={size} gravado={offset}"
                )
            rename_started = time.perf_counter()
            temporary.replace(destination)
            rename_seconds = time.perf_counter() - rename_started
            self._incremental_digests[destination] = incremental_digest.hexdigest()
            validate_started = time.perf_counter()
            self._validate_xlsx(destination)
            validate_seconds = time.perf_counter() - validate_started
        except Exception:
            self._incremental_digests.pop(destination, None)
            temporary.unlink(missing_ok=True)
            destination.unlink(missing_ok=True)
            raise
        finally:
            clear_started = time.perf_counter()
            if use_prefetch:
                self._browser.execute_async_script(_CLEAR_PREFETCH_SCRIPT, prefetch_token)
                consumed = self._prefetch_slots.pop(prefetch_token or "", None)
                if consumed is not None:
                    logger.info(
                        "PERF prefetch_slot_consumido versao=%s idade=%.3fs",
                        version.number,
                        time.perf_counter() - float(consumed["started_at"]),
                    )
                self._log_prefetch_buffer()
            else:
                self._browser.execute_async_script(_CLEAR_DOWNLOAD_SCRIPT)
            clear_seconds = time.perf_counter() - clear_started

        total_seconds = time.perf_counter() - total_started
        logger.info(
            "PERF download_sharepoint planilha=%s versao=%s url=%s modo=%s bytes=%d blocos=%d "
            "bloco_bytes=%d espera_fetch=%.3fs transferencia_blocos=%.3fs "
            "base64_serializacao=%.3fs base64_decode=%.3fs escrita=%.3fs rename=%.3fs "
            "validacao_zip=%.3fs limpeza=%.3fs prefetch_idade=%.3fs round_trips=%d total=%.3fs",
            spreadsheet.name,
            version.number,
            url,
            "prefetch" if use_prefetch else "normal",
            size,
            chunk_calls,
            _DOWNLOAD_CHUNK_SIZE,
            begin_seconds,
            chunk_transfer_seconds,
            serialization_seconds,
            decode_seconds,
            write_seconds,
            rename_seconds,
            validate_seconds,
            clear_seconds,
            prefetch_age,
            chunk_calls + 2,
            total_seconds,
        )
        logger.debug(
            "Versão adquirida planilha=%s identidade=%s versao=%s atual=%s modo=%s",
            spreadsheet.name,
            spreadsheet.drive_item_id,
            version.number,
            version.is_current,
            "prefetch" if use_prefetch else "normal",
        )
        return destination

    def get_version(self, spreadsheet: SpreadsheetInfo, version: VersionInfo) -> Path:
        url = self._version_download_url(spreadsheet, version)
        alternate_url = self._historical_url_fallback(version)
        destination = self._workspace.filename(
            spreadsheet.site_id,
            spreadsheet.drive_id,
            spreadsheet.drive_item_id,
            version.id,
        )
        prefetch_token = next(
            (token for token, slot in self._prefetch_slots.items()
             if slot["url"] == url and slot["version_id"] == version.id
             and slot["version"] == version.number and slot["size"] == version.size),
            None,
        )
        use_prefetch = prefetch_token is not None
        if use_prefetch:
            last_prefetch_error: Exception | None = None
            for attempt in range(1, PREFETCH_RETRIES + 2):
                try:
                    return self._download_to_destination(
                        spreadsheet, version, url, destination,
                        use_prefetch=True, attempt=attempt,
                        prefetch_token=prefetch_token,
                    )
                except SharePointReadError as error:
                    last_prefetch_error = error
                    if attempt <= PREFETCH_RETRIES:
                        logger.warning(
                            "PERF fetch_retry planilha=%s versao=%s modo=prefetch "
                            "tentativa=%d motivo=%s",
                            spreadsheet.name, version.number, attempt + 1,
                            "timeout" if "timeout" in str(error) else "erro",
                        )
                        self._notify_status(
                            f"Prefetch excedeu {FETCH_OPERATION_TIMEOUT_SECONDS} s. "
                            f"Retry {attempt}/{PREFETCH_RETRIES}..."
                        )
                        self.cancel_prefetch(prefetch_token, reason="retry")
                        if not self.prefetch_version(spreadsheet, version):
                            break
                        prefetch_token = next(reversed(self._prefetch_slots))
            assert last_prefetch_error is not None
            recovery_reason = (
                "timeout" if "timeout" in str(last_prefetch_error).lower() else "webdriver_error"
            )
            self._recover_edge_session(
                spreadsheet, version, reason=recovery_reason
            )
            # O refresh destruiu o contexto anterior. Um prefetch inteiramente
            # novo mantém URL/VersionInfo, mas recebe token e blob novos.
            if self.prefetch_version(spreadsheet, version):
                prefetch_token = next(reversed(self._prefetch_slots))
                try:
                    return self._download_to_destination(
                        spreadsheet,
                        version,
                        url,
                        destination,
                        use_prefetch=True,
                        attempt=1,
                        prefetch_token=prefetch_token,
                    )
                except SharePointReadError as error:
                    last_prefetch_error = error
            logger.warning(
                    "PERF fetch_fallback planilha=%s versao=%s origem=prefetch "
                    "destino=normal motivo=%s",
                    spreadsheet.name, version.number,
                    "timeout" if "timeout" in str(last_prefetch_error) else "erro",
                )
            self._notify_status(
                f"Fallback para download normal da versão {version.number}..."
            )
            self.cancel_prefetch()

        last_error: Exception | None = None
        normal_urls = (
            tuple(dict.fromkeys((url, alternate_url)))
            if alternate_url
            else (url,)
        )
        for attempt in range(1, NORMAL_DOWNLOAD_RETRIES + 2):
            for route_index, normal_url in enumerate(normal_urls, start=1):
                try:
                    if route_index > 1:
                        logger.warning(
                            "PERF fetch_fallback planilha=%s versao=%s "
                            "origem=versions_value destino=fileversion_url tentativa=%d",
                            spreadsheet.name,
                            version.number,
                            attempt,
                        )
                        self._notify_status(
                            f"Tentando rota histórica alternativa para a versão "
                            f"{version.number}..."
                        )
                    return self._download_to_destination(
                        spreadsheet, version, normal_url, destination,
                        use_prefetch=False, attempt=attempt,
                    )
                except SharePointReadError as error:
                    last_error = error
                    logger.warning(
                        "PERF fetch_retry planilha=%s versao=%s modo=normal "
                        "tentativa=%d rota=%d/%d motivo=%s",
                        spreadsheet.name, version.number, attempt,
                        route_index, len(normal_urls),
                        "timeout" if "timeout" in str(error) else "erro",
                    )
            if attempt <= NORMAL_DOWNLOAD_RETRIES:
                assert last_error is not None
                if len(normal_urls) == 1:
                    self._notify_status(
                        "SharePoint demorando para responder. Tentando novamente..."
                    )
        logger.warning(
            "PERF fetch_falha_final planilha=%s versao=%s tentativas=%d rotas=%d erro=%s",
            spreadsheet.name,
            version.number,
            NORMAL_DOWNLOAD_RETRIES + 1,
            len(normal_urls),
            last_error,
        )
        assert last_error is not None
        self._notify_status(
            f"Falha ao obter a versão {version.number} após tentativas. "
            "Auditoria interrompida com checkpoint preservado."
        )
        raise last_error

    def release_version(self, path: Path) -> None:
        self._incremental_digests.pop(path, None)
        self._workspace.release(path)

    def verify_download_digest(self, path: Path, reference_digest: str) -> None:
        """Compara o SHA incremental da transferência com a leitura de referência."""
        incremental = self._incremental_digests.get(path)
        if incremental is None:
            raise SharePointReadError("SHA-256 incremental do download indisponível")
        if incremental != reference_digest:
            raise SharePointReadError(
                "SHA-256 incremental diverge do arquivo gravado em disco"
            )
        logger.info(
            "PERF sha256_download arquivo=%s incremental=%s referencia=%s equivalente=true",
            path.name,
            incremental,
            reference_digest,
        )

    @staticmethod
    def _validate_xlsx(path: Path) -> None:
        if not path.is_file() or path.stat().st_size == 0:
            raise SharePointReadError("Versão vazia ou ausente")
        try:
            with zipfile.ZipFile(path) as archive:
                if (
                    "[Content_Types].xml" not in archive.namelist()
                    or "xl/workbook.xml" not in archive.namelist()
                ):
                    raise SharePointReadError(
                        "Arquivo não é Open XML com xl/workbook.xml"
                    )
        except zipfile.BadZipFile as error:
            raise SharePointReadError("Versão não é um XLSX/ZIP válido") from error

    def _validate_spreadsheet(self, spreadsheet: SpreadsheetInfo) -> None:
        if (
            spreadsheet.site_id != self.site_url
            or spreadsheet.drive_id != self.REST_CONTEXT
            or not spreadsheet.drive_item_id
            or not spreadsheet.path
        ):
            raise ValueError("A planilha não pertence ao site REST configurado")


SharePointSource = BrowserSharePointSource
