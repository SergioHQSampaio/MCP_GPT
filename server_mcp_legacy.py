from __future__ import annotations

import base64
import io
import mimetypes
import os
import posixpath
import re
from dataclasses import dataclass
from typing import Any, Literal

from dotenv import load_dotenv
from google.oauth2 import service_account
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaInMemoryUpload
from mcp.server.fastmcp import FastMCP

load_dotenv()

SCOPES = [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/documents",
    "https://www.googleapis.com/auth/spreadsheets",
]

MCP_HOST = os.environ.get("MCP_HOST", "0.0.0.0")
MCP_PORT = int(os.environ.get("MCP_PORT", "8787"))
MCP_MOUNT_PATH = os.environ.get("MCP_MOUNT_PATH", "/")
MCP_STREAMABLE_HTTP_PATH = os.environ.get("MCP_STREAMABLE_HTTP_PATH", "/mcp")
ROOT_FOLDER_ID = os.environ.get("GDRIVE_ROOT_FOLDER_ID", "").strip()
DEFAULT_TEXT_MIME = "text/markdown"

if not ROOT_FOLDER_ID:
    raise RuntimeError("GDRIVE_ROOT_FOLDER_ID nao configurado no ambiente.")

SERVER_INSTRUCTIONS = """
Este conector opera exclusivamente dentro de uma pasta-raiz especifica do Google Drive.

Use estas ferramentas quando o usuario pedir para:
- listar arquivos ou subpastas;
- criar subpastas;
- criar ou editar Google Docs;
- criar ou editar Google Sheets;
- salvar arquivos de texto simples (.txt, .md, .json, .csv);
- enviar arquivos binarios pequenos em base64 (PDF, imagens, outros);
- renomear, mover, pesquisar ou enviar itens para a lixeira.

Regras obrigatorias:
- Nunca opere fora da pasta-raiz configurada.
- Nunca use IDs externos sem validar se pertencem a subarvore permitida.
- Para PDF e imagens, trate edicao como substituicao/upload de nova versao, nao como edicao interna.
- Quando o usuario informar apenas um caminho relativo, resolva sempre a partir da pasta-raiz.
- Para Google Docs e Google Sheets, prefira ferramentas nativas de leitura/escrita em vez de upload binario.
""".strip()

mcp = FastMCP(
    name="Google Drive Bridge",
    instructions=SERVER_INSTRUCTIONS,
    host=MCP_HOST,
    port=MCP_PORT,
    mount_path=MCP_MOUNT_PATH,
    streamable_http_path=MCP_STREAMABLE_HTTP_PATH,
)


def _slug(name: str) -> str:
    name = (name or "").strip()
    name = re.sub(r"[\\/:*?\"<>|]+", "-", name)
    name = re.sub(r"\s+", " ", name).strip()
    return name or "item"


@dataclass
class GoogleClients:
    drive: Any
    docs: Any
    sheets: Any


def _build_credentials() -> Credentials:
    service_account_json = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON", "").strip()
    service_account_file = os.environ.get("GOOGLE_SERVICE_ACCOUNT_FILE", "").strip()
    delegated_user = os.environ.get("GOOGLE_DELEGATED_USER", "").strip()

    if service_account_json:
        info = __import__("json").loads(service_account_json)
        creds = service_account.Credentials.from_service_account_info(info, scopes=SCOPES)
        if delegated_user:
            creds = creds.with_subject(delegated_user)
        return creds

    if service_account_file:
        creds = service_account.Credentials.from_service_account_file(service_account_file, scopes=SCOPES)
        if delegated_user:
            creds = creds.with_subject(delegated_user)
        return creds

    refresh_token = os.environ.get("GOOGLE_REFRESH_TOKEN", "").strip()
    client_id = os.environ.get("GOOGLE_CLIENT_ID", "").strip()
    client_secret = os.environ.get("GOOGLE_CLIENT_SECRET", "").strip()
    token_uri = os.environ.get("GOOGLE_TOKEN_URI", "https://oauth2.googleapis.com/token").strip()

    if refresh_token and client_id and client_secret:
        creds = Credentials(
            token=None,
            refresh_token=refresh_token,
            token_uri=token_uri,
            client_id=client_id,
            client_secret=client_secret,
            scopes=SCOPES,
        )
        creds.refresh(Request())
        return creds

    raise RuntimeError(
        "Configure autenticacao Google: GOOGLE_SERVICE_ACCOUNT_FILE / GOOGLE_SERVICE_ACCOUNT_JSON "
        "ou GOOGLE_REFRESH_TOKEN + GOOGLE_CLIENT_ID + GOOGLE_CLIENT_SECRET."
    )


def _clients() -> GoogleClients:
    creds = _build_credentials()
    return GoogleClients(
        drive=build("drive", "v3", credentials=creds, cache_discovery=False),
        docs=build("docs", "v1", credentials=creds, cache_discovery=False),
        sheets=build("sheets", "v4", credentials=creds, cache_discovery=False),
    )


def _root_metadata(clients: GoogleClients) -> dict[str, Any]:
    return clients.drive.files().get(
        fileId=ROOT_FOLDER_ID,
        fields="id,name,mimeType,parents,trashed",
        supportsAllDrives=True,
    ).execute()


def _normalize_relpath(path: str | None) -> str:
    raw = (path or "").strip()
    raw = raw.replace("\\", "/")
    raw = posixpath.normpath(raw)
    if raw in (".", "/"):
        return ""
    if raw.startswith("../") or raw == "..":
        raise ValueError("Caminho relativo invalido; sai da pasta-raiz.")
    return raw.strip("/")


def _get_file(clients: GoogleClients, file_id: str, fields: str = "id,name,mimeType,parents,trashed") -> dict[str, Any]:
    return clients.drive.files().get(
        fileId=file_id,
        fields=fields,
        supportsAllDrives=True,
    ).execute()


def _list_children(clients: GoogleClients, parent_id: str) -> list[dict[str, Any]]:
    results = []
    page_token = None
    while True:
        resp = clients.drive.files().list(
            q=f"'{parent_id}' in parents and trashed = false",
            fields="nextPageToken, files(id,name,mimeType,parents,modifiedTime,size,webViewLink,iconLink)",
            pageToken=page_token,
            pageSize=1000,
            includeItemsFromAllDrives=True,
            supportsAllDrives=True,
        ).execute()
        results.extend(resp.get("files", []))
        page_token = resp.get("nextPageToken")
        if not page_token:
            break
    return results


def _child_by_name(clients: GoogleClients, parent_id: str, name: str) -> dict[str, Any] | None:
    escaped = name.replace("'", "\\'")
    resp = clients.drive.files().list(
        q=f"'{parent_id}' in parents and name = '{escaped}' and trashed = false",
        fields="files(id,name,mimeType,parents,modifiedTime,size,webViewLink)",
        pageSize=10,
        includeItemsFromAllDrives=True,
        supportsAllDrives=True,
    ).execute()
    files = resp.get("files", [])
    return files[0] if files else None


def _resolve_path(clients: GoogleClients, path: str | None, create_missing_folders: bool = False, final_folder: bool = True) -> dict[str, Any]:
    rel = _normalize_relpath(path)
    current = _root_metadata(clients)
    if not rel:
        return current
    parts = [p for p in rel.split("/") if p]
    for idx, part in enumerate(parts):
        child = _child_by_name(clients, current["id"], part)
        is_last = idx == len(parts) - 1
        if child is None:
            if create_missing_folders or (final_folder and is_last):
                child = clients.drive.files().create(
                    body={
                        "name": part,
                        "mimeType": "application/vnd.google-apps.folder",
                        "parents": [current["id"]],
                    },
                    fields="id,name,mimeType,parents,webViewLink",
                    supportsAllDrives=True,
                ).execute()
            else:
                raise FileNotFoundError(f"Caminho nao encontrado: {rel}")
        current = child
    return current


def _ensure_inside_root(clients: GoogleClients, file_id: str) -> dict[str, Any]:
    current = _get_file(clients, file_id, fields="id,name,mimeType,parents,trashed,webViewLink")
    visited = set()
    probe = current
    while True:
        if probe["id"] == ROOT_FOLDER_ID:
            return current
        if probe["id"] in visited:
            break
        visited.add(probe["id"])
        parents = probe.get("parents") or []
        if not parents:
            break
        probe = _get_file(clients, parents[0], fields="id,name,mimeType,parents")
    raise PermissionError("O item solicitado esta fora da pasta-raiz permitida.")


def _metadata(file_obj: dict[str, Any]) -> dict[str, Any]:
    return {
        "id": file_obj.get("id"),
        "name": file_obj.get("name"),
        "mimeType": file_obj.get("mimeType"),
        "parents": file_obj.get("parents"),
        "modifiedTime": file_obj.get("modifiedTime"),
        "size": file_obj.get("size"),
        "webViewLink": file_obj.get("webViewLink"),
    }


def _doc_text(clients: GoogleClients, document_id: str) -> str:
    doc = clients.docs.documents().get(documentId=document_id).execute()
    chunks: list[str] = []
    for el in doc.get("body", {}).get("content", []):
        para = el.get("paragraph")
        if not para:
            continue
        for pe in para.get("elements", []):
            text_run = pe.get("textRun")
            if text_run and text_run.get("content"):
                chunks.append(text_run["content"])
    return "".join(chunks)


def _clear_and_write_doc(clients: GoogleClients, document_id: str, content: str) -> None:
    doc = clients.docs.documents().get(documentId=document_id).execute()
    end_index = doc.get("body", {}).get("content", [{}])[-1].get("endIndex", 1)
    requests = []
    if end_index and end_index > 2:
        requests.append({
            "deleteContentRange": {
                "range": {"startIndex": 1, "endIndex": end_index - 1}
            }
        })
    if content:
        requests.append({
            "insertText": {
                "location": {"index": 1},
                "text": content,
            }
        })
    if requests:
        clients.docs.documents().batchUpdate(documentId=document_id, body={"requests": requests}).execute()


def _sheet_values(clients: GoogleClients, spreadsheet_id: str, a1_range: str) -> list[list[Any]]:
    resp = clients.sheets.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=a1_range,
    ).execute()
    return resp.get("values", [])


def _sheet_write(clients: GoogleClients, spreadsheet_id: str, a1_range: str, values: list[list[Any]]) -> None:
    clients.sheets.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range=a1_range,
        valueInputOption="USER_ENTERED",
        body={"values": values},
    ).execute()


def _guess_mime(filename: str, fallback: str = "application/octet-stream") -> str:
    guessed, _ = mimetypes.guess_type(filename)
    return guessed or fallback


@mcp.tool(
    name="gdrive_health",
    title="Saude do conector Google Drive",
    description="Valida autenticacao, acesso a pasta-raiz e configuracao basica do conector.",
)
def gdrive_health() -> dict[str, Any]:
    clients = _clients()
    root = _root_metadata(clients)
    return {"ok": True, "root": _metadata(root)}


@mcp.tool(
    name="list_items",
    title="Listar itens",
    description="Lista arquivos e subpastas de um caminho relativo dentro da pasta-raiz do Google Drive.",
)
def list_items(path: str = "") -> dict[str, Any]:
    clients = _clients()
    folder = _resolve_path(clients, path, final_folder=False)
    if folder["mimeType"] != "application/vnd.google-apps.folder":
        raise ValueError("O caminho informado nao e uma pasta.")
    items = [_metadata(x) for x in _list_children(clients, folder["id"])]
    return {"ok": True, "path": _normalize_relpath(path), "items": items}


@mcp.tool(
    name="create_folder",
    title="Criar pasta",
    description="Cria uma subpasta dentro de um caminho relativo da pasta-raiz do Google Drive.",
)
def create_folder(parent_path: str, name: str) -> dict[str, Any]:
    clients = _clients()
    parent = _resolve_path(clients, parent_path, final_folder=False)
    if parent["mimeType"] != "application/vnd.google-apps.folder":
        raise ValueError("parent_path precisa apontar para uma pasta.")
    name = _slug(name)
    existing = _child_by_name(clients, parent["id"], name)
    if existing and existing["mimeType"] == "application/vnd.google-apps.folder":
        return {"ok": True, "created": False, "folder": _metadata(existing)}
    folder = clients.drive.files().create(
        body={
            "name": name,
            "mimeType": "application/vnd.google-apps.folder",
            "parents": [parent["id"]],
        },
        fields="id,name,mimeType,parents,webViewLink",
        supportsAllDrives=True,
    ).execute()
    return {"ok": True, "created": True, "folder": _metadata(folder)}


@mcp.tool(
    name="create_doc",
    title="Criar Google Doc",
    description="Cria um Google Doc nativo dentro de uma pasta da subarvore permitida e, opcionalmente, grava conteudo inicial.",
)
def create_doc(parent_path: str, title: str, content: str = "") -> dict[str, Any]:
    clients = _clients()
    parent = _resolve_path(clients, parent_path, final_folder=False)
    doc = clients.drive.files().create(
        body={
            "name": _slug(title),
            "mimeType": "application/vnd.google-apps.document",
            "parents": [parent["id"]],
        },
        fields="id,name,mimeType,parents,webViewLink",
        supportsAllDrives=True,
    ).execute()
    if content:
        _clear_and_write_doc(clients, doc["id"], content)
    return {"ok": True, "doc": _metadata(doc)}


@mcp.tool(
    name="read_doc",
    title="Ler Google Doc",
    description="Le o texto integral de um Google Doc nativo dentro da pasta-raiz permitida.",
)
def read_doc(file_id: str) -> dict[str, Any]:
    clients = _clients()
    file_obj = _ensure_inside_root(clients, file_id)
    if file_obj["mimeType"] != "application/vnd.google-apps.document":
        raise ValueError("O item informado nao e um Google Doc.")
    return {"ok": True, "file": _metadata(file_obj), "text": _doc_text(clients, file_id)}


@mcp.tool(
    name="write_doc",
    title="Sobrescrever Google Doc",
    description="Sobrescreve integralmente o conteudo de um Google Doc nativo dentro da pasta-raiz permitida.",
)
def write_doc(file_id: str, content: str) -> dict[str, Any]:
    clients = _clients()
    file_obj = _ensure_inside_root(clients, file_id)
    if file_obj["mimeType"] != "application/vnd.google-apps.document":
        raise ValueError("O item informado nao e um Google Doc.")
    _clear_and_write_doc(clients, file_id, content)
    return {"ok": True, "file": _metadata(file_obj)}


@mcp.tool(
    name="create_sheet",
    title="Criar Google Sheet",
    description="Cria uma planilha Google Sheets dentro de uma pasta da subarvore permitida.",
)
def create_sheet(parent_path: str, title: str, sheet_name: str = "Pagina1") -> dict[str, Any]:
    clients = _clients()
    parent = _resolve_path(clients, parent_path, final_folder=False)
    spreadsheet = clients.sheets.spreadsheets().create(
        body={
            "properties": {"title": _slug(title)},
            "sheets": [{"properties": {"title": _slug(sheet_name)}}],
        }
    ).execute()
    file_id = spreadsheet["spreadsheetId"]
    # move para pasta destino
    current = _get_file(clients, file_id, fields="id,name,mimeType,parents,webViewLink")
    prev_parents = ",".join(current.get("parents", []))
    moved = clients.drive.files().update(
        fileId=file_id,
        addParents=parent["id"],
        removeParents=prev_parents,
        fields="id,name,mimeType,parents,webViewLink",
        supportsAllDrives=True,
    ).execute()
    return {"ok": True, "sheet": _metadata(moved)}


@mcp.tool(
    name="read_sheet_range",
    title="Ler intervalo de planilha",
    description="Le um intervalo A1 de uma Google Sheet dentro da pasta-raiz permitida.",
)
def read_sheet_range(file_id: str, a1_range: str) -> dict[str, Any]:
    clients = _clients()
    file_obj = _ensure_inside_root(clients, file_id)
    if file_obj["mimeType"] != "application/vnd.google-apps.spreadsheet":
        raise ValueError("O item informado nao e uma Google Sheet.")
    values = _sheet_values(clients, file_id, a1_range)
    return {"ok": True, "file": _metadata(file_obj), "range": a1_range, "values": values}


@mcp.tool(
    name="write_sheet_range",
    title="Gravar intervalo de planilha",
    description="Escreve valores em um intervalo A1 de uma Google Sheet dentro da pasta-raiz permitida.",
)
def write_sheet_range(file_id: str, a1_range: str, values: list[list[Any]]) -> dict[str, Any]:
    clients = _clients()
    file_obj = _ensure_inside_root(clients, file_id)
    if file_obj["mimeType"] != "application/vnd.google-apps.spreadsheet":
        raise ValueError("O item informado nao e uma Google Sheet.")
    _sheet_write(clients, file_id, a1_range, values)
    return {"ok": True, "file": _metadata(file_obj), "range": a1_range}


@mcp.tool(
    name="upload_text_file",
    title="Salvar arquivo de texto simples",
    description="Cria ou substitui um arquivo de texto simples (.txt, .md, .json, .csv) dentro de uma pasta da subarvore permitida.",
)
def upload_text_file(parent_path: str, filename: str, content: str, mime_type: str = DEFAULT_TEXT_MIME, replace_if_exists: bool = True) -> dict[str, Any]:
    clients = _clients()
    parent = _resolve_path(clients, parent_path, final_folder=False)
    filename = _slug(filename)
    media = MediaInMemoryUpload(content.encode("utf-8"), mimetype=mime_type, resumable=False)
    existing = _child_by_name(clients, parent["id"], filename)
    if existing:
        if not replace_if_exists:
            raise ValueError("O arquivo ja existe e replace_if_exists=false.")
        updated = clients.drive.files().update(
            fileId=existing["id"],
            media_body=media,
            fields="id,name,mimeType,parents,webViewLink,modifiedTime,size",
            supportsAllDrives=True,
        ).execute()
        return {"ok": True, "created": False, "file": _metadata(updated)}
    created = clients.drive.files().create(
        body={"name": filename, "parents": [parent["id"]]},
        media_body=media,
        fields="id,name,mimeType,parents,webViewLink,modifiedTime,size",
        supportsAllDrives=True,
    ).execute()
    return {"ok": True, "created": True, "file": _metadata(created)}


@mcp.tool(
    name="upload_base64_file",
    title="Salvar arquivo binario em base64",
    description="Cria ou substitui um arquivo binario pequeno (PDF, imagem, outros) a partir de conteudo base64, dentro da subarvore permitida.",
)
def upload_base64_file(parent_path: str, filename: str, base64_content: str, mime_type: str = "application/octet-stream", replace_if_exists: bool = True) -> dict[str, Any]:
    clients = _clients()
    parent = _resolve_path(clients, parent_path, final_folder=False)
    filename = _slug(filename)
    binary = base64.b64decode(base64_content)
    media = MediaInMemoryUpload(binary, mimetype=mime_type or _guess_mime(filename), resumable=False)
    existing = _child_by_name(clients, parent["id"], filename)
    if existing:
        if not replace_if_exists:
            raise ValueError("O arquivo ja existe e replace_if_exists=false.")
        updated = clients.drive.files().update(
            fileId=existing["id"],
            media_body=media,
            fields="id,name,mimeType,parents,webViewLink,modifiedTime,size",
            supportsAllDrives=True,
        ).execute()
        return {"ok": True, "created": False, "file": _metadata(updated)}
    created = clients.drive.files().create(
        body={"name": filename, "parents": [parent["id"]]},
        media_body=media,
        fields="id,name,mimeType,parents,webViewLink,modifiedTime,size",
        supportsAllDrives=True,
    ).execute()
    return {"ok": True, "created": True, "file": _metadata(created)}


@mcp.tool(
    name="download_file_metadata",
    title="Ler metadados de arquivo",
    description="Le metadados de um item dentro da subarvore permitida do Google Drive.",
)
def download_file_metadata(file_id: str) -> dict[str, Any]:
    clients = _clients()
    file_obj = _ensure_inside_root(clients, file_id)
    return {"ok": True, "file": _metadata(file_obj)}


@mcp.tool(
    name="read_text_file",
    title="Ler arquivo de texto simples",
    description="Le o conteudo de um arquivo de texto simples dentro da pasta-raiz permitida.",
)
def read_text_file(file_id: str) -> dict[str, Any]:
    clients = _clients()
    file_obj = _ensure_inside_root(clients, file_id)
    mime_type = file_obj.get("mimeType", "")
    if mime_type.startswith("application/vnd.google-apps"):
        raise ValueError("Use as ferramentas nativas de Google Docs ou Google Sheets para este arquivo.")
    request = clients.drive.files().get_media(fileId=file_id, supportsAllDrives=True)
    buffer = io.BytesIO()
    downloader = __import__("googleapiclient.http").http.MediaIoBaseDownload(buffer, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    text = buffer.getvalue().decode("utf-8", errors="replace")
    return {"ok": True, "file": _metadata(file_obj), "text": text}


@mcp.tool(
    name="move_item",
    title="Mover item",
    description="Move um arquivo ou subpasta para outra pasta da subarvore permitida.",
)
def move_item(file_id: str, destination_folder_path: str) -> dict[str, Any]:
    clients = _clients()
    file_obj = _ensure_inside_root(clients, file_id)
    dest = _resolve_path(clients, destination_folder_path, final_folder=False)
    prev_parents = ",".join(file_obj.get("parents", []))
    moved = clients.drive.files().update(
        fileId=file_id,
        addParents=dest["id"],
        removeParents=prev_parents,
        fields="id,name,mimeType,parents,webViewLink",
        supportsAllDrives=True,
    ).execute()
    return {"ok": True, "file": _metadata(moved)}


@mcp.tool(
    name="rename_item",
    title="Renomear item",
    description="Renomeia um arquivo ou pasta dentro da subarvore permitida.",
)
def rename_item(file_id: str, new_name: str) -> dict[str, Any]:
    clients = _clients()
    _ensure_inside_root(clients, file_id)
    updated = clients.drive.files().update(
        fileId=file_id,
        body={"name": _slug(new_name)},
        fields="id,name,mimeType,parents,webViewLink",
        supportsAllDrives=True,
    ).execute()
    return {"ok": True, "file": _metadata(updated)}


@mcp.tool(
    name="trash_item",
    title="Enviar item para a lixeira",
    description="Envia um arquivo ou pasta para a lixeira dentro da subarvore permitida.",
)
def trash_item(file_id: str) -> dict[str, Any]:
    clients = _clients()
    _ensure_inside_root(clients, file_id)
    trashed = clients.drive.files().update(
        fileId=file_id,
        body={"trashed": True},
        fields="id,name,mimeType,trashed,parents,webViewLink",
        supportsAllDrives=True,
    ).execute()
    return {"ok": True, "file": _metadata(trashed)}


@mcp.tool(
    name="restore_item",
    title="Restaurar item da lixeira",
    description="Restaura um arquivo ou pasta da lixeira, desde que o item pertença a subarvore permitida.",
)
def restore_item(file_id: str) -> dict[str, Any]:
    clients = _clients()
    _ensure_inside_root(clients, file_id)
    restored = clients.drive.files().update(
        fileId=file_id,
        body={"trashed": False},
        fields="id,name,mimeType,trashed,parents,webViewLink",
        supportsAllDrives=True,
    ).execute()
    return {"ok": True, "file": _metadata(restored)}


@mcp.tool(
    name="search_items",
    title="Pesquisar itens",
    description="Pesquisa itens por nome dentro da subarvore permitida do Google Drive. Opcionalmente filtra por mime type.",
)
def search_items(query: str, path: str = "", mime_type: str = "") -> dict[str, Any]:
    clients = _clients()
    folder = _resolve_path(clients, path, final_folder=False)
    query = (query or "").strip()
    if not query:
        raise ValueError("query e obrigatoria.")
    clauses = [f"'{folder['id']}' in parents", "trashed = false"]
    for token in query.split():
        token = token.replace("'", "\\'")
        clauses.append(f"name contains '{token}'")
    if mime_type:
        mime_type = mime_type.replace("'", "\\'")
        clauses.append(f"mimeType = '{mime_type}'")
    resp = clients.drive.files().list(
        q=" and ".join(clauses),
        fields="files(id,name,mimeType,parents,modifiedTime,size,webViewLink)",
        pageSize=100,
        includeItemsFromAllDrives=True,
        supportsAllDrives=True,
    ).execute()
    return {"ok": True, "items": [_metadata(x) for x in resp.get("files", [])]}


if __name__ == "__main__":
    mcp.run(transport="streamable-http")
