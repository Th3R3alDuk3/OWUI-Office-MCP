from collections.abc import AsyncGenerator
from contextlib import asynccontextmanager, suppress
from json import dumps

from fastmcp.exceptions import ToolError
from httpx2 import AsyncClient, HTTPStatusError, RequestError

from config import get_settings

_settings = get_settings()

_REQUEST_TIMEOUT_SECONDS = 60.0

# OpenWebUI's API takes only its own tokens.
_TOKEN_EXCHANGE_PATH = "/api/v1/auths/oauth/oidc/token/exchange"
_FILE_SEARCH_PATH = "/api/v1/files/user/search"
_FILE_DOWNLOAD_PATH = "/api/v1/files/{file_id}/content"
# Projects and exports are never chat context, so no extraction or embeddings.
_FILE_UPLOAD_PATH = "/api/v1/files/?process=false"
_FILE_DELETE_PATH = "/api/v1/files/{file_id}"

# Shared, so calls reuse connections.
_client = AsyncClient(
    base_url=_settings.owui_base_url,
    timeout=_REQUEST_TIMEOUT_SECONDS,
)


@asynccontextmanager
async def _http_errors(
    action: str,
) -> AsyncGenerator[None]:

    try:
        yield
    except HTTPStatusError as error:
        response = error.response
        detail = f"{response.status_code} {response.reason_phrase}"
        with suppress(ValueError, KeyError, TypeError):
            detail = str(response.json()["detail"]).rstrip(".")
        raise ToolError(
            f"OpenWebUI rejected the {action}: {detail}."
        ) from error
    except RequestError as error:
        raise ToolError("Could not reach OpenWebUI.") from error


async def exchange_token(
    token: str,
) -> str:

    async with _http_errors("token exchange"):
        response = await _client.post(
            _TOKEN_EXCHANGE_PATH,
            json={"token": token},
        )
        response.raise_for_status()

    if not (owui_token := response.json().get("token")):
        raise ToolError(
            "OpenWebUI answered the token exchange without a token."
        )

    return owui_token


async def find_files(
    pattern: str,
    owui_token: str,
) -> list[dict]:

    async with _http_errors(f"search for file '{pattern}'"):
        response = await _client.get(
            _FILE_SEARCH_PATH,
            params={"filename": pattern},
            headers={"Authorization": f"Bearer {owui_token}"},
        )
        response.raise_for_status()

    # Newest first.
    return response.json()


async def download_file(
    file_id: str,
    owui_token: str,
    max_bytes: int,
) -> tuple[bytes, str]:

    async with (
        _http_errors(f"download of file '{file_id}'"),
        _client.stream(
            "GET",
            _FILE_DOWNLOAD_PATH.format(file_id=file_id),
            headers={"Authorization": f"Bearer {owui_token}"},
        ) as response,
    ):
        if response.is_error:
            await response.aread()

        response.raise_for_status()

        data = bytearray()

        async for chunk in response.aiter_bytes():
            data += chunk
            if len(data) > max_bytes:
                raise ToolError(
                    f"File '{file_id}' exceeds {max_bytes // 2**20} MB."
                )

    return (
        bytes(data),
        response.headers.get("content-type", "").partition(";")[0],
    )


async def upload_file(
    file_name: str,
    data: bytes,
    metadata: dict,
    owui_token: str,
) -> tuple[str, str]:

    async with _http_errors("upload"):
        response = await _client.post(
            _FILE_UPLOAD_PATH,
            headers={"Authorization": f"Bearer {owui_token}"},
            data={"metadata": dumps(metadata)},
            files={"file": (file_name, data)},
        )
        response.raise_for_status()

    if not (file_id := response.json().get("id")):
        raise ToolError("OpenWebUI answered the upload without a file ID.")

    # Relative to OpenWebUI.
    return file_id, _FILE_DOWNLOAD_PATH.format(file_id=file_id)


async def delete_file(
    file_id: str,
    owui_token: str,
) -> None:

    async with _http_errors(f"deletion of file '{file_id}'"):
        response = await _client.delete(
            _FILE_DELETE_PATH.format(file_id=file_id),
            headers={"Authorization": f"Bearer {owui_token}"},
        )
        response.raise_for_status()
