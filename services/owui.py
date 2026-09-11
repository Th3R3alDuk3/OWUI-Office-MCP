from asyncio import timeout, to_thread
from contextlib import suppress
from pathlib import Path

from httpx2 import AsyncClient, HTTPError, HTTPStatusError, RequestError

from config import get_settings

_settings = get_settings()

_REQUEST_TIMEOUT_SECONDS = 60.0

_FILE_UPLOAD_URL = "{base_url}/api/v1/files/"
_FILE_DOWNLOAD_URL = "{base_url}/api/v1/files/{file_id}/content"
_MESSAGE_EVENT_URL = (
    "{base_url}/api/v1/chats/{chat_id}/messages/{message_id}/event"
)


def _client() -> AsyncClient:
    return AsyncClient(
        verify=_settings.owui_verify_tls,
        timeout=_REQUEST_TIMEOUT_SECONDS,
    )


async def download_file(
    file_id: str,
    token: str,
    target: Path,
    max_bytes: int,
) -> str:

    try:
        async with (
            timeout(_REQUEST_TIMEOUT_SECONDS),
            _client() as client,
            client.stream(
                "GET",
                url=_FILE_DOWNLOAD_URL.format(
                    base_url=_settings.owui_base_url,
                    file_id=file_id,
                ),
                headers={"Authorization": f"Bearer {token}"},
            ) as response,
        ):
            response.raise_for_status()

            size = 0

            with target.open("wb") as file:
                async for chunk in response.aiter_bytes(chunk_size=64 * 1024):
                    # One byte over is enough for the caller to reject it.
                    chunk = chunk[:max_bytes + 1 - size]
                    await to_thread(file.write, chunk)
                    size += len(chunk)
                    if size > max_bytes:
                        break

        return response.headers.get("content-type", "").partition(";")[0]

    except HTTPStatusError as error:
        raise RuntimeError(
            "OpenWebUI rejected the download."
        ) from error
    except (RequestError, TimeoutError) as error:
        raise RuntimeError(
            "Could not reach OpenWebUI."
        ) from error


async def upload_file(
    file: Path,
    file_name: str,
    content_type: str,
    token: str,
    chat_id: str | None,
    message_id: str | None,
) -> str:

    try:
        async with _client() as client:

            async with timeout(_REQUEST_TIMEOUT_SECONDS):
                with file.open("rb") as content:
                    response = await client.post(
                        url=_FILE_UPLOAD_URL.format(
                            base_url=_settings.owui_base_url,
                        ),
                        headers={"Authorization": f"Bearer {token}"},
                        files={"file": (file_name, content, content_type)},
                    )
            response.raise_for_status()

            file_id = response.json().get("id")

            if not file_id:
                raise RuntimeError(
                    "OpenWebUI upload response carries no file id.")

            # The chip on the chat message is a bonus; the link always works.
            if chat_id and message_id:
                with suppress(HTTPError, TimeoutError):
                    async with timeout(10):
                        response = await client.post(
                            url=_MESSAGE_EVENT_URL.format(
                                base_url=_settings.owui_base_url,
                                chat_id=chat_id,
                                message_id=message_id,
                            ),
                            headers={"Authorization": f"Bearer {token}"},
                            json={
                                "type": "files",
                                "data": {"files": [{
                                    "type": "file",
                                    "id": file_id,
                                    "name": file_name,
                                    "url": f"/api/v1/files/{file_id}",
                                }]},
                            },
                        )
                        response.raise_for_status()

        return _FILE_DOWNLOAD_URL.format(
            base_url=_settings.owui_public_url or _settings.owui_base_url,
            file_id=file_id,
        )

    except HTTPStatusError as error:
        raise RuntimeError(
            "OpenWebUI rejected the upload."
        ) from error
    except (RequestError, TimeoutError) as error:
        raise RuntimeError(
            "Could not reach OpenWebUI."
        ) from error
