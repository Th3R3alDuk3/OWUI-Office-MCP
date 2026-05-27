from httpx import AsyncClient, HTTPStatusError, RequestError

from models.owui import OWUIFile


async def upload_file(
    filename: str,
    data: bytes,
    content_type: str,
    token: str,
    base_url: str,
) -> OWUIFile:

    try:

        async with AsyncClient(verify=False) as client:
            response = await client.post(
                url=f"{base_url}/api/v1/files/",
                headers={"Authorization": f"Bearer {token}"},
                files={"file": (filename, data, content_type)},
            )

        response.raise_for_status()

        response_json = response.json()
        return OWUIFile.model_validate(response_json)

    except HTTPStatusError as error:
        raise RuntimeError(
            f"OpenWebUI rejected the upload."
            f" Detail: {error}"
        ) from error
    except RequestError as error:
        raise RuntimeError(
            f"Could not reach OpenWebUI at {base_url}."
            f" Detail: {error}"
        ) from error
