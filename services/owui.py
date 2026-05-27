import httpx

from models.owui import OWUIFile


async def upload_file(
    filename: str,
    data: bytes,
    content_type: str,
    token: str,
    base_url: str,
    timeout: float = 30.0,
) -> OWUIFile:
    async with httpx.AsyncClient(timeout=timeout) as client:
        response = await client.post(
            url=f"{base_url}/api/v1/files/",
            headers={"Authorization": f"Bearer {token}"},
            files={"file": (filename, data, content_type)},
        )

    response.raise_for_status()
    return OWUIFile.model_validate(response.json())
