import httpx

from models.owui import OWUIFile

PPTX_MIME = (
    "application/vnd.openxmlformats-officedocument.presentationml.presentation"
)


async def upload_to_owui(
    filename: str,
    data: bytes,
    token: str,
    base_url: str,
    timeout: float = 30.0,
) -> OWUIFile:

    async with httpx.AsyncClient(timeout=timeout) as client:
        response = await client.post(
            url=f"{base_url.rstrip('/')}/api/v1/files/",
            headers={"Authorization": f"Bearer {token}"},
            files={"file": (filename, data, PPTX_MIME)},
        )

    response.raise_for_status()
    
    response_json = response.json()
    return OWUIFile.model_validate(response_json)
