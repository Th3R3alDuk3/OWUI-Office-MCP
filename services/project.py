from asyncio import to_thread
from collections.abc import AsyncGenerator
from contextlib import asynccontextmanager
from dataclasses import dataclass
from mimetypes import guess_extension
from pathlib import Path
from secrets import token_hex
from shutil import copyfile
from tempfile import TemporaryDirectory
from types import ModuleType

from fastmcp import FastMCP
from fastmcp.exceptions import ToolError
from fastmcp.utilities.logging import get_logger
from pydantic import ValidationError

from models.office import AnyInventory, PptxInventory
from services import docx, pptx, xlsx
from services.image import shrink_image
from services.officecli import LANDLOCK_ABI, READ_COMMANDS, run_batch
from services.owui import (
    delete_file,
    download_file,
    find_files,
    upload_file,
)

logger = get_logger(__name__)

_FORMATS: dict[str, ModuleType] = {
    module.FORMAT: module for module in (pptx, docx, xlsx)
}

_TEMPLATES = Path("templates")

# A project is the newest file of this name; OpenWebUI cannot replace one.
_PROJECT_FILE = "project-{project_id}.{format}"

_MAX_DOCUMENT_BYTES = 100 * 2**20

_MAX_ASSET_BYTES = 20 * 2**20
_IMAGE_TYPES = {
    "image/png", "image/jpeg", "image/gif", "image/bmp", "image/tiff",
    "image/webp",
}


@dataclass(frozen=True)
class StoredTemplate:
    module: ModuleType
    document: bytes
    inventory: AnyInventory


@dataclass(frozen=True)
class Project:
    id: str
    file_id: str
    owui_token: str
    document: Path
    module: ModuleType
    inventory: AnyInventory


templates: dict[str, StoredTemplate] = {}


async def fetch_document(
    directory: Path,
    file_id: str,
    owui_token: str,
) -> tuple[ModuleType, Path]:

    data, content_type = await download_file(
        file_id, owui_token, _MAX_DOCUMENT_BYTES,
    )

    suffix = guess_extension(content_type) or ""

    if (module := _FORMATS.get(suffix.removeprefix("."))) is None:
        raise ToolError(
            f"File '{file_id}' is no `.pptx`, `.docx` or `.xlsx` document."
        )

    document = directory / f"document.{module.FORMAT}"
    await to_thread(document.write_bytes, data)

    return module, document


async def _save(
    project_id: str,
    previous_file_id: str | None,
    document: Path,
    inventory: AnyInventory,
    owui_token: str,
) -> None:

    data = await to_thread(document.read_bytes)

    if len(data) > _MAX_DOCUMENT_BYTES:
        raise ToolError(
            f"The document would exceed {_MAX_DOCUMENT_BYTES // 2**20} MB; "
            "nothing was changed."
        )

    document_format = document.suffix.removeprefix(".")
    file_name = _PROJECT_FILE.format(
        project_id=project_id, format=document_format,
    )

    file_id, _ = await upload_file(
        file_name=file_name,
        data=data,
        metadata={
            "format": document_format,
            "inventory": inventory.model_dump_json(),
        },
        owui_token=owui_token,
    )

    if previous_file_id is None:
        return

    # No lock: a version saved by another chat since this edit started wins.
    for file in await find_files(file_name, owui_token):
        if file["id"] == previous_file_id:
            break
        if file["id"] != file_id:
            await delete_file(file_id, owui_token)
            raise ToolError(
                f"Another chat saved project '{project_id}' meanwhile; this "
                "edit was discarded. Read the project again, then repeat the "
                "edit."
            )

    try:
        await delete_file(previous_file_id, owui_token)
    except ToolError as error:
        # The edit is saved; only the previous version stays behind.
        logger.warning(
            "previous version of project %s kept: %s", project_id, error,
        )


async def create_project(
    document: Path,
    inventory: AnyInventory,
    owui_token: str,
) -> str:

    project_id = token_hex(8)

    await _save(
        project_id=project_id,
        previous_file_id=None,
        document=document,
        inventory=inventory,
        owui_token=owui_token,
    )

    return project_id


@asynccontextmanager
async def load_project(
    project_id: str,
    owui_token: str,
) -> AsyncGenerator[Project]:

    files = await find_files(
        _PROJECT_FILE.format(project_id=project_id, format="*"), owui_token,
    )

    if not files:
        raise ToolError(
            f"Project '{project_id}' not found. Pass the `project_id` from "
            "`start_project`, or start a new one."
        )

    file_id = files[0]["id"]
    # What the upload passed as `metadata`.
    fields = (files[0].get("meta") or {}).get("data") or {}

    try:
        module = _FORMATS[fields.get("format", "")]
        inventory = module.INVENTORY.model_validate_json(
            fields.get("inventory", ""),
        )
    except (KeyError, ValidationError):
        raise ToolError(
            f"Project '{project_id}' comes from an older server version. Start "
            "a new one with `start_project`."
        ) from None

    data, _ = await download_file(file_id, owui_token, _MAX_DOCUMENT_BYTES)

    with TemporaryDirectory() as scratch:

        document = Path(scratch) / f"document.{module.FORMAT}"
        await to_thread(document.write_bytes, data)

        yield Project(
            id=project_id,
            file_id=file_id,
            owui_token=owui_token,
            document=document,
            module=module,
            inventory=inventory,
        )


async def add_asset(
    project: Project,
    file_id: str,
) -> str:

    data, content_type = await download_file(
        file_id, project.owui_token, _MAX_ASSET_BYTES,
    )

    if content_type not in _IMAGE_TYPES:
        raise ToolError(
            f"File '{file_id}' is no supported image "
            "(PNG, JPEG, GIF, BMP, TIFF, WebP)."
        )

    suffix = guess_extension(content_type)

    if shrunk := await to_thread(shrink_image, data):
        data, suffix = shrunk

    assets = project.document.parent / "assets"
    assets.mkdir(exist_ok=True)
    name = f"{file_id}{suffix}"
    await to_thread((assets / name).write_bytes, data)

    return f"assets/{name}"


async def apply_batch(
    project: Project,
    batch: list[dict],
) -> list[dict]:

    if all(command["command"] in READ_COMMANDS for command in batch):
        return await run_batch(project.document, batch)

    if isinstance(project.inventory, PptxInventory):
        pptx.resolve_layouts(project.inventory, batch)
        *results, deck = await run_batch(project.document, [*batch, pptx.DECK])
        pptx.check(deck, project.inventory)
    else:
        results = await run_batch(project.document, batch)

    await _save(
        project_id=project.id,
        previous_file_id=project.file_id,
        document=project.document,
        inventory=project.inventory,
        owui_token=project.owui_token,
    )

    return results


@asynccontextmanager
async def project_lifespan(
    server: FastMCP,
) -> AsyncGenerator[None]:

    # Below ABI 4 (Linux 6.7) Landlock cannot restrict the network.
    if LANDLOCK_ABI < 4:
        logger.warning(
            "Landlock ABI %d: OfficeCLI is confined only as far as this kernel "
            "allows and keeps network access.", LANDLOCK_ABI,
        )

    for source in sorted(_TEMPLATES.iterdir()):

        if (module := _FORMATS.get(source.suffix.removeprefix("."))) is None:
            continue

        try:
            with TemporaryDirectory() as scratch:
                # OfficeCLI picks its handler by suffix.
                document = Path(scratch) / f"document.{module.FORMAT}"
                await to_thread(copyfile, source, document)
                await module.prepare(document)
                templates[source.name] = StoredTemplate(
                    module=module,
                    document=await to_thread(document.read_bytes),
                    inventory=await module.inventory(document),
                )
        except Exception as error:
            logger.warning("template %s skipped: %s", source.name, error)

    yield
