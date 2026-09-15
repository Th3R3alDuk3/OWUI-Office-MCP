from asyncio import to_thread
from collections.abc import AsyncGenerator
from contextlib import asynccontextmanager, suppress
from dataclasses import dataclass
from io import BytesIO
from json import dumps, loads
from mimetypes import guess_extension
from pathlib import Path
from secrets import token_hex
from tempfile import TemporaryDirectory
from types import ModuleType
from xml.etree import ElementTree
from zipfile import BadZipFile, ZipFile

from fastmcp.exceptions import ToolError
from fastmcp.utilities.logging import get_logger
from valkey.asyncio import Valkey
from valkey.asyncio.lock import Lock
from valkey.exceptions import LockError

from config import get_settings
from models.inventory import DocxInventory, Inventory, PptxInventory, XlsxInventory
from tools._office import docx, guard, officecli, pptx, xlsx
from services.owui import download_file

_settings = get_settings()
_logger = get_logger("office")

FORMATS: dict[str, ModuleType] = {
    module.FORMAT: module for module in (pptx, docx, xlsx)
}

_MANIFEST = "[Content_Types].xml"
_MAX_MANIFEST_BYTES = 2**20

_TEMPLATES = Path("templates")

_PROJECT_KEY = "project:{user_id}:{project_id}"
_LOCK_KEY = "lock:{user_id}:{project_id}"
# A replica that dies mid-edit leaves its lock behind for this long.
_LOCK_TTL_SECONDS = 600

_MAX_UPLOAD_BYTES = 100 * 2**20
_MAX_ASSET_BYTES = 20 * 2**20

_IMAGE_TYPES = {
    "image/png", "image/jpeg", "image/gif", "image/bmp", "image/tiff",
    "image/webp",
}

valkey = Valkey.from_url(_settings.valkey_url)


@dataclass(frozen=True)
class Prepared:
    module: ModuleType
    document: bytes
    inventory: PptxInventory | DocxInventory | XlsxInventory
    baseline: list[str]


@dataclass(frozen=True)
class Project:
    key: str
    document: Path
    module: ModuleType
    inventory: Inventory
    baseline: list[str]


templates: dict[str, Prepared] = {}


def detect_format(
    data: bytes,
) -> ModuleType:

    # OfficeCLI rejects bombs and broken packages; macros only show here.
    try:

        with ZipFile(BytesIO(data)) as archive:

            if archive.getinfo(_MANIFEST).file_size > _MAX_MANIFEST_BYTES:
                raise ToolError("The file is too large when unpacked.")

            names = set(archive.namelist())
            content_types = {
                entry.get("ContentType")
                for entry in ElementTree.fromstring(archive.read(_MANIFEST))
                if entry.get("PartName", "").lstrip("/") in names
            }

    except (BadZipFile, KeyError, ElementTree.ParseError):
        raise ToolError("The file is not a valid Office document.") from None

    supported = [
        module for module in FORMATS.values()
        if module.CONTENT_TYPE in content_types
    ]

    if len(supported) != 1:
        raise ToolError(
            "Expected one `.pptx`, `.docx` or `.xlsx` document "
            "(no macros or template formats such as `.potx`)."
        )

    return supported[0]


async def prepare_templates() -> None:

    # Once per start; changed templates need a restart.
    for source in sorted(_TEMPLATES.iterdir()):

        if source.suffix.removeprefix(".") not in FORMATS:
            continue

        try:
            with TemporaryDirectory() as scratch:
                data = await to_thread(source.read_bytes)
                module = detect_format(data)
                # OfficeCLI picks its handler by suffix.
                document = Path(scratch) / f"document.{module.FORMAT}"
                await to_thread(document.write_bytes, data)
                await module.prepare(document)
                inventory = await module.inventory(document)
                baseline = await officecli.findings(document)
                templates[source.name] = Prepared(
                    module=module,
                    document=await to_thread(document.read_bytes),
                    inventory=inventory,
                    baseline=baseline,
                )
        except (ToolError, RuntimeError) as error:
            _logger.warning("template %s skipped: %s", source.name, error)


async def check_limit(
    user_id: str,
) -> None:

    pattern = _PROJECT_KEY.format(user_id=user_id, project_id="*")
    held = [key async for key in valkey.scan_iter(match=pattern, count=1000)]

    if len(held) >= _settings.project_max_per_user:
        raise ToolError(
            f"You have {_settings.project_max_per_user} projects, the "
            "maximum. Old projects expire after inactivity; retry later."
        )


async def _download(
    file_id: str,
    token: str,
    max_bytes: int,
) -> tuple[bytes, str]:

    try:
        data, content_type = await download_file(
            file_id=file_id,
            token=token,
            max_bytes=max_bytes,
        )
    except RuntimeError as error:
        raise ToolError(
            f"Could not fetch file '{file_id}' from OpenWebUI. Check that the "
            "ID belongs to a file the user actually attached."
        ) from error

    if len(data) > max_bytes:
        raise ToolError(
            f"File '{file_id}' is too large. Limit is {max_bytes // 2**20} MB."
        )

    return data, content_type


async def fetch_upload(
    directory: Path,
    file_id: str,
    token: str,
) -> ModuleType:

    data, _ = await _download(file_id, token, _MAX_UPLOAD_BYTES)
    module = detect_format(data)
    await to_thread((directory / f"document.{module.FORMAT}").write_bytes, data)

    return module


async def _store(
    key: str,
    document: Path,
    inventory: Inventory,
    baseline: list[str],
) -> None:

    # One transaction, so a project never exists without its TTL.
    async with valkey.pipeline() as pipe:
        pipe.hset(key, mapping={
            "document": await to_thread(document.read_bytes),
            "format": document.suffix.removeprefix("."),
            "inventory": inventory.model_dump_json(),
            "baseline": dumps(baseline),
        })
        pipe.expire(key, _settings.project_ttl)
        await pipe.execute()


async def publish(
    user_id: str,
    document: Path,
    inventory: Inventory,
    baseline: list[str],
) -> str:

    project_id = token_hex(8)
    await _store(
        _PROJECT_KEY.format(user_id=user_id, project_id=project_id),
        document, inventory, baseline,
    )

    return project_id


@asynccontextmanager
async def load(
    user_id: str,
    project_id: str,
) -> AsyncGenerator[Project]:

    key = _PROJECT_KEY.format(user_id=user_id, project_id=project_id)
    # valkey-py types the async commands like the sync ones.
    fields: dict[bytes, bytes] = await valkey.hgetall(key)  # pyright: ignore[reportGeneralTypeIssues]

    if not fields:
        raise ToolError(
            f"Project '{project_id}' not found. Start one with "
            "`create_project` or `open_project`."
        )

    # Every access counts as activity.
    await valkey.expire(key, _settings.project_ttl)

    module = FORMATS[fields[b"format"].decode()]

    with TemporaryDirectory() as scratch:

        document = Path(scratch) / f"document.{module.FORMAT}"
        await to_thread(document.write_bytes, fields[b"document"])

        yield Project(
            key=key,
            document=document,
            module=module,
            inventory=module.INVENTORY.model_validate_json(fields[b"inventory"]),
            baseline=loads(fields[b"baseline"]),
        )


@asynccontextmanager
async def locked(
    user_id: str,
    project_id: str,
) -> AsyncGenerator[Lock]:

    lock = valkey.lock(
        _LOCK_KEY.format(user_id=user_id, project_id=project_id),
        timeout=_LOCK_TTL_SECONDS,
        blocking_timeout=_settings.officecli_queue_timeout,
    )

    if not await lock.acquire():
        raise ToolError(
            "The project is busy; nothing was changed. Retry in a moment."
        )

    try:
        yield lock
    finally:
        # Releases only its own token; a lock lost meanwhile is no error here.
        with suppress(LockError):
            await lock.release()


async def add_asset(
    project: Project,
    file_id: str,
    token: str,
) -> str:

    data, content_type = await _download(file_id, token, _MAX_ASSET_BYTES)

    if content_type not in _IMAGE_TYPES:
        raise ToolError(
            f"File '{file_id}' is no supported image "
            "(PNG, JPEG, GIF, BMP, TIFF, WebP)."
        )

    assets = project.document.parent / "assets"
    assets.mkdir(exist_ok=True)
    name = f"{file_id}{guess_extension(content_type)}"
    await to_thread((assets / name).write_bytes, data)

    return f"assets/{name}"


async def apply(
    project: Project,
    batch: list[dict],
    lock: Lock,
) -> list[dict]:

    if all(command["command"] in guard.READ_COMMANDS for command in batch):
        return await officecli.run(project.document, batch)

    if isinstance(project.inventory, PptxInventory):
        await to_thread(pptx.check, project.document, project.inventory, batch)

    results = await officecli.run(project.document, batch)

    if isinstance(project.inventory, PptxInventory):
        await to_thread(pptx.check, project.document, project.inventory)

    # An expired or evicted lock means another call may hold the project now.
    if not await lock.owned():
        raise ToolError(
            "Lost the project lock during the batch; nothing was saved. Send "
            "the whole batch again."
        )

    # All fields, so an evicted project comes back whole with this edit.
    await _store(project.key, project.document, project.inventory, project.baseline)

    return results
