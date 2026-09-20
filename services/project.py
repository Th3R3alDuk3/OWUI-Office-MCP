from asyncio import to_thread
from collections.abc import AsyncGenerator
from contextlib import asynccontextmanager, suppress
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
from valkey.asyncio import Valkey
from valkey.asyncio.lock import Lock
from valkey.exceptions import LockError, WatchError

from config import get_settings
from models.office import AnyInventory, PptxInventory
from services import docx, pptx, xlsx
from services.officecli import LANDLOCK_ABI, READ_COMMANDS, run_batch
from services.owui import download_file

_settings = get_settings()
logger = get_logger(__name__)

FORMATS: dict[str, ModuleType] = {
    module.FORMAT: module for module in (pptx, docx, xlsx)
}
# Macro and template variants carry other types and fall through.
_FORMATS_BY_TYPE = {module.MIME: module for module in FORMATS.values()}

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


@dataclass(frozen=True)
class StoredTemplate:
    module: ModuleType
    document: bytes
    inventory: AnyInventory


@dataclass(frozen=True)
class Project:
    key: str
    document: Path
    module: ModuleType
    inventory: AnyInventory


_valkey = Valkey.from_url(_settings.valkey_url)
templates: dict[str, StoredTemplate] = {}


async def prepare_templates() -> None:

    for source in sorted(_TEMPLATES.iterdir()):

        if (module := FORMATS.get(source.suffix.removeprefix("."))) is None:
            continue

        try:
            with TemporaryDirectory() as scratch:
                # OfficeCLI picks its handler by suffix.
                document = Path(scratch) / f"document.{module.FORMAT}"
                await to_thread(copyfile, source, document)
                await module.prepare(document)
                # xlsx's inventory writes cell formats, so it goes first.
                inventory = await module.inventory(document)
                templates[source.name] = StoredTemplate(
                    module=module,
                    document=await to_thread(document.read_bytes),
                    inventory=inventory,
                )
        except (ToolError, RuntimeError) as error:
            logger.warning("template %s skipped: %s", source.name, error)


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
) -> tuple[ModuleType, Path]:

    data, content_type = await _download(file_id, token, _MAX_UPLOAD_BYTES)

    if (module := _FORMATS_BY_TYPE.get(content_type)) is None:
        raise ToolError(
            f"File '{file_id}' is no `.pptx`, `.docx` or `.xlsx` document."
        )

    document = directory / f"document.{module.FORMAT}"
    await to_thread(document.write_bytes, data)

    return module, document


async def bind_master(
    document: Path,
    inventory: AnyInventory,
    master: int | None,
) -> AnyInventory:

    if isinstance(inventory, PptxInventory):
        inventory = pptx.bind(inventory, master)
        # Existing slides must belong to the bound master.
        await to_thread(pptx.check, document, inventory)
    elif master is not None:
        raise ToolError("`master` applies to pptx designs only.")

    return inventory


async def _store(
    key: str,
    document: Path,
    inventory: AnyInventory,
    lock: Lock | None = None,
) -> None:

    fields = {
        "document": await to_thread(document.read_bytes),
        "format": document.suffix.removeprefix("."),
        "inventory": inventory.model_dump_json(),
    }

    try:
        # One transaction: no project without TTL, no write under a lost lock.
        async with _valkey.pipeline() as pipe:
            if lock is not None:
                await pipe.watch(lock.name)
                if await pipe.get(lock.name) != lock.local.token:
                    raise WatchError
                pipe.multi()
            pipe.hset(key, mapping=fields)
            pipe.expire(key, _settings.project_ttl)
            await pipe.execute()
    except WatchError:
        raise ToolError(
            "Lost the project lock during the batch; nothing was saved. Send "
            "the whole batch again."
        ) from None


async def publish_project(
    user_id: str,
    document: Path,
    inventory: AnyInventory,
) -> str:

    project_id = token_hex(8)
    await _store(
        _PROJECT_KEY.format(user_id=user_id, project_id=project_id),
        document, inventory,
    )

    return project_id


@asynccontextmanager
async def load_project(
    user_id: str,
    project_id: str,
) -> AsyncGenerator[Project]:

    key = _PROJECT_KEY.format(user_id=user_id, project_id=project_id)
    # valkey-py types the async commands like the sync ones.
    fields: dict[bytes, bytes] = await _valkey.hgetall(key)  # pyright: ignore[reportGeneralTypeIssues]

    if not fields:
        raise ToolError(
            f"Project '{project_id}' not found. Start one with `start_project`."
        )

    # Every access counts as activity.
    await _valkey.expire(key, _settings.project_ttl)

    module = FORMATS[fields[b"format"].decode()]

    with TemporaryDirectory() as scratch:

        document = Path(scratch) / f"document.{module.FORMAT}"
        await to_thread(document.write_bytes, fields[b"document"])

        yield Project(
            key=key,
            document=document,
            module=module,
            inventory=module.INVENTORY.model_validate_json(fields[b"inventory"]),
        )


@asynccontextmanager
async def lock_project(
    user_id: str,
    project_id: str,
) -> AsyncGenerator[Lock]:

    lock = _valkey.lock(
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


async def apply_batch(
    project: Project,
    batch: list[dict],
    lock: Lock,
) -> list[dict]:

    if all(command["command"] in READ_COMMANDS for command in batch):
        return await run_batch(project.document, batch)

    if isinstance(project.inventory, PptxInventory):
        await to_thread(pptx.disambiguate_layouts, project.document, batch)

    results = await run_batch(project.document, batch)

    if isinstance(project.inventory, PptxInventory):
        await to_thread(pptx.check, project.document, project.inventory)

    # All fields, so an evicted project comes back whole with this edit.
    await _store(project.key, project.document, project.inventory, lock)

    return results


@asynccontextmanager
async def project_lifespan(
    server: FastMCP,
) -> AsyncGenerator[None]:

    # Below ABI 4 (Linux 6.7) Landlock cannot restrict the network.
    if LANDLOCK_ABI < 4:
        raise RuntimeError("Landlock ABI 4 or newer is needed to confine OfficeCLI.")

    await _valkey.ping()
    await prepare_templates()

    try:
        yield
    finally:
        await _valkey.aclose()
