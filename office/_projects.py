from asyncio import Lock, sleep, timeout, to_thread
from collections.abc import AsyncGenerator
from contextlib import asynccontextmanager, suppress
from dataclasses import dataclass
from hashlib import sha256
from json import dumps, loads
from mimetypes import guess_extension
from os import utime
from pathlib import Path
from secrets import token_hex
from shutil import copyfile, rmtree
from time import time
from types import ModuleType
from weakref import WeakValueDictionary

from fastmcp.exceptions import ToolError

from config import get_settings
from models.inventory import Inventory, PptxInventory
from office import _guard, _officecli, pptx
from office._formats import FORMATS, detect_format
from services.owui import download_file

_settings = get_settings()

_DATA = Path("data")
_USERS = _DATA / "users"
_INVENTORY_FILE = "inventory.json"
_BASELINE_FILE = "baseline.json"
_ASSETS_DIR = "assets"

_MAX_UPLOAD_BYTES = 100 * 2**20
_MAX_ASSET_BYTES = 20 * 2**20
_MAX_USER_BYTES = _settings.data_max_size_per_user * 10**6
_MAX_DATA_BYTES = _settings.data_max_size * 10**6

_IMAGE_TYPES = {
    "image/png", "image/jpeg", "image/gif", "image/bmp", "image/tiff",
    "image/webp",
}

_locks: WeakValueDictionary[str, Lock] = WeakValueDictionary()


@dataclass(frozen=True)
class Project:
    directory: Path
    module: ModuleType
    inventory: Inventory
    baseline: list[str]

    @property
    def document(self) -> Path:
        return self.directory / f"document.{self.module.FORMAT}"


def _user_directory(
    user_id: str,
) -> Path:

    return _USERS / sha256(user_id.encode()).hexdigest()


def _size(
    directory: Path,
) -> int:

    size = 0

    for file in directory.rglob("*"):
        # Engine temporary files may vanish mid-walk.
        with suppress(OSError):
            size += file.stat().st_size if file.is_file() else 0

    return size


def _check_quota(
    user_directory: Path,
) -> None:

    # Checked before each write, so one write may overshoot by its own size.
    if _size(user_directory) > _MAX_USER_BYTES:
        raise ToolError(
            "Your storage quota is used up. Old projects expire after "
            "inactivity; retry later."
        )

    if _size(_DATA) > _MAX_DATA_BYTES:
        raise ToolError("Server storage is full. Retry later.")


@asynccontextmanager
async def create(
    user_id: str,
) -> AsyncGenerator[Path]:

    user_directory = _user_directory(user_id)
    user_directory.mkdir(parents=True, exist_ok=True)

    if len(list(user_directory.iterdir())) >= _settings.project_max_per_user:
        raise ToolError(
            f"You have {_settings.project_max_per_user} projects, the "
            "maximum. Old projects expire after inactivity; retry later."
        )

    # Created before the next await, so a parallel creation counts it.
    directory = user_directory / token_hex(8)
    directory.mkdir()

    try:
        await to_thread(_check_quota, user_directory)
        yield directory
    except BaseException:
        await to_thread(rmtree, directory, ignore_errors=True)
        raise


async def _download(
    file_id: str,
    token: str,
    target: Path,
    max_bytes: int,
) -> str:

    try:
        content_type = await download_file(
            file_id=file_id,
            token=token,
            target=target,
            max_bytes=max_bytes,
        )
    except RuntimeError as error:
        raise ToolError(
            f"Could not fetch file '{file_id}' from OpenWebUI. Check that the "
            "ID belongs to a file the user actually attached."
        ) from error

    if target.stat().st_size > max_bytes:
        raise ToolError(
            f"File '{file_id}' is too large. Limit is {max_bytes // 2**20} MB."
        )

    return content_type


async def receive(
    directory: Path,
    file_id: str,
    token: str,
) -> ModuleType:

    download = directory / "upload"
    await _download(file_id, token, download, _MAX_UPLOAD_BYTES)

    module = await to_thread(detect_format, download)
    # OfficeCLI picks its handler by suffix.
    download.rename(directory / f"document.{module.FORMAT}")

    return module


def publish(
    directory: Path,
    inventory: Inventory,
    baseline: list[str],
) -> None:

    directory.joinpath(_BASELINE_FILE).write_text(dumps(baseline))

    # A project exists once its inventory does.
    staging = directory / f".{_INVENTORY_FILE}"
    staging.write_text(inventory.model_dump_json())
    staging.replace(directory / _INVENTORY_FILE)


def get(
    user_id: str,
    project_id: str,
) -> Project:

    directory = _user_directory(user_id) / project_id

    if not directory.joinpath(_INVENTORY_FILE).exists():
        raise ToolError(
            f"Project '{project_id}' not found. Start one with "
            "`create_project` or `open_project`."
        )

    # Every access counts as activity for the sweep.
    utime(directory)

    document = next(directory.glob("document.*"))
    module = FORMATS[document.suffix.removeprefix(".")]

    return Project(
        directory=directory,
        module=module,
        inventory=module.INVENTORY.model_validate_json(
            directory.joinpath(_INVENTORY_FILE).read_bytes()
        ),
        baseline=loads(directory.joinpath(_BASELINE_FILE).read_bytes()),
    )


@asynccontextmanager
async def locked(
    project: Project,
) -> AsyncGenerator[None]:

    # Weak entries vanish once no request holds or awaits the lock.
    lock = _locks.setdefault(str(project.directory), Lock())

    try:
        async with timeout(_settings.officecli_queue_timeout):
            await lock.acquire()
    except TimeoutError:
        raise ToolError(
            "The project is busy; nothing was changed. Retry in a moment."
        ) from None

    try:
        yield
    finally:
        lock.release()


async def add_asset(
    project: Project,
    file_id: str,
    token: str,
) -> str:

    assets = project.directory / _ASSETS_DIR

    if existing := next(assets.glob(f"{file_id}.*"), None):
        return f"{_ASSETS_DIR}/{existing.name}"

    await to_thread(_check_quota, project.directory.parent)

    assets.mkdir(exist_ok=True)
    download = assets / f".{file_id}"

    try:

        content_type = await _download(
            file_id, token, download, _MAX_ASSET_BYTES,
        )

        if content_type not in _IMAGE_TYPES:
            raise ToolError(
                f"File '{file_id}' is no supported image "
                "(PNG, JPEG, GIF, BMP, TIFF, WebP)."
            )

        name = f"{file_id}{guess_extension(content_type)}"
        download.rename(assets / name)

    finally:
        download.unlink(missing_ok=True)

    return f"{_ASSETS_DIR}/{name}"


async def apply(
    project: Project,
    batch: list[dict],
) -> list[dict]:

    if all(command["command"] in _guard.READ_COMMANDS for command in batch):
        return await _officecli.run(project.document, batch)

    await to_thread(_check_quota, project.directory.parent)

    # Edits go live only after they completed.
    work = project.directory / f"work.{project.module.FORMAT}"
    await to_thread(copyfile, project.document, work)

    try:
        if isinstance(project.inventory, PptxInventory):
            await to_thread(pptx.check, work, project.inventory, batch)
        results = await _officecli.run(work, batch)
        if isinstance(project.inventory, PptxInventory):
            await to_thread(pptx.check, work, project.inventory)
        work.replace(project.document)
    finally:
        work.unlink(missing_ok=True)

    return results


async def sweep() -> None:

    while True:

        await sleep(_settings.project_sweep_interval)
        deadline = time() - _settings.project_ttl

        for directory in await to_thread(lambda: list(_USERS.glob("*/*"))):

            lock = _locks.setdefault(str(directory), Lock())

            # A failed creation may have removed the directory since the glob.
            with suppress(OSError):
                if not lock.locked() and directory.stat().st_mtime < deadline:
                    async with lock:
                        await to_thread(rmtree, directory, ignore_errors=True)


def start() -> None:

    for directory in _USERS.glob("*/*"):

        # Unpublished projects come from creations interrupted by a restart.
        if not directory.joinpath(_INVENTORY_FILE).exists():
            rmtree(directory, ignore_errors=True)
            continue

        for leftover in (
            *directory.glob("work.*"),
            *directory.glob(".*"),
            *directory.glob(f"{_ASSETS_DIR}/.*"),
        ):
            leftover.unlink(missing_ok=True)
