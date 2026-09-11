from asyncio import to_thread
from dataclasses import dataclass
from pathlib import Path
from shutil import copyfile, rmtree
from types import ModuleType

from fastmcp.exceptions import ToolError
from fastmcp.utilities.logging import get_logger

from models.inventory import Inventory
from office import _officecli
from office._formats import FORMATS, detect_format

_logger = get_logger("office")

_TEMPLATES = Path("templates")
_PREPARED = Path("data/templates")


@dataclass(frozen=True)
class Prepared:
    module: ModuleType
    document: Path
    inventory: Inventory
    baseline: list[str]


stored: dict[str, Prepared] = {}


async def start() -> None:

    # Prepared once per start; added or changed templates need a restart.
    await to_thread(rmtree, _PREPARED, ignore_errors=True)
    _PREPARED.mkdir(parents=True)

    for source in sorted(_TEMPLATES.iterdir()):

        if source.suffix.removeprefix(".") not in FORMATS:
            continue

        try:
            module = await to_thread(detect_format, source)
            # OfficeCLI picks its handler by suffix.
            document = _PREPARED / f"{source.stem}.{module.FORMAT}"
            await to_thread(copyfile, source, document)
            await module.prepare(document)
            inventory = await module.inventory(document)
            baseline = await _officecli.findings(document)
        except (ToolError, RuntimeError) as error:
            _logger.warning("template %s skipped: %s", source.name, error)
            continue

        stored[source.name] = Prepared(
            module=module,
            document=document,
            inventory=inventory,
            baseline=baseline,
        )
