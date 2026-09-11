import re
from collections.abc import Awaitable, Callable
from json import dumps
from types import ModuleType

from fastmcp.exceptions import ToolError

from models.guard import Command
from models.inventory import Inventory
from models.project import DesignMode

READ_COMMANDS = {"get", "query", "view", "validate"}

_TARGET_FIELDS = ("path", "parent", "to", "path2")

# Attached images are fetched from OpenWebUI into the project's assets.
_SOURCE_PROPS = {"src", "image"}
_FILE_REFERENCE = re.compile(r"file:([A-Za-z0-9-]{1,64})")

_MAX_BATCH_BYTES = 2**20


async def check(
    commands: list[Command],
    module: ModuleType,
    inventory: Inventory,
    design_mode: DesignMode,
    fetch: Callable[[str], Awaitable[str]],
) -> list[dict]:

    batch = [
        command.model_dump(by_alias=True, exclude_none=True)
        for command in commands
    ]

    if len(dumps(batch)) > _MAX_BATCH_BYTES:
        raise ToolError(
            f"The batch exceeds {_MAX_BATCH_BYTES // 2**20} MB; split it."
        )

    allowed_props = module.CONTENT_PROPS | (
        module.APPEARANCE_PROPS if design_mode == "custom" else set()
    )
    checked: list[dict] = []

    for index, command in enumerate(batch):

        verb = command["command"]
        props = command.get("props", {})

        try:

            if verb not in READ_COMMANDS and any(
                command.get(field, "").lower().startswith(module.PROTECTED_PATHS)
                for field in _TARGET_FIELDS
            ):
                raise ToolError(
                    "Master, layout and style definitions are read-only."
                )

            if verb == "add" and command.get("type") not in module.ELEMENT_TYPES:
                raise ToolError(
                    f"Type `{command.get('type')}` is not supported. Use one "
                    f"of: {', '.join(sorted(module.ELEMENT_TYPES))}."
                )

            for key, value in props.items():

                if key not in allowed_props:
                    raise ToolError(
                        f"Prop `{key}` is not supported"
                        + (
                            " in template mode; appearance props need "
                            '`design_mode="custom"` and an explicit user '
                            "request."
                            if key in module.APPEARANCE_PROPS else "."
                        )
                    )

                if key in _SOURCE_PROPS:
                    if not _FILE_REFERENCE.fullmatch(str(value)):
                        raise ToolError(
                            f"Prop `{key}` only takes `file:<file_id>` of an "
                            "image the user attached."
                        )

            checked.append(module.adapt(command, inventory))

        except ToolError as error:
            raise ToolError(f"Command {index}: {error}") from None

    # Invalid later commands must not trigger downloads for earlier ones.
    for index, command in enumerate(checked):

        props = command.get("props", {})

        try:
            for key in _SOURCE_PROPS & props.keys():
                props[key] = await fetch(str(props[key])[5:])
        except ToolError as error:
            raise ToolError(f"Command {index}: {error}") from None

    return checked
