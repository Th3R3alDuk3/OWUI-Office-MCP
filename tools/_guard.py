import re
from collections.abc import Awaitable, Callable
from json import dumps

from fastmcp.exceptions import ToolError

from models.office import AnyInventory, Command, XlsxInventory
from services import xlsx

# Whole files and the process, not a project.
_FILE_COMMANDS = {"create", "open", "close", "save", "merge", "watch", "mcp", "install"}

_SOURCE_PROPS = {"src", "image"}
_FILE_REFERENCE = re.compile(r"file:([A-Za-z0-9-]{1,64})")

_MAX_BATCH_BYTES = 2**20


async def check_commands(
    commands: list[Command],
    inventory: AnyInventory,
    fetch: Callable[[str], Awaitable[str]],
) -> list[dict]:

    batch = [command.model_dump() for command in commands]

    if len(dumps(batch)) > _MAX_BATCH_BYTES:
        raise ToolError(
            f"The batch exceeds {_MAX_BATCH_BYTES // 2**20} MB; split it."
        )

    checked: list[dict] = []

    for index, command in enumerate(batch):

        # OfficeCLI treats verbs and prop names case-insensitively.
        verb = command["command"] = str(command["command"]).lower()
        props = command.get("props", {})

        try:

            if not isinstance(props, dict):
                raise ToolError("`props` must be an object of prop name -> value.")

            props = command["props"] = {
                str(key).lower(): value for key, value in props.items()
            }

            if verb in _FILE_COMMANDS:
                raise ToolError(f"`{verb}` works on files, not on a project.")

            for key in _SOURCE_PROPS & props.keys():
                if not _FILE_REFERENCE.fullmatch(str(props[key])):
                    raise ToolError(
                        f"Prop `{key}` only takes `file:<file_id>` of an image "
                        "the user attached."
                    )

            # Named cell styles are the one thing the engine has no prop for.
            if isinstance(inventory, XlsxInventory):
                command = xlsx.resolve_style(command, inventory)

            checked.append(command)

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
