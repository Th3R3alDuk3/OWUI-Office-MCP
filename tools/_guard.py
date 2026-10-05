import re
from collections.abc import Awaitable, Callable
from json import dumps

from fastmcp.exceptions import ToolError

from models.office import Command

_SOURCE_PROPS = {"src", "image"}
_FILE_REFERENCE = re.compile(r"file:[A-Za-z0-9-]{1,64}")

_MAX_BATCH_BYTES = 2**20
_MAX_FILES = 10


async def check_commands(
    commands: list[Command],
    fetch: Callable[[str], Awaitable[str]],
) -> list[dict]:

    batch = [command.model_dump() for command in commands]

    if len(dumps(batch)) > _MAX_BATCH_BYTES:
        raise ToolError(
            f"The batch exceeds {_MAX_BATCH_BYTES // 2**20} MB; split it."
        )

    for index, command in enumerate(batch):

        props = command.get("props", {})

        try:

            if not isinstance(props, dict):
                raise ToolError("`props` must be an object of prop name -> value.")

            # OfficeCLI reads prop names case-insensitively.
            props = command["props"] = {
                str(key).lower(): value for key, value in props.items()
            }

            for key in _SOURCE_PROPS & props.keys():
                if not _FILE_REFERENCE.fullmatch(str(props[key])):
                    raise ToolError(
                        f"Prop `{key}` only takes `file:<file_id>` of an image "
                        "the user attached."
                    )

        except ToolError as error:
            raise ToolError(f"Command {index}: {error}") from None

    file_ids = {
        str(command["props"][key]).removeprefix("file:")
        for command in batch
        for key in _SOURCE_PROPS & command["props"].keys()
    }

    if len(file_ids) > _MAX_FILES:
        raise ToolError(
            f"The batch uses over {_MAX_FILES} attached files; split it."
        )

    # Invalid later commands must not trigger downloads for earlier ones.
    fetched = {file_id: await fetch(file_id) for file_id in file_ids}

    for command in batch:
        props = command["props"]
        for key in _SOURCE_PROPS & props.keys():
            props[key] = fetched[str(props[key]).removeprefix("file:")]

    return batch
