from asyncio import Semaphore, create_subprocess_exec, timeout, to_thread
from asyncio.subprocess import PIPE
from contextlib import suppress
from ctypes import CDLL
from json import JSONDecodeError, dumps, loads
from os import killpg
from pathlib import Path
from signal import SIGKILL
from tempfile import TemporaryDirectory
from time import perf_counter

from fastmcp.exceptions import ToolError
from fastmcp.utilities.logging import get_logger
from pydantic import JsonValue

from config import get_settings
from models.inventory import Theme

_settings = get_settings()
_logger = get_logger("office")

# Both fetched by bin/download.sh, pinned by version and checksum there.
_LANDRUN = "/usr/local/bin/landrun"
_OFFICECLI = "/usr/local/bin/officecli"

# landlock_create_ruleset(NULL, 0, LANDLOCK_CREATE_RULESET_VERSION) answers the ABI.
LANDLOCK_ABI = CDLL(None, use_errno=True).syscall(444, None, 0, 1)

_ENVIRONMENT = {
    "PATH": "/usr/bin:/bin",
    "LANG": "C.UTF-8",
    # The .NET runtime inside OfficeCLI needs ICU unless it runs invariant.
    "DOTNET_SYSTEM_GLOBALIZATION_INVARIANT": "1",
    # A runaway document kills its own run, not the server.
    "DOTNET_GCHeapHardLimit": f"{_settings.officecli_max_memory * 2**20:#x}",
    "OFFICECLI_NO_AUTO_RESIDENT": "1",
    "OFFICECLI_NO_AUTO_INSTALL": "1",
    "OFFICECLI_SKIP_UPDATE": "1",
}

_MAX_SPILL_BYTES = 64 * 2**20

# The rest of a reference documents OfficeCLI's own sources.
_REFERENCE_KEYS = {"element", "parent", "operations", "paths", "children"}
_PROPERTY_KEYS = {"type", "values", "description", "examples", "add", "set"}

_semaphore = Semaphore(_settings.officecli_max_processes)


async def _execute(
    *arguments: str,
    cwd: Path,
    temp: Path,
    stdin: bytes = b"",
) -> tuple[int | None, str, str]:

    try:
        async with timeout(_settings.officecli_queue_timeout):
            await _semaphore.acquire()
    except TimeoutError:
        raise ToolError(
            "The server is busy; nothing was changed. Retry in a moment."
        ) from None

    started = perf_counter()

    try:

        process = await create_subprocess_exec(
            _LANDRUN, "--best-effort",
            # /etc/fonts, or screenshots fall back to one font.
            "--rox", "/usr", "--ro", "/proc,/dev,/etc/fonts",
            "--rw", f"{cwd},{temp}",
            *(
                f"--env={name}={value}"
                for name, value in {
                    **_ENVIRONMENT, "HOME": str(temp), "TMPDIR": str(temp),
                }.items()
            ),
            "--", _OFFICECLI, *arguments,
            cwd=cwd,
            env={},
            stdin=PIPE,
            stdout=PIPE,
            stderr=PIPE,
            # Own process group, so a kill also takes its children.
            start_new_session=True,
        )

        # OfficeCLI keeps stdout small and spills large output into TMPDIR.
        try:
            async with timeout(_settings.officecli_timeout):
                stdout, stderr = await process.communicate(stdin)
        except BaseException:
            with suppress(ProcessLookupError):
                killpg(process.pid, SIGKILL)
            # Drains the pipes, or wait() would hang on a paused reader.
            await process.communicate()
            raise

    except TimeoutError:
        raise ToolError(
            f"OfficeCLI timed out after {_settings.officecli_timeout:g} s."
        ) from None

    finally:
        _semaphore.release()
        _logger.info(
            "officecli %s: %.2f s", arguments[0], perf_counter() - started,
        )

    return (
        process.returncode,
        stdout.decode(errors="replace"),
        stderr.decode(errors="replace"),
    )


async def run(
    file: Path,
    commands: list[dict],
) -> list[dict]:

    with TemporaryDirectory(prefix="oc-") as temp:

        _, stdout, stderr = await _execute(
            "batch", file.name, "--input", "-", "--json",
            # Relative sources resolve inside the project.
            cwd=file.parent,
            temp=Path(temp),
            stdin=dumps(commands).encode(),
        )

        try:
            envelope = loads(stdout)
        except JSONDecodeError:
            raise RuntimeError(
                f"OfficeCLI failed. Detail: {stderr.strip()[-500:]}"
            ) from None

        if error := envelope.get("error"):
            raise ToolError(f"{error['error']} ({error['code']})")

        data = envelope["data"]

        # OfficeCLI spills large output into its TMPDIR.
        if output_file := data.get("outputFile"):

            spill = Path(output_file)

            if not spill.is_relative_to(temp):
                raise RuntimeError(
                    "OfficeCLI spilled its output to an unexpected path."
                )

            if spill.stat().st_size > _MAX_SPILL_BYTES:
                raise ToolError(
                    "Output too large. Narrow the read: a more specific path "
                    "or selector, or a smaller depth."
                )

            data = await to_thread(lambda: loads(spill.read_bytes()))

    results = data["results"]

    if failures := [result for result in results if not result["success"]]:
        raise ToolError(
            "\n".join(
                f"Command {failure['index']} failed "
                f"({failure.get('code', 'error')}): {failure['error']}"
                for failure in failures
            )
            + "\nThe batch was rolled back, nothing changed. Fix it and "
            "send the whole batch again."
        )

    return results


async def findings(
    file: Path,
) -> list[str]:

    validation, issues = await run(file, [
        {"command": "validate"},
        {"command": "view", "mode": "issues"},
    ])

    # Batch validation is text: a headline, then one indented block per error.
    errors = validation["output"].split("\n  [")[1:]

    return [
        *(" ".join(f"[{error}".split()) for error in errors),
        *(
            f"{issue['path']}: {issue['message']}"
            for issue in issues["output"]["issues"]
        ),
    ]


async def screenshot(
    file: Path,
    page: int | None,
) -> bytes:

    with TemporaryDirectory(prefix="oc-") as temp:

        image = Path(temp) / "page.png"
        returncode, stdout, stderr = await _execute(
            "view", file.name, "screenshot",
            # Without a page, all slides or pages as one contact sheet.
            *(("--page", str(page)) if page else ("--grid",)),
            "-o", str(image),
            cwd=file.parent,
            temp=Path(temp),
        )

        if returncode:
            detail = (stderr or stdout).strip().removeprefix("Error: ")
            raise ToolError(
                f"Could not render {f'page {page}' if page else 'the overview'}: "
                f"{detail[-300:]}"
            )

        return await to_thread(image.read_bytes)


async def help(
    format: str,
    topic: str,
    props: set[str],
) -> JsonValue:

    with TemporaryDirectory(prefix="oc-") as temp:
        returncode, stdout, stderr = await _execute(
            "help", format, *topic.split(), "--json",
            cwd=Path(temp),
            temp=Path(temp),
        )

    if returncode:
        detail = (stderr or stdout).strip()
        # The JSON error envelope may suggest a topic ("Did you mean: chart?").
        with suppress(ValueError, KeyError, TypeError):
            detail = loads(detail)["error"]["error"]
        raise ToolError(
            f"No reference for '{topic}'. {detail.partition('\nUse:')[0]}"
        )

    try:
        reference = loads(stdout)
    except JSONDecodeError:
        # Verb overviews are plain text.
        return stdout.strip()

    properties: dict[str, JsonValue] = {}

    for name, prop in reference.get("properties", {}).items():

        # Listed under the name `props` accepts, which may be an alias.
        if accepted := [
            alias for alias in (name, *prop.get("aliases", []))
            if alias.lower() in props
        ]:
            properties.setdefault(accepted[0], {
                key: value for key, value in prop.items() if key in _PROPERTY_KEYS
            })

    return {
        **{key: value for key, value in reference.items() if key in _REFERENCE_KEYS},
        "properties": properties,
    }


def theme(
    root: dict,
) -> Theme:

    properties = root["output"]["results"][0]["format"]

    return Theme(
        colors={
            key.removeprefix("theme.color."): value
            for key, value in properties.items()
            if key.startswith("theme.color.")
        },
        fonts=list(dict.fromkeys(
            font for key in ("theme.font.major.latin", "theme.font.minor.latin")
            if (font := properties.get(key))
        )),
    )
