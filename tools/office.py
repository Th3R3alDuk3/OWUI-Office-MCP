from asyncio import to_thread
from collections import Counter
from collections.abc import AsyncGenerator
from contextlib import asynccontextmanager
from functools import partial
from pathlib import Path
from tempfile import TemporaryDirectory

from fastmcp.dependencies import CurrentAccessToken, TokenClaim
from fastmcp.exceptions import ToolError
from fastmcp.server.auth import AccessToken
from fastmcp.tools import ToolResult, tool
from fastmcp.utilities.logging import get_logger
from fastmcp.utilities.types import File, Image
from mcp.types import TextContent, ToolAnnotations
from pydantic import Field

from config import get_settings
from models.office import (
    Commands,
    CommandsResult,
    ExportResult,
    Format,
    PptxInventory,
    ReferenceResult,
    StartResult,
    Template,
    TemplatesResult,
)
from services import pptx
from services.officecli import lookup_reference, screenshot
from services.owui import upload_file
from services.project import (
    add_asset,
    apply_batch,
    fetch_upload,
    load_project,
    lock_project,
    publish_project,
    templates,
)
from tools._guard import check_commands

_settings = get_settings()
logger = get_logger(__name__)

_active: Counter[str] = Counter()


@asynccontextmanager
async def _admitted(
    user_id: str,
) -> AsyncGenerator[None]:

    if (
        _active.total() >= _settings.max_concurrent_requests
        or _active[user_id] >= _settings.max_concurrent_requests_per_user
    ):
        raise ToolError(
            "The server is busy; nothing was changed. Retry in a moment."
        )

    _active[user_id] += 1

    try:
        yield
    except ToolError as error:
        # What the model got wrong.
        logger.warning("user %s: %s", user_id, error)
        raise
    finally:

        _active[user_id] -= 1

        if not _active[user_id]:
            del _active[user_id]


@tool(
    name="list_templates",
    description=(
        "List the stored templates with their format and, for PPTX, their "
        "slide masters. Use this when the user did not attach a file, then "
        "pick one for `start_project`."
    ),
    annotations=ToolAnnotations(read_only_hint=True, open_world_hint=False),
)
async def list_templates() -> TemplatesResult:

    listed: list[Template] = []

    for name, stored in templates.items():

        inventory = stored.inventory
        listed.append(Template(
            name=name,
            format=stored.module.FORMAT,
            slide_masters=(
                {master.index: master.name for master in inventory.slide_masters}
                if isinstance(inventory, PptxInventory) else None
            ),
        ))

    return TemplatesResult(
        hint=(
            "Pick a template of the format the user wants and call "
            "`start_project`; for a PPTX template with several `slide_masters`, "
            "also pass the one the user named. If several could fit and the "
            "user named none, ask the user instead of guessing, template and "
            "slide master in one question."
            if listed else
            "No stored templates. Ask the administrator to add some, or ask "
            "the user to attach a file whose design to use."
        ),
        templates=listed,
    )


@tool(
    name="start_project",
    description=(
        "Start a project from a design: a stored template (`template_name`) or "
        "an attached file (`file_id`), exactly one. Of an attached file only the "
        "design is used and its slides or body text are dropped, unless "
        "`keep_content` is true; the attachment itself never changes. A PPTX "
        "project is bound to one slide master: with several, pass "
        "`slide_master`; without it the call lists them."
    ),
    annotations=ToolAnnotations(destructive_hint=False, open_world_hint=True),
)
async def start_project(
    template_name: str | None = Field(
        default=None,
        description="Stored template from `list_templates`.",
    ),
    file_id: str | None = Field(
        default=None,
        pattern=r"^[A-Za-z0-9-]{1,64}$",
        description=(
            "OpenWebUI file ID of the attached file; never a template name, "
            "never invented."
        ),
    ),
    keep_content: bool = Field(
        default=False,
        description=(
            "Edit the attached file as it is instead of using only its "
            "design. Workbooks keep their content either way."
        ),
    ),
    slide_master: int | None = Field(
        default=None,
        ge=1,
        description=(
            "PPTX only: index of the slide master the project is bound to. "
            "Required when the design has several slide masters."
        ),
    ),
    token: AccessToken = CurrentAccessToken(),
    user_id: str = TokenClaim("id"),
) -> StartResult:

    async with _admitted(user_id):

        with TemporaryDirectory() as scratch:

            directory = Path(scratch)

            match template_name, file_id:

                case str(), None:
                    if (stored := templates.get(template_name)) is None:
                        raise ToolError(
                            f"Template '{template_name}' not found. Pick one "
                            "from `list_templates`."
                        )
                    if keep_content:
                        raise ToolError(
                            "`keep_content` applies to attached files only; "
                            "leave it out for templates."
                        )
                    module, inventory = stored.module, stored.inventory
                    document = directory / f"document.{module.FORMAT}"
                    await to_thread(document.write_bytes, stored.document)
                    origin = f"template '{template_name}'"

                case None, str():
                    module, document = await fetch_upload(
                        directory, file_id, token.token,
                    )
                    if not keep_content:
                        await module.prepare(document)
                    inventory = await module.inventory(document)
                    origin = (
                        f"attached file '{file_id}'" if keep_content
                        else f"the design of attached file '{file_id}'"
                    )

                case _:
                    raise ToolError("Pass either `template_name` or `file_id`.")

            if isinstance(inventory, PptxInventory):
                inventory = pptx.bind(inventory, slide_master)
                # Only kept slides can stray from the bound slide master.
                if keep_content:
                    await pptx.check(document, inventory)
            elif slide_master is not None:
                raise ToolError("`slide_master` applies to PPTX designs only.")

            project_id = await publish_project(user_id, document, inventory)

    return StartResult(
        hint=(
            f"Project started from {origin}. "
            + (
                'Read its content first, e.g. {"command": "view", "mode": '
                '"outline"}. ' if keep_content else ""
            )
            + module.START_HINT
        ),
        project_id=project_id,
        format=module.FORMAT,
        inventory=inventory,
    )


@tool(
    name="get_reference",
    description=(
        "Look up the OfficeCLI reference of a format: an element (e.g. "
        "`chart`, `placeholder`, `table`), a verb (e.g. `add`) or both (e.g. "
        "`add picture`). An element's reference lists its paths, props and "
        "aliases."
    ),
    annotations=ToolAnnotations(read_only_hint=True, open_world_hint=False),
)
async def get_reference(
    format: Format = Field(
        description="Document format.",
    ),
    topic: str = Field(
        pattern=r"^[a-z][a-z0-9-]*( [a-z][a-z0-9-]*)?$",
        description="Element and/or verb, e.g. `chart` or `add picture`.",
    ),
    user_id: str = TokenClaim("id"),
) -> ReferenceResult:

    async with _admitted(user_id):
        reference = await lookup_reference(format, topic)

    return ReferenceResult(
        hint="Use the props as documented.",
        reference=reference,
    )


@tool(
    name="run_commands",
    description=(
        "Apply one atomic batch of OfficeCLI commands: if one fails, nothing is "
        "applied and its error comes back with the command index. Each command "
        "names its verb in `command` as in an OfficeCLI batch; reads (get, "
        "query, view, validate, raw) return their output in `results`. Send a "
        "whole step as one batch, not command by command."
    ),
    annotations=ToolAnnotations(destructive_hint=True, open_world_hint=True),
)
async def run_commands(
    project_id: str = Field(
        pattern=r"^[0-9a-f]{16}$",
        description="From `start_project`.",
    ),
    commands: Commands = Field(
        min_length=1,
        max_length=500,
        description=(
            'OfficeCLI commands, e.g. [{"command": "add", "parent": "/", '
            '"type": "slide", "props": {"layout": "3"}}].'
        ),
    ),
    token: AccessToken = CurrentAccessToken(),
    user_id: str = TokenClaim("id"),
) -> CommandsResult:

    async with (
        _admitted(user_id),
        lock_project(user_id, project_id),
        load_project(user_id, project_id) as project,
    ):
        batch = await check_commands(
            commands,
            fetch=partial(add_asset, project, token=token.token),
        )
        results = await apply_batch(project, batch)

    return CommandsResult(
        hint=(
            "Continue; once the user's request is fully applied, call "
            "`export_project` once."
        ),
        results=[result.get("output") for result in results],
        warnings=[
            f"Command {result['index']}: {warning['message']}"
            for result in results for warning in result.get("warnings", [])
        ],
    )


@tool(
    name="preview_project",
    description=(
        "Render the project as an image: the whole document as thumbnails, or "
        "one slide (PPTX) or page (DOCX) in full size; XLSX shows the first "
        "sheet. "
        + (
            "Check it for overflowing or overlapping content, empty "
            "placeholders and broken layouts. Use it after building or larger "
            "edits, not after every change."
            if _settings.preview_to_model else
            "The image is shown to the user, not to you; use it when the user "
            "wants to see the state, then ask what to change."
        )
    ),
    annotations=ToolAnnotations(read_only_hint=True, open_world_hint=False),
)
async def preview_project(
    project_id: str = Field(
        pattern=r"^[0-9a-f]{16}$",
        description="From `start_project`.",
    ),
    page: int | None = Field(
        default=None,
        ge=1,
        description=(
            "Slide or page number for one page in full size; omit for all "
            "slides or pages as thumbnails."
        ),
    ),
    user_id: str = TokenClaim("id"),
) -> ToolResult:

    async with _admitted(user_id), load_project(user_id, project_id) as project:
        image = await screenshot(project.document, page)

    subject = f"Page {page}" if page else "The whole document"

    # OpenWebUI shows image content to the user, image resources to the model.
    return ToolResult(content=[
        TextContent(
            type="text",
            text=(
                f"{subject} is attached as an image and shown to the user. "
                "It approximates Office's rendering: judge layout, overflow "
                "and missing or unreadable content, not exact colors or pixel "
                "details. Fix what is wrong with `run_commands`, then call "
                "`export_project`."
                if _settings.preview_to_model else
                f"{subject} is shown to the user as an image. Ask the user "
                "what to change, fix it with `run_commands`, then call "
                "`export_project`."
            ),
        ),
        *(
            [File(data=image, name=f"page-{page or 'all'}.png").to_resource_content(
                mime_type="image/png",
            )] if _settings.preview_to_model else []
        ),
        Image(data=image).to_image_content(),
    ])


@tool(
    name="export_project",
    description=(
        "Upload the project's document to OpenWebUI as the user's download. "
        "Call it once the user's request is fully applied, not after each "
        "change; run `validate` first if in doubt."
    ),
    annotations=ToolAnnotations(destructive_hint=False, open_world_hint=True),
)
async def export_project(
    project_id: str = Field(
        pattern=r"^[0-9a-f]{16}$",
        description="From `start_project`.",
    ),
    file_name: str = Field(
        min_length=1,
        max_length=60,
        pattern=r'^[^/\\:*?"<>|]+$',
        description="Stem without extension.",
    ),
    token: AccessToken = CurrentAccessToken(),
    user_id: str = TokenClaim("id"),
) -> ExportResult:

    async with _admitted(user_id), load_project(user_id, project_id) as project:

        upload_name = f"{file_name}.{project.module.FORMAT}"

        try:
            owui_url = await upload_file(
                file_name=upload_name,
                data=await to_thread(project.document.read_bytes),
                content_type=project.module.MIME,
                token=token.token,
            )
        except RuntimeError as error:
            raise ToolError(
                "Could not upload the file to OpenWebUI. Retry in a moment."
            ) from error

    return ExportResult(
        hint=(
            "Share `owui_url` with the user as the download link. The project "
            "stays editable; export it again after later edits."
        ),
        file_name=upload_name,
        owui_url=owui_url,
    )
