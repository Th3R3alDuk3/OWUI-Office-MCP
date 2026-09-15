from asyncio import to_thread
from collections import Counter
from collections.abc import AsyncGenerator
from contextlib import asynccontextmanager
from functools import partial
from json import dumps
from pathlib import Path
from tempfile import TemporaryDirectory

from fastmcp import FastMCP
from fastmcp.dependencies import CurrentAccessToken, TokenClaim
from fastmcp.exceptions import ToolError
from fastmcp.server.auth import AccessToken
from fastmcp.server.lifespan import lifespan
from fastmcp.tools import ToolResult, tool
from fastmcp.utilities.types import File, Image
from mcp.types import TextContent, ToolAnnotations
from pydantic import Field, JsonValue

from config import get_settings
from models.inventory import PptxInventory
from models.project import (
    Command,
    CommandsResult,
    DesignMode,
    ExportResult,
    Format,
    ReferenceResult,
    StartResult,
    Template,
    TemplatesResult,
)
from tools._office import guard, officecli, pptx, projects
from services.owui import upload_file

_settings = get_settings()

_MAX_OUTPUT_CHARS = 50_000

_active: Counter[str] = Counter()


@lifespan
async def office_lifespan(
    server: FastMCP,
):

    # Below ABI 4 (Linux 6.7) Landlock cannot restrict the network.
    if officecli.LANDLOCK_ABI < 4:
        raise RuntimeError("Landlock ABI 4 or newer is needed to confine OfficeCLI.")

    await projects.valkey.ping()
    await projects.prepare_templates()

    try:
        yield
    finally:
        await projects.valkey.aclose()


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
    finally:

        _active[user_id] -= 1

        if not _active[user_id]:
            del _active[user_id]


@tool(
    name="list_templates",
    tags={"office", "templates"},
    description=(
        "List the stored templates with their format and, for PPTX, their "
        "slide masters. Use this when the user did NOT attach a file, then "
        "pick one for `create_project`."
    ),
    annotations=ToolAnnotations(read_only_hint=True, open_world_hint=False),
)
async def list_templates() -> TemplatesResult:

    templates: list[Template] = []

    for name, prepared in projects.templates.items():

        inventory = prepared.inventory
        templates.append(Template(
            name=name,
            format=prepared.module.FORMAT,
            masters=(
                {master.index: master.name for master in inventory.masters}
                if isinstance(inventory, PptxInventory) else None
            ),
        ))

    return TemplatesResult(
        hint=(
            "Pick a template of the format the user wants and call "
            "`create_project`; for a PPTX template with several `masters`, "
            "also pass the master the user named. If several could fit and "
            "the user named none, ask the user instead of guessing, template "
            "and master in one question."
            if templates else
            "No stored templates. Ask the administrator to add some, or ask "
            "the user to attach a file whose design to use."
        ),
        templates=templates,
    )


@tool(
    name="create_project",
    tags={"office", "project"},
    description=(
        "Create a new project from a design: a stored template "
        "(`template_name` from `list_templates`) or, if the user attached a "
        "file as the design (e.g. their own slide master), that file "
        "(`file_id`). Pass exactly one of them. The design's masters, layouts "
        "and styles are kept; PPTX and DOCX content is dropped, XLSX keeps its "
        "sheets, values and formulas. A PPTX project is bound "
        "to one slide master: if the design has several, pass the one the "
        "user named as `master`; without it the call lists them. The result "
        "carries the `project_id` and an overview of what the design offers."
    ),
    annotations=ToolAnnotations(destructive_hint=False, open_world_hint=True),
)
async def create_project(
    template_name: str | None = Field(
        default=None,
        description="Stored template from `list_templates`.",
    ),
    file_id: str | None = Field(
        default=None,
        pattern=r"^[A-Za-z0-9-]{1,64}$",
        description=(
            "OpenWebUI file ID of the attached design; never a template name, "
            "never invented."
        ),
    ),
    master: int | None = Field(
        default=None,
        ge=1,
        description=(
            "PPTX only: index of the slide master the project is bound to. "
            "Required when the design has several masters."
        ),
    ),
    token: AccessToken = CurrentAccessToken(),
    user_id: str = TokenClaim("id"),
) -> StartResult:

    async with _admitted(user_id):

        await projects.check_limit(user_id)

        with TemporaryDirectory() as scratch:

            directory = Path(scratch)

            match template_name, file_id:

                case str(), None:
                    if (prepared := projects.templates.get(template_name)) is None:
                        raise ToolError(
                            f"Template '{template_name}' not found. Pick one "
                            "from `list_templates`."
                        )
                    module, inventory, baseline = (
                        prepared.module, prepared.inventory, prepared.baseline,
                    )
                    document = directory / f"document.{module.FORMAT}"
                    await to_thread(document.write_bytes, prepared.document)
                    origin = f"template '{template_name}'"

                case None, str():
                    module = await projects.fetch_upload(directory, file_id, token.token)
                    document = directory / f"document.{module.FORMAT}"
                    await module.prepare(document)
                    inventory = await module.inventory(document)
                    baseline = await officecli.findings(document)
                    origin = f"the design of attached file '{file_id}'"

                case _:
                    raise ToolError("Pass either `template_name` or `file_id`.")

            if isinstance(inventory, PptxInventory):
                inventory = pptx.bind(inventory, master)
                await to_thread(pptx.check, document, inventory)
            elif master is not None:
                raise ToolError("`master` applies to pptx designs only.")
            project_id = await projects.publish(user_id, document, inventory, baseline)

    return StartResult(
        hint=f"Project created from {origin}. {module.START_HINT}",
        project_id=project_id,
        format=module.FORMAT,
        inventory=inventory,
    )


@tool(
    name="open_project",
    tags={"office", "project"},
    description=(
        "Open a `.pptx`, `.docx` or `.xlsx` the user attached in OpenWebUI to "
        "edit it as it is, by its `file_id`. Use only when the user actually "
        "attached a file to edit; to use a file only as the design, call "
        "`create_project` with it instead. The attachment itself stays "
        "unchanged. `master` binds a PPTX project as in `create_project`."
    ),
    annotations=ToolAnnotations(destructive_hint=False, open_world_hint=True),
)
async def open_project(
    file_id: str = Field(
        pattern=r"^[A-Za-z0-9-]{1,64}$",
        description=(
            "OpenWebUI file ID of the attached file; never a template name, "
            "never invented."
        ),
    ),
    master: int | None = Field(
        default=None,
        ge=1,
        description=(
            "PPTX only: index of the slide master the project is bound to. "
            "Required when the file has several masters."
        ),
    ),
    token: AccessToken = CurrentAccessToken(),
    user_id: str = TokenClaim("id"),
) -> StartResult:

    async with _admitted(user_id):

        await projects.check_limit(user_id)

        with TemporaryDirectory() as scratch:

            directory = Path(scratch)
            module = await projects.fetch_upload(directory, file_id, token.token)
            document = directory / f"document.{module.FORMAT}"
            inventory = await module.inventory(document)
            if isinstance(inventory, PptxInventory):
                inventory = pptx.bind(inventory, master)
                await to_thread(pptx.check, document, inventory)
            elif master is not None:
                raise ToolError("`master` applies to pptx designs only.")
            project_id = await projects.publish(
                user_id, document, inventory, await officecli.findings(document),
            )

    return StartResult(
        hint=(
            f"Project opened from attached file '{file_id}'. Read its content "
            'first, e.g. {"command": "view", "mode": "outline"}. '
            f"{module.START_HINT}"
        ),
        project_id=project_id,
        format=module.FORMAT,
        inventory=inventory,
    )


@tool(
    name="run_commands",
    tags={"office", "edit"},
    description=(
        "Apply one batch of OfficeCLI commands to a project, atomically: if a "
        "command fails, nothing is applied. Each command names its verb in "
        "`command` and passes its arguments as sibling fields. Reads (get, "
        "query, view) return their output in `results`. Images come from "
        "attached files as `file:<file_id>`. Only a supported subset of "
        "OfficeCLI is exposed; unsupported input is rejected with the allowed "
        "options. Put all changes of one step into one batch."
    ),
    annotations=ToolAnnotations(destructive_hint=True, open_world_hint=True),
)
async def run_commands(
    project_id: str = Field(
        pattern=r"^[0-9a-f]{16}$",
        description="From `create_project` or `open_project`.",
    ),
    commands: list[Command] = Field(
        min_length=1,
        max_length=500,
        description=(
            'OfficeCLI commands, e.g. [{"command": "add", "parent": "/", '
            '"type": "slide", "props": {"layout": "3"}}].'
        ),
    ),
    design_mode: DesignMode = Field(
        default="template",
        description=(
            '"template" keeps the design\'s layouts, styles and appearance. '
            '"custom" also allows appearance props like colors and fonts; use '
            "it only when the user explicitly asks for such changes."
        ),
    ),
    token: AccessToken = CurrentAccessToken(),
    user_id: str = TokenClaim("id"),
) -> CommandsResult:

    async with (
        _admitted(user_id),
        projects.locked(user_id, project_id) as lock,
        projects.load(user_id, project_id) as project,
    ):
        batch = await guard.check(
            commands,
            module=project.module,
            inventory=project.inventory,
            design_mode=design_mode,
            fetch=partial(projects.add_asset, project, token=token.token),
        )
        results = await projects.apply(project, batch, lock)

    outputs: list[JsonValue] = []

    for result in results:

        output = result.get("output")

        if (size := len(dumps(output))) > _MAX_OUTPUT_CHARS:
            output = (
                f"Output too large ({size} characters). Narrow the read: a "
                "more specific path or selector, or a smaller depth."
            )

        outputs.append(output)

    return CommandsResult(
        hint=(
            "Continue the edit batch; when the user's request is fully "
            "applied, call `export_project` once."
        ),
        results=outputs,
        warnings=[
            f"Command {result['index']}: {warning['message']}"
            for result in results for warning in result.get("warnings", [])
        ],
    )


@tool(
    name="get_reference",
    tags={"office", "reference"},
    description=(
        "Look up the OfficeCLI reference of a format: an element (e.g. "
        "`chart`, `placeholder`, `table`), a verb (e.g. `add`) or both (e.g. "
        "`add picture`). An element's reference lists only the props "
        "`run_commands` accepts."
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

    module = projects.FORMATS[format]

    async with _admitted(user_id):
        reference = await officecli.help(
            format, topic, module.CONTENT_PROPS | module.APPEARANCE_PROPS,
        )

    return ReferenceResult(
        hint=(
            "Appearance props need `design_mode=\"custom\"` and an explicit "
            "user request."
        ),
        reference=reference,
    )


@tool(
    name="preview_project",
    tags={"office", "preview"},
    description=(
        "Render the project as an image to check the result visually: the "
        "whole document as thumbnails, or one slide (PPTX) or page (DOCX) in "
        "full size; XLSX shows the first sheet. Look for overflowing or "
        "overlapping content, empty placeholders and broken layouts. Use it "
        "after building or larger edits, not after every change."
    ),
    annotations=ToolAnnotations(read_only_hint=True, open_world_hint=False),
)
async def preview_project(
    project_id: str = Field(
        pattern=r"^[0-9a-f]{16}$",
        description="From `create_project` or `open_project`.",
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

    async with _admitted(user_id), projects.load(user_id, project_id) as project:
        image = await officecli.screenshot(project.document, page)

    subject = f"Page {page}" if page else "The whole document"

    # OpenWebUI passes image resources to the model, image content only to
    # the user, so the image goes out as both.
    return ToolResult(content=[
        TextContent(
            type="text",
            text=(
                f"{subject} is attached as an image and shown to the user. "
                "It approximates Office's rendering: judge layout, overflow "
                "and missing content, not pixel details. Text colors from "
                "the master may render wrong, so never change colors because "
                "of the preview alone. Fix what is wrong with `run_commands`, "
                "then call `export_project`."
            ),
        ),
        File(data=image, name=f"page-{page or 'all'}.png").to_resource_content(
            mime_type="image/png",
        ),
        Image(data=image).to_image_content(),
    ])


@tool(
    name="export_project",
    tags={"office", "export"},
    description=(
        "Check the project's document and upload it to OpenWebUI. Call it once "
        "the user's request is fully applied, not after each change. The "
        "project stays editable; export it again after later edits."
    ),
    annotations=ToolAnnotations(destructive_hint=False, open_world_hint=True),
)
async def export_project(
    project_id: str = Field(
        pattern=r"^[0-9a-f]{16}$",
        description="From `create_project` or `open_project`.",
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

    async with _admitted(user_id), projects.load(user_id, project_id) as project:

        upload_name = f"{file_name}.{project.module.FORMAT}"
        findings = await officecli.findings(project.document)

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
            "Share `owui_url` with the user as the download link and mention "
            "serious `warnings`. The project stays editable; export it again "
            "after later edits."
        ),
        file_name=upload_name,
        owui_url=owui_url,
        warnings=[
            finding for finding in findings if finding not in project.baseline
        ],
    )
