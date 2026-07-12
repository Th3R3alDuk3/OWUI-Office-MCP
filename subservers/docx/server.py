from asyncio import to_thread
from io import BytesIO
from zipfile import BadZipFile

from docx import Document
from fastmcp import FastMCP
from fastmcp.dependencies import CurrentAccessToken, Depends, TokenClaim
from fastmcp.server.auth import AccessToken
from pydantic import Field

from config import get_settings
from models.docx import (
    FinalizeResult,
    Project,
    ScriptResult,
    StartResult,
    TemplatesResult,
)
from services.owui import download_file, upload_file
from subservers._sandbox import run_sandboxed
from subservers._store import ProjectStore
from subservers.docx._facade import SCRIPT_API, script_functions
from subservers.docx._ops import (
    clear_document,
    count_blocks,
    list_style_infos,
    list_template_names,
)

_DOCX_MIME = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
)

_EDIT_HINT = (
    "Continue the edit batch; when the user's request is fully applied, "
    "call `finalize_project` once."
)

_FINALIZE_HINT = (
    "Share `owui_url` with the user as the download link. The project stays "
    "active; end a later edit batch with `finalize_project` again."
)


_settings = get_settings()

_store = ProjectStore[Project](
    max_size=1_000,
    ttl=_settings.project_ttl,
    sweep_interval=_settings.project_sweep_interval,
)


mcp = FastMCP(name="docx", lifespan=_store.lifespan)


@mcp.tool(
    name="list_templates",
    description=(
        "List the available templates. Use this when the user did NOT attach "
        "a file, then pick one for `create_project`. If the user attached a "
        "file, use `open_project` instead."
    ),
)
async def list_templates() -> TemplatesResult:

    templates = await to_thread(list_template_names, _settings.templates_dir)

    hint = (
        "Pick a template and call `create_project`."
    ) if templates else (
        "No templates available. Ask the administrator to add `.docx` templates."
    )

    return TemplatesResult(hint=hint, templates=templates)


@mcp.tool(
    name="create_project",
    description=(
        "Create a new, empty in-memory project from a template. Use this by "
        "default when the user did NOT attach a file. Overwrites any existing "
        "project for the user. The result lists the template's styles for "
        "`run_script`."
    ),
)
async def create_project(
    template_name: str = Field(description="From `list_templates`."),
    user_id: str = TokenClaim("id"),
) -> StartResult:

    templates = await to_thread(list_template_names, _settings.templates_dir)

    if template_name not in templates:
        raise ValueError(
            f"Template '{template_name}' not found. "
            "Pick one from `list_templates`."
        )

    template_file = _settings.templates_dir.joinpath(template_name)

    document = await to_thread(Document, template_file)
    clear_document(document)

    _store.set(user_id, Project(document=document))

    return StartResult(
        hint=(
            f"Empty project created from template '{template_name}'. "
            "Build the document with `run_script`, using the styles below."
        ),
        block_count=0,
        styles=list_style_infos(document),
    )


@mcp.tool(
    name="open_project",
    description=(
        "Open a `.docx` the user attached in OpenWebUI, by its `file_id`. Use "
        "only when the user actually attached a file; if none was given, use "
        "`create_project` instead. Overwrites any existing project for the "
        "user. The result lists the file's styles for `run_script`."
    ),
)
async def open_project(
    file_id: str = Field(
        description=(
            "OpenWebUI file ID of the file the user attached. NOT a template "
            "name from `list_templates`, and never invented."
        ),
    ),
    token: AccessToken = CurrentAccessToken(),
    user_id: str = TokenClaim("id"),
) -> StartResult:

    file_content = await download_file(
        file_id=file_id,
        token=token.token,
    )

    try:
        document = await to_thread(Document, BytesIO(file_content))
    except (BadZipFile, ValueError) as error:
        raise ValueError(
            f"File '{file_id}' is not a valid `.docx` document."
        ) from error

    _store.set(user_id, Project(document=document))

    return StartResult(
        hint=(
            f"Project opened from attached file '{file_id}'. "
            "Inspect and edit it with `run_script` (`list_blocks()` shows "
            "the existing content), using the styles below for new blocks."
        ),
        block_count=count_blocks(document),
        styles=list_style_infos(document),
    )


@mcp.tool(
    name="run_script",
    description=SCRIPT_API,
)
async def run_script(
    code: str = Field(description="Python script for the sandbox."),
    token: AccessToken = CurrentAccessToken(),
    user_id: str = TokenClaim("id"),
    project: Project = Depends(_store.require),
) -> ScriptResult:

    async with project.lock:

        output = await run_sandboxed(
            code,
            functions=script_functions(project.document, token.token),
        )

        _store.touch(user_id, project)

        return ScriptResult(
            hint=_EDIT_HINT,
            block_count=count_blocks(project.document),
            output=output,
        )


@mcp.tool(
    name="finalize_project",
    description=(
        "Final step of an edit batch: serialize the current project to "
        "`.docx` and upload it to OpenWebUI. Call exactly once after the "
        "requested batch of changes, not after each one. The project stays "
        "active afterwards, so a later edit batch ends with another single "
        "`finalize_project` call."
    ),
)
async def finalize_project(
    file_name: str = Field(
        min_length=1, max_length=60,
        description="Stem without `.docx`.",
    ),
    token: AccessToken = CurrentAccessToken(),
    user_id: str = TokenClaim("id"),
    project: Project = Depends(_store.require),
) -> FinalizeResult:

    async with project.lock:

        upload_name = f"{file_name}.docx"

        buffer = BytesIO()
        await to_thread(project.document.save, buffer)

        _store.touch(user_id, project)

        block_count = count_blocks(project.document)

    owui_url = await upload_file(
        file_name=upload_name,
        data=buffer.getvalue(),
        content_type=_DOCX_MIME,
        token=token.token,
    )

    return FinalizeResult(
        hint=_FINALIZE_HINT,
        file_name=upload_name,
        block_count=block_count,
        owui_url=owui_url,
    )
