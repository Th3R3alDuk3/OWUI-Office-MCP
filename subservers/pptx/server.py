from asyncio import to_thread
from contextlib import suppress
from io import BytesIO
from zipfile import BadZipFile

from fastmcp import FastMCP
from fastmcp.dependencies import CurrentAccessToken, Depends, TokenClaim
from fastmcp.server.auth import AccessToken
from pptx import Presentation
from pydantic import Field

from config import get_settings
from models.pptx import (
    LayoutInfo,
    PlaceholderText,
    Project,
    ProjectResponse,
    SlideInfo,
)
from services.owui import download_file, upload_file
from subservers._store import ProjectStore
from subservers.pptx._utils import (
    count_slides,
    drop_all_slides,
    drop_slides,
    list_layout_infos,
    list_master_names,
    list_slide_infos,
    list_template_names,
    move_slide as _move_slide,
)

PPTX_MIME = (
    "application/vnd.openxmlformats-officedocument.presentationml.presentation"
)


_settings = get_settings()

_store = ProjectStore(
    max_size=1_000,
    ttl=_settings.project_ttl_seconds,
    sweep_interval=_settings.project_sweep_interval_seconds,
)

lifespan = _store.lifespan


def _get_project(
    user_id: str = TokenClaim("id"),
) -> Project:

    project = _store.get(user_id)

    if project is None:
        raise ValueError(
            "No project for this user. "
            "Call `create_project` or `open_project` first."
        )

    return project


mcp = FastMCP(name="pptx")


@mcp.tool(
    name="list_templates",
    description=(
        "List the available templates. Use this when the user did NOT attach "
        "a file, then pick one for `create_project`. If the user attached a "
        "file, use `open_project` instead."
    ),
)
async def list_templates() -> list[str]:
    return await to_thread(
        list_template_names, _settings.templates_dir)


@mcp.tool(
    name="create_project",
    description=(
        "Create a new, empty in-memory project from a template. Use this by "
        "default when the user did NOT attach a file. Overwrites any existing "
        "project for the user. Then call `list_masters`."
    ),
)
async def create_project(
    template_name: str = Field(description="From `list_templates`."),
    user_id: str = TokenClaim("id"),
) -> None:

    templates = await to_thread(list_template_names, _settings.templates_dir)

    if template_name not in templates:
        raise ValueError(
            f"Template '{template_name}' not found. "
            "Pick one from `list_templates`."
        )

    template_file = _settings.templates_dir.joinpath(template_name)

    presentation = await to_thread(Presentation, template_file)
    drop_all_slides(presentation)

    _store.set(user_id, Project(presentation=presentation))


@mcp.tool(
    name="open_project",
    description=(
        "Open a `.pptx` the user attached in OpenWebUI, by its `file_id`. Use "
        "only when the user actually attached a file; if none was given, use "
        "`create_project` instead. Overwrites any existing project for the "
        "user. Then call `list_masters` to add slides, or `list_slides` to "
        "edit existing ones."
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
) -> None:

    file_content = await download_file(
        file_id=file_id,
        token=token.token,
        base_url=_settings.owui_base_url,
    )

    try:
        presentation = await to_thread(Presentation, BytesIO(file_content))
    except (BadZipFile, ValueError) as error:
        raise ValueError(
            f"File '{file_id}' is not a valid `.pptx` presentation."
        ) from error

    _store.set(user_id, Project(presentation=presentation))


@mcp.tool(
    name="list_masters",
    description=(
        "List the slide masters of the current project. Pick one, then call "
        "`list_layouts`."
    ),
)
async def list_masters(
    project: Project = Depends(_get_project),
) -> dict[int, str]:
    async with project.lock:
        return list_master_names(project.presentation)


@mcp.tool(
    name="list_layouts",
    description=(
        "List the layouts of a master with their placeholders (idx, name, "
        "type). Then call `insert_slide`."
    ),
)
async def list_layouts(
    master_index: int = Field(description="From `list_masters`."),
    project: Project = Depends(_get_project),
) -> dict[str, LayoutInfo]:
    async with project.lock:
        try:
            return list_layout_infos(project.presentation, master_index)
        except IndexError:
            raise ValueError(f"Master '{master_index}' not found.") from None


@mcp.tool(
    name="insert_slide",
    description=(
        "Insert a slide using a layout from `list_layouts`, optionally "
        "filling text placeholders by `idx`. Without `slide_index` the slide "
        "is appended; otherwise it is inserted at that zero-based position. "
        "After the requested batch of edits, call `finalize_project` once "
        "(not after each change)."
    ),
)
async def insert_slide(
    master_index: int = Field(description="From `list_masters`."),
    layout_name: str = Field(description="From `list_layouts`."),
    placeholders: list[PlaceholderText] = Field(
        default_factory=list,
        description="Text placeholders to fill, each targeting one `idx`.",
    ),
    slide_index: int | None = Field(
        default=None,
        description=(
            "Zero-based position to insert at. If omitted, the slide is "
            "appended at the end."
        ),
    ),
    user_id: str = TokenClaim("id"),
    project: Project = Depends(_get_project),
) -> int:

    async with project.lock:

        try:
            master = project.presentation.slide_masters[master_index]
        except IndexError:
            raise ValueError(
                f"Master '{master_index}' not found."
            ) from None

        layout = master.slide_layouts.get_by_name(layout_name)

        if layout is None:
            raise ValueError(
                f"Layout '{layout_name}' not found in master index {master_index}."
            )

        slide = project.presentation.slides.add_slide(layout)

        for placeholder in placeholders:
            with suppress(KeyError):
                slide.placeholders[placeholder.idx].text = placeholder.text

        if slide_index is not None:
            _move_slide(
                project.presentation,
                count_slides(project.presentation) - 1,
                slide_index,
            )

        _store.touch(user_id, project)
        return count_slides(project.presentation)


@mcp.tool(
    name="list_slides",
    description=(
        "List the current slides in order. The list position is the zero-based "
        "index; each entry has the layout name and slide text. Use it to target "
        "`edit_slide`, `move_slide`, or `remove_slides`."
    ),
)
async def list_slides(
    project: Project = Depends(_get_project),
) -> list[SlideInfo]:
    async with project.lock:
        return list_slide_infos(project.presentation)


@mcp.tool(
    name="edit_slide",
    description=(
        "Update text placeholders on an existing slide by zero-based index "
        "(from `list_slides`). Only the listed `idx` are changed; others stay "
        "as-is. Changes stay in memory only. After the requested batch of "
        "edits, call `finalize_project` once (not after each change)."
    ),
)
async def edit_slide(
    slide_index: int = Field(description="Zero-based slide index."),
    placeholders: list[PlaceholderText] = Field(
        description="Placeholders to update, each targeting one `idx`.",
    ),
    user_id: str = TokenClaim("id"),
    project: Project = Depends(_get_project),
) -> None:

    async with project.lock:

        try:
            slide = project.presentation.slides[slide_index]
        except IndexError:
            raise ValueError(
                f"Slide index {slide_index} out of range."
            ) from None

        for placeholder in placeholders:
            with suppress(KeyError):
                slide.placeholders[placeholder.idx].text = placeholder.text

        _store.touch(user_id, project)


@mcp.tool(
    name="move_slide",
    description=(
        "Move a slide to a new position by zero-based index. Negative "
        "`to_index` counts from the end. Changes stay in memory only. After "
        "the requested batch of edits, call `finalize_project` once (not "
        "after each change)."
    ),
)
async def move_slide(
    from_index: int = Field(
        description="Zero-based current slide index.",
    ),
    to_index: int = Field(
        description="Zero-based target slide index.",
    ),
    user_id: str = TokenClaim("id"),
    project: Project = Depends(_get_project),
) -> int:

    async with project.lock:

        _move_slide(project.presentation, from_index, to_index)

        _store.touch(user_id, project)
        return count_slides(project.presentation)


@mcp.tool(
    name="remove_slides",
    description=(
        "Remove slides by zero-based index (from `list_slides`). Indices refer "
        "to positions before removal; duplicates are ignored. Changes stay in "
        "memory only. After the requested batch of edits, call "
        "`finalize_project` once (not after each change)."
    ),
)
async def remove_slides(
    indices: list[int] = Field(description="Zero-based slide indices."),
    user_id: str = TokenClaim("id"),
    project: Project = Depends(_get_project),
) -> int:

    async with project.lock:

        drop_slides(project.presentation, indices)

        _store.touch(user_id, project)

        return count_slides(project.presentation)


@mcp.tool(
    name="finalize_project",
    description=(
        "Final step of an edit batch: serialize the current project to "
        "`.pptx` and upload it to OpenWebUI. Call exactly once after the "
        "requested batch of changes, not after each one. The project stays "
        "active afterwards, so a later edit batch ends with another single "
        "`finalize_project` call."
    ),
)
async def finalize_project(
    file_name: str = Field(
        min_length=1, max_length=60,
        description="Stem without `.pptx`.",
    ),
    token: AccessToken = CurrentAccessToken(),
    user_id: str = TokenClaim("id"),
    project: Project = Depends(_get_project),
) -> ProjectResponse:

    async with project.lock:

        out_name = f"{file_name}.pptx"

        buffer = BytesIO()
        await to_thread(project.presentation.save, buffer)

        _store.touch(user_id, project)
        slide_count = count_slides(project.presentation)

    uploaded = await upload_file(
        file_name=out_name,
        data=buffer.getvalue(),
        content_type=PPTX_MIME,
        token=token.token,
        base_url=_settings.owui_base_url,
    )

    return ProjectResponse(
        file_name=out_name,
        slide_count=slide_count,
        owui_url=(
            f"{_settings.owui_base_url}/api/v1/files/{uploaded.id}/content"
        ),
    )
