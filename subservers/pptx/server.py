from asyncio import CancelledError, create_task, sleep, to_thread
from collections.abc import AsyncIterator
from contextlib import asynccontextmanager, suppress
from io import BytesIO

from cachetools import TTLCache
from fastmcp import FastMCP
from fastmcp.dependencies import CurrentAccessToken, Depends, TokenClaim
from fastmcp.server.auth import AccessToken
from pptx import Presentation
from pydantic import Field

from config import get_settings
from models.pptx import DownloadProjectResponse, LayoutInfo, Project
from services.owui import upload_file
from subservers.pptx._utils import (
    drop_all_slides,
    drop_slide,
    list_layout_infos,
    list_master_names,
    list_template_names,
)

PPTX_MIME = (
    "application/vnd.openxmlformats-officedocument.presentationml.presentation"
)


_settings = get_settings()

_projects: TTLCache[str, Project] = TTLCache(
    maxsize=10_000, ttl=_settings.project_ttl_seconds)


async def _ttl_task(interval: float) -> None:
    while True:
        await sleep(interval)
        _projects.expire()


def _get_project(user_id: str = TokenClaim("id")) -> Project:

    project = _projects.get(user_id)

    if project is None:
        raise ValueError("No project for this user. Call `create_project` first.")

    return project


@asynccontextmanager
async def lifespan(server: FastMCP) -> AsyncIterator[None]:

    task = create_task(
        _ttl_task(_settings.project_sweep_interval_seconds))

    try:
        yield
    finally:

        task.cancel()

        with suppress(CancelledError):
            await task


mcp = FastMCP(name="powerpoint")


@mcp.tool(
    name="list_templates",
    description=(
        "Step 1: list available templates. Pick one, then call `list_masters`."
    ),
)
async def list_templates() -> list[str]:
    return await to_thread(
        list_template_names, _settings.templates_dir)


@mcp.tool(
    name="list_masters",
    description=(
        "Step 2: list slide masters of a template. Pick one, then call "
        "`list_layouts`."
    ),
)
async def list_masters(
    template_name: str = Field(description="From `list_templates`."),
) -> dict[int, str]:
    return await to_thread(
        list_master_names, _settings.templates_dir, template_name)


@mcp.tool(
    name="list_layouts",
    description=(
        "Step 3: list slide layouts of a master with their placeholders "
        "(idx, name, type). Then call `create_project` and `append_slide`."
    ),
)
async def list_layouts(
    template_name: str = Field(description="From `list_templates`."),
    master_index: int = Field(description="From `list_masters`."),
) -> dict[str, LayoutInfo]:
    return await to_thread(
        list_layout_infos, _settings.templates_dir, template_name, master_index)


@mcp.tool(
    name="create_project",
    description=(
        "Step 4: create an empty in-memory project from a template. Overwrites "
        "any existing project for the user. Then call `append_slide` per slide."
    ),
)
async def create_project(
    template_name: str = Field(description="From `list_templates`."),
    user_id: str = TokenClaim("id"),
) -> None:

    template_file = _settings.templates_dir.joinpath(template_name)

    presentation = await to_thread(Presentation, template_file)
    drop_all_slides(presentation)

    _projects[user_id] = Project(presentation=presentation)


@mcp.tool(
    name="append_slide",
    description=(
        "Step 5: append a slide using a layout from `list_layouts`, "
        "optionally filling text placeholders by `idx`. Repeat per slide, "
        "then call `download_project`."
    ),
)
async def append_slide(
    master_index: int = Field(description="From `list_masters`."),
    layout_name: str = Field(description="From `list_layouts`."),
    placeholders: dict[int, str] = Field(
        default_factory=dict,
        description="Placeholder `idx` -> text. Missing keys stay empty.",
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
            )

        layout = master.slide_layouts.get_by_name(layout_name)

        if layout is None:
            raise ValueError(
                f"Layout '{layout_name}' not found in master index {master_index}."
            )

        slide = project.presentation.slides.add_slide(layout)

        for idx, text in placeholders.items():
            with suppress(KeyError):
                slide.placeholders[idx].text = text

        _projects[user_id] = project
        return len(project.presentation.slides)


@mcp.tool(
    name="edit_slide",
    description=(
        "Update text placeholders on an existing slide by zero-based index. "
        "Only listed `idx` keys are touched; others stay as-is."
    ),
)
async def edit_slide(
    slide_index: int = Field(description="Zero-based slide index."),
    placeholders: dict[int, str] = Field(
        description="Placeholder `idx` -> new text.",
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
            )

        for idx, text in placeholders.items():
            with suppress(KeyError):
                slide.placeholders[idx].text = text

        _projects[user_id] = project


@mcp.tool(
    name="remove_slides",
    description=(
        "Remove slides by zero-based index. Indices refer to positions "
        "before removal; duplicates are ignored."
    ),
)
async def remove_slides(
    indices: list[int] = Field(description="Zero-based slide indices."),
    user_id: str = TokenClaim("id"),
    project: Project = Depends(_get_project),
) -> int:

    async with project.lock:

        for i in sorted(set(indices), reverse=True):
            drop_slide(project.presentation, i)

        _projects[user_id] = project

        return len(project.presentation.slides)


@mcp.tool(
    name="download_project",
    description=(
        "Step 6: serialize the project to `.pptx` and upload it to OpenWebUI. "
        "The project stays active after saving."
    ),
)
async def download_project(
    file_name: str = Field(
        min_length=3, max_length=30,
        description="Stem without `.pptx`.",
    ),
    user_id: str = TokenClaim("id"),
    project: Project = Depends(_get_project),
    token: AccessToken = CurrentAccessToken(),
) -> DownloadProjectResponse:

    async with project.lock:

        out_name = f"{file_name.strip()}.pptx"

        buffer = BytesIO()
        await to_thread(project.presentation.save, buffer)

        _projects[user_id] = project
        slide_count = len(project.presentation.slides)

    base_url = _settings.owui_base_url.rstrip("/")

    uploaded = await upload_file(
        filename=out_name,
        data=buffer.getvalue(),
        content_type=PPTX_MIME,
        token=token.token,
        base_url=base_url,
    )

    return DownloadProjectResponse(
        filename=out_name,
        item_count=item_count,
        owui_url=f"{base_url}/api/v1/files/{uploaded.id}/content",
    )
