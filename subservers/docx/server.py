from asyncio import CancelledError, create_task, sleep, to_thread
from collections.abc import AsyncIterator
from contextlib import asynccontextmanager, suppress
from io import BytesIO

from cachetools import TTLCache
from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from fastmcp import FastMCP
from fastmcp.dependencies import CurrentAccessToken, Depends, TokenClaim
from fastmcp.server.auth import AccessToken
from pydantic import Field

from config import get_settings
from models.docx import DownloadProjectResponse, Project, StyleInfo
from services.owui import upload_file
from subservers.docx._utils import (
    count_blocks,
    drop_all_blocks,
    drop_block,
    list_style_infos,
    list_template_names,
)

DOCX_MIME = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
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


mcp = FastMCP(name="docx")


@mcp.tool(
    name="list_templates",
    description=(
        "Step 1: list available templates. Pick one, then call `list_styles`."
    ),
)
async def list_templates() -> list[str]:
    return await to_thread(
        list_template_names, _settings.templates_dir)


@mcp.tool(
    name="list_styles",
    description=(
        "Step 2: list paragraph and table styles of a template "
        "(name -> type, builtin). Then call `create_project` and "
        "`append_paragraph` / `append_table`."
    ),
)
async def list_styles(
    template_name: str = Field(description="From `list_templates`."),
) -> dict[str, StyleInfo]:
    return await to_thread(
        list_style_infos, _settings.templates_dir, template_name)


@mcp.tool(
    name="create_project",
    description=(
        "Step 3: create an empty in-memory project from a template. Overwrites "
        "any existing project for the user. Then call `append_paragraph` / "
        "`append_table` per block."
    ),
)
async def create_project(
    template_name: str = Field(description="From `list_templates`."),
    user_id: str = TokenClaim("id"),
) -> None:

    template_file = _settings.templates_dir.joinpath(template_name)

    document = await to_thread(Document, template_file)
    drop_all_blocks(document)

    _projects[user_id] = Project(document=document)


@mcp.tool(
    name="append_paragraph",
    description=(
        "Step 4a: append a paragraph. Use a paragraph style from `list_styles` "
        "(e.g. `Heading 1`, `Normal`) to control formatting. Returns the new "
        "block count."
    ),
)
async def append_paragraph(
    text: str = Field(description="Paragraph text."),
    style: str | None = Field(
        default=None,
        description="Paragraph style name from `list_styles`. None = default.",
    ),
    user_id: str = TokenClaim("id"),
    project: Project = Depends(_get_project),
) -> int:

    async with project.lock:

        try:
            project.document.styles.get_style_id(style, WD_STYLE_TYPE.PARAGRAPH)
        except (KeyError, ValueError):
            raise ValueError(
                f"Style '{style}' not found."
            )

        project.document.add_paragraph(text, style=style)

        _projects[user_id] = project
        return count_blocks(project.document)


@mcp.tool(
    name="append_table",
    description=(
        "Step 4b: append a table with `rows` x `cols` cells. Optionally fill "
        "cells from `data` (row-major; extra rows/cols are ignored). Returns "
        "the new block count."
    ),
)
async def append_table(
    rows: int = Field(gt=0, description="Number of rows."),
    cols: int = Field(gt=0, description="Number of columns."),
    style: str | None = Field(
        default=None,
        description="Table style name from `list_styles`. None = default.",
    ),
    data: list[list[str]] = Field(
        default_factory=list,
        description="Row-major cell text. Missing cells stay empty.",
    ),
    user_id: str = TokenClaim("id"),
    project: Project = Depends(_get_project),
) -> int:

    async with project.lock:

        try:
            project.document.styles.get_style_id(style, WD_STYLE_TYPE.TABLE)
        except (KeyError, ValueError):
            raise ValueError(
                f"Style '{style}' not found."
            )

        table = project.document.add_table(rows=rows, cols=cols)

        if style is not None:
            table.style = style

        for r, row_data in enumerate(data[:rows]):
            for c, cell_text in enumerate(row_data[:cols]):
                table.cell(r, c).text = cell_text

        _projects[user_id] = project
        return count_blocks(project.document)


@mcp.tool(
    name="remove_blocks",
    description=(
        "Remove body blocks (paragraphs and tables) by zero-based index. "
        "Indices refer to positions before removal; duplicates are ignored."
    ),
)
async def remove_blocks(
    indices: list[int] = Field(description="Zero-based block indices."),
    user_id: str = TokenClaim("id"),
    project: Project = Depends(_get_project),
) -> int:

    async with project.lock:

        for i in sorted(set(indices), reverse=True):
            drop_block(project.document, i)

        _projects[user_id] = project

        return count_blocks(project.document)


@mcp.tool(
    name="download_project",
    description=(
        "Step 5: serialize the project to `.docx` and upload it to OpenWebUI. "
        "The project stays active after saving."
    ),
)
async def download_project(
    file_name: str = Field(
        min_length=3, max_length=30,
        description="Stem without `.docx`.",
    ),
    user_id: str = TokenClaim("id"),
    project: Project = Depends(_get_project),
    token: AccessToken = CurrentAccessToken(),
) -> DownloadProjectResponse:

    async with project.lock:

        out_name = f"{file_name.strip()}.docx"

        buffer = BytesIO()
        await to_thread(project.document.save, buffer)

        _projects[user_id] = project
        block_count = count_blocks(project.document)

    base_url = _settings.owui_base_url.rstrip("/")

    uploaded = await upload_file(
        filename=out_name,
        data=buffer.getvalue(),
        content_type=DOCX_MIME,
        token=token.token,
        base_url=base_url,
    )

    return DownloadProjectResponse(
        filename=out_name,
        block_count=block_count,
        owui_url=f"{base_url}/api/v1/files/{uploaded.id}/content",
    )
