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
from models.docx import BlockInfo, DownloadProjectResponse, Project, StyleInfo
from services.owui import upload_file
from subservers.docx._utils import (
    count_blocks,
    drop_all_blocks,
    drop_blocks,
    list_block_infos,
    list_style_infos,
    list_template_names,
    move_block as _move_block,
)

DOCX_MIME = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
)


_settings = get_settings()

_projects: TTLCache[str, Project] = TTLCache(
    maxsize=1_000, ttl=_settings.project_ttl_seconds)

async def _ttl_task(
    interval: float,
) -> None:
    while True:
        await sleep(interval)
        _projects.expire()


def _get_project(
    user_id: str = TokenClaim("id"),
) -> Project:

    project = _projects.get(user_id)

    if project is None:
        raise ValueError("No project for this user. Call `create_project` first.")

    return project


def _touch(
    user_id: str,
    project: Project,
) -> None:
    if _projects.get(user_id) is project:
        _projects[user_id] = project


@asynccontextmanager
async def lifespan(
    server: FastMCP,
) -> AsyncIterator[None]:

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
        "`insert_paragraph` / `insert_table`."
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
        "any existing project for the user. Then call `insert_paragraph` / "
        "`insert_table` per block."
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
    name="insert_paragraph",
    description=(
        "Step 4a: insert a paragraph. Use a paragraph style from `list_styles` "
        "(e.g. `Heading 1`, `Normal`) to control formatting. Without "
        "`block_index` the paragraph is appended; otherwise it is inserted at "
        "that zero-based position. Returns the new block count. After the "
        "user's requested batch of edits is finished, always call "
        "`download_project` exactly once — not after every individual "
        "insert/move/remove."
    ),
)
async def insert_paragraph(
    text: str = Field(description="Paragraph text."),
    style: str | None = Field(
        default=None,
        description="Paragraph style name from `list_styles`. None = default.",
    ),
    block_index: int | None = Field(
        default=None,
        description=(
            "Zero-based position to insert at. If omitted, the block is "
            "appended at the end."
        ),
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

        if block_index is not None:
            _move_block(
                project.document, count_blocks(project.document) - 1, block_index)

        _touch(user_id, project)
        return count_blocks(project.document)


@mcp.tool(
    name="insert_table",
    description=(
        "Step 4b: insert a table with `rows` x `cols` cells. Optionally fill "
        "cells from `data` (row-major; extra rows/cols are ignored). Without "
        "`block_index` the table is appended; otherwise it is inserted at that "
        "zero-based position. Returns the new block count. After the user's "
        "requested batch of edits is finished, always call `download_project` "
        "exactly once — not after every individual insert/move/remove."
    ),
)
async def insert_table(
    rows: int = Field(
        gt=0, le=50,
        description="Number of rows."
    ),
    cols: int = Field(
        gt=0, le=10,
        description="Number of columns."
    ),
    style: str | None = Field(
        default=None,
        description="Table style name from `list_styles`. None = default.",
    ),
    data: list[list[str]] = Field(
        default_factory=list,
        description="Row-major cell text. Missing cells stay empty.",
    ),
    block_index: int | None = Field(
        default=None,
        description=(
            "Zero-based position to insert at. If omitted, the block is "
            "appended at the end."
        ),
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

        if block_index is not None:
            _move_block(
                project.document, count_blocks(project.document) - 1, block_index)

        _touch(user_id, project)
        return count_blocks(project.document)


@mcp.tool(
    name="insert_page_break",
    description=(
        "Step 4c: insert a page break as a body block. Without `block_index` "
        "it is appended; otherwise it is inserted at that zero-based position. "
        "Returns the new block count. After the user's requested batch of "
        "edits is finished, always call `download_project` exactly once — not "
        "after every individual insert/move/remove."
    ),
)
async def insert_page_break(
    block_index: int | None = Field(
        default=None,
        description=(
            "Zero-based position to insert at. If omitted, the block is "
            "appended at the end."
        ),
    ),
    user_id: str = TokenClaim("id"),
    project: Project = Depends(_get_project),
) -> int:

    async with project.lock:

        project.document.add_page_break()

        if block_index is not None:
            _move_block(
                project.document, count_blocks(project.document) - 1, block_index)

        _touch(user_id, project)
        return count_blocks(project.document)


@mcp.tool(
    name="list_blocks",
    description=(
        "List the current body blocks in order. The list position is the "
        "zero-based index; each entry has a type (`paragraph` or `table`) and "
        "a text preview. Use it to target `move_block` or `remove_blocks`."
    ),
)
async def list_blocks(
    project: Project = Depends(_get_project),
) -> list[BlockInfo]:
    async with project.lock:
        return list_block_infos(project.document)


@mcp.tool(
    name="move_block",
    description=(
        "Move a body block (paragraph or table) to a new position by "
        "zero-based index. Negative `to_index` counts from the end. Changes "
        "stay in memory only. After the user's requested batch of edits is "
        "finished, always call `download_project` exactly once — not after "
        "every individual insert/move/remove."
    ),
)
async def move_block(
    from_index: int = Field(
        description="Zero-based current block index.",
    ),
    to_index: int = Field(
        description="Zero-based target block index.",
    ),
    user_id: str = TokenClaim("id"),
    project: Project = Depends(_get_project),
) -> int:

    async with project.lock:

        _move_block(project.document, from_index, to_index)

        _touch(user_id, project)
        return count_blocks(project.document)


@mcp.tool(
    name="remove_blocks",
    description=(
        "Remove body blocks (paragraphs and tables) by zero-based index "
        "(from `list_blocks`). Indices refer to positions before removal; "
        "duplicates are ignored. Changes stay in memory only. After the "
        "user's requested batch of edits is finished, always call "
        "`download_project` exactly once — not after every individual "
        "insert/move/remove."
    ),
)
async def remove_blocks(
    indices: list[int] = Field(
        description="Zero-based block indices."
    ),
    user_id: str = TokenClaim("id"),
    project: Project = Depends(_get_project),
) -> int:

    async with project.lock:

        drop_blocks(project.document, indices)

        _touch(user_id, project)

        return count_blocks(project.document)


@mcp.tool(
    name="download_project",
    description=(
        "Step 5 — final step after a completed project editing request: "
        "serialize the current project to `.docx` and upload it to OpenWebUI. "
        "Always call this exactly once after the requested batch of inserts, "
        "moves, or removals is finished. Do not call it after every "
        "individual change when multiple changes belong to the same request. "
        "The project stays active afterwards, so a later editing request ends "
        "with another single `download_project` call."
    ),
)
async def download_project(
    file_name: str = Field(
        min_length=3, max_length=30,
        description="Stem without `.docx`.",
    ),
    token: AccessToken = CurrentAccessToken(),
    user_id: str = TokenClaim("id"),
    project: Project = Depends(_get_project),
) -> DownloadProjectResponse:

    async with project.lock:

        out_name = f"{file_name.strip()}.docx"

        buffer = BytesIO()
        await to_thread(project.document.save, buffer)

        _touch(user_id, project)
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
