from asyncio import to_thread
from io import BytesIO
from zipfile import BadZipFile

from fastmcp import FastMCP
from fastmcp.dependencies import CurrentAccessToken, Depends, TokenClaim
from fastmcp.server.auth import AccessToken
from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException
from pydantic import Field

from config import get_settings
from models.xlsx import (
    CellInput,
    Project,
    ProjectResponse,
    SheetInfo,
)
from services.owui import DOWNLOAD_FILE_URL, download_file, upload_file
from subservers._store import ProjectStore
from subservers.xlsx._utils import (
    insert_sheet as _insert_sheet,
    clear_workbook,
    count_sheets,
    drop_sheets,
    list_sheet_infos,
    list_template_names,
    move_sheet as _move_sheet,
    read_sheet_rows,
    write_cells as _write_cells,
    write_rows as _write_rows,
)

XLSX_MIME = (
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
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


mcp = FastMCP(name="xlsx")


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
        "project for the user. Then call `list_sheets`."
    ),
)
async def create_project(
    template_name: str = Field(description="From `list_templates`."),
    user_id: str = TokenClaim("id"),
) -> str:

    templates = await to_thread(list_template_names, _settings.templates_dir)

    if template_name not in templates:
        raise ValueError(
            f"Template '{template_name}' not found. "
            "Pick one from `list_templates`."
        )

    template_file = _settings.templates_dir.joinpath(template_name)

    workbook = await to_thread(load_workbook, template_file)
    clear_workbook(workbook)

    _store.set(user_id, Project(workbook=workbook))

    return (
        f"Project created from template '{template_name}'. "
        "Call `list_sheets` to see the initial sheet and `list_styles` for "
        "available styles, then `insert_sheet`, `write_rows`, etc. to edit it."
    )


@mcp.tool(
    name="open_project",
    description=(
        "Open a `.xlsx` the user attached in OpenWebUI, by its `file_id`. Use "
        "only when the user actually attached a file; if none was given, use "
        "`create_project` instead. Overwrites any existing project for the "
        "user. Then call `list_sheets` to inspect or extend it."
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
) -> str:

    file_content = await download_file(
        file_id=file_id,
        token=token.token,
        base_url=_settings.owui_base_url,
    )

    try:
        workbook = await to_thread(load_workbook, BytesIO(file_content))
    except (BadZipFile, InvalidFileException) as error:
        raise ValueError(
            f"File '{file_id}' is not a valid `.xlsx` workbook."
        ) from error

    _store.set(user_id, Project(workbook=workbook))

    return (
        f"Project opened from attached file '{file_id}'. "
        "Call `list_styles` to see available styles and `list_sheets` to see its initial state, "
        "then `insert_sheet`, `write_rows`, etc. to edit it."
    )


@mcp.tool(
    name="list_sheets",
    description=(
        "List the worksheets of the current project with their used extent "
        "(`rows` x `cols`). Use the titles to target the other tools."
    ),
)
async def list_sheets(
    project: Project = Depends(_get_project),
) -> list[SheetInfo]:
    async with project.lock:
        return list_sheet_infos(project.workbook)


@mcp.tool(
    name="list_styles",
    description=(
        "List the names of the named cell styles in the current project. Pass "
        "one to `write_cells` or `write_rows` to format cells. Optional — "
        "cells without a style use the default formatting."
    ),
)
async def list_styles(
    project: Project = Depends(_get_project),
) -> list[str]:
    async with project.lock:
        return list(project.workbook.named_styles)


@mcp.tool(
    name="insert_sheet",
    description=(
        "Insert an empty worksheet. Without `index` it is appended; otherwise it "
        "is inserted at that zero-based position. Returns the new sheet count as `{sheet_count}`. "
        "After the requested batch of edits, call `finalize_project` once "
        "(not after each change)."
    ),
)
async def insert_sheet(
    title: str = Field(description="Title for the new worksheet."),
    index: int | None = Field(
        default=None,
        description=(
            "Zero-based position to insert at. If omitted, the sheet is "
            "appended at the end."
        ),
    ),
    user_id: str = TokenClaim("id"),
    project: Project = Depends(_get_project),
) -> dict[str, int]:

    async with project.lock:

        _insert_sheet(project.workbook, title, index)

        _store.touch(user_id, project)

        return {"sheet_count": count_sheets(project.workbook)}


@mcp.tool(
    name="write_rows",
    description=(
        "Fill a contiguous block of cells from `rows` (row-major), starting "
        "at `anchor` (top-left, default `A1`). Best for tables and bulk data; "
        "for scattered or individually-styled cells use `write_cells`. An "
        "optional `style` from `list_styles` formats every written cell. "
        "Changes stay in memory only. After the requested batch of edits, "
        "call `finalize_project` once (not after each change)."
    ),
)
async def write_rows(
    sheet: str = Field(description="Worksheet title from `list_sheets`."),
    rows: list[list[bool | int | float | str | None]] = Field(
        description=(
            "Row-major values. Numbers stay numeric, a string starting with "
            "`=` becomes a formula, and `null` clears the cell."
        ),
    ),
    anchor: str = Field(
        default="A1",
        description="A1 reference of the top-left cell.",
    ),
    style: str | None = Field(
        default=None,
        description="Named cell style from `list_styles` for every cell.",
    ),
    user_id: str = TokenClaim("id"),
    project: Project = Depends(_get_project),
) -> dict[str, int]:

    async with project.lock:

        _write_rows(project.workbook, sheet, anchor, rows, style)

        _store.touch(user_id, project)

    return {
        "rows": len(rows),
        "cols": len(rows[0]) if rows else 0,
    }


@mcp.tool(
    name="write_cells",
    description=(
        "Write values into individual cells by A1 reference, optionally "
        "formatting each with a named style from `list_styles`. Best for "
        "scattered edits or styling specific cells; to fill a contiguous "
        "table use `write_rows`. At most a limited number of cells per call. "
        "Only the listed cells change; the rest stay as-is. Changes stay in "
        "memory only. "
        "After the requested batch of edits, call `finalize_project` once "
        "(not after each change)."
    ),
)
async def write_cells(
    sheet: str = Field(description="Worksheet title from `list_sheets`."),
    cells: list[CellInput] = Field(
        max_length=100,
        description="Cells to write, each targeting one A1 reference.",
    ),
    user_id: str = TokenClaim("id"),
    project: Project = Depends(_get_project),
) -> dict[str, int]:

    async with project.lock:

        _write_cells(project.workbook, sheet, cells)

        _store.touch(user_id, project)

    return {"cells": len(cells)}


@mcp.tool(
    name="read_sheet",
    description=(
        "Read a worksheet's used range as rows of plain text (row-major, empty "
        "cells as ``). Formulas are returned as their formula string, not the "
        "computed result."
    ),
)
async def read_sheet(
    sheet: str = Field(description="Worksheet title from `list_sheets`."),
    project: Project = Depends(_get_project),
) -> list[list[str]]:
    async with project.lock:
        return read_sheet_rows(project.workbook, sheet)


@mcp.tool(
    name="move_sheet",
    description=(
        "Move a worksheet to a new position by zero-based index. Negative "
        "`to_index` counts from the end. Changes stay in memory only. After "
        "the requested batch of edits, call `finalize_project` once (not "
        "after each change)."
    ),
)
async def move_sheet(
    title: str = Field(description="Worksheet title from `list_sheets`."),
    to_index: int = Field(description="Zero-based target position."),
    user_id: str = TokenClaim("id"),
    project: Project = Depends(_get_project),
) -> list[SheetInfo]:

    async with project.lock:

        _move_sheet(project.workbook, title, to_index)

        _store.touch(user_id, project)

        return list_sheet_infos(project.workbook)


@mcp.tool(
    name="remove_sheets",
    description=(
        "Remove worksheets by title (from `list_sheets`). A workbook must keep "
        "at least one sheet. Changes stay in memory only. After the requested "
        "batch of edits, call `finalize_project` once (not after each change)."
    ),
)
async def remove_sheets(
    titles: list[str] = Field(description="Worksheet titles to remove."),
    user_id: str = TokenClaim("id"),
    project: Project = Depends(_get_project),
) -> dict[str, int]:

    async with project.lock:

        drop_sheets(project.workbook, titles)

        _store.touch(user_id, project)

        return {"sheet_count": count_sheets(project.workbook)}


@mcp.tool(
    name="finalize_project",
    description=(
        "Final step of an edit batch: serialize the current project to "
        "`.xlsx` and upload it to OpenWebUI. Call exactly once after the "
        "requested batch of changes, not after each one. The project stays "
        "active afterwards, so a later edit batch ends with another single "
        "`finalize_project` call."
    ),
)
async def finalize_project(
    file_name: str = Field(
        min_length=1, max_length=60,
        description="Stem without `.xlsx`.",
    ),
    token: AccessToken = CurrentAccessToken(),
    user_id: str = TokenClaim("id"),
    project: Project = Depends(_get_project),
) -> ProjectResponse:

    async with project.lock:

        out_name = f"{file_name}.xlsx"

        buffer = BytesIO()
        await to_thread(project.workbook.save, buffer)

        _store.touch(user_id, project)

        sheet_count = count_sheets(project.workbook)

    uploaded = await upload_file(
        file_name=out_name,
        data=buffer.getvalue(),
        content_type=XLSX_MIME,
        token=token.token,
        base_url=_settings.owui_base_url,
    )

    return ProjectResponse(
        file_name=out_name,
        sheet_count=sheet_count,
        owui_url=DOWNLOAD_FILE_URL.format(
            base_url=_settings.owui_base_url, file_id=uploaded.id)
    )
