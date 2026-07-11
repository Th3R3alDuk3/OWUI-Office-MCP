from asyncio import to_thread
from io import BytesIO
from zipfile import BadZipFile

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from fastmcp import FastMCP
from fastmcp.dependencies import CurrentAccessToken, Depends, TokenClaim
from fastmcp.exceptions import ToolError
from fastmcp.server.auth import AccessToken
from pydantic import Field

from config import get_settings
from models._base import TemplatesResult
from models.docx import (
    BlocksResult,
    FinalizeResult,
    Project,
    ProjectResult,
    StylesResult,
)
from services.owui import download_file, upload_file
from subservers._store import ProjectStore
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

def _get_project(
    user_id: str = TokenClaim("id"),
) -> Project:

    project = _store.get(user_id)

    if project is None:
        # ToolError passes dependency resolution unchanged; other exceptions
        # get swallowed into a generic "failed to resolve" message.
        raise ToolError(
            "No project for this user. "
            "Call `create_project` or `open_project` first."
        )

    return project


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
        "project for the user. Then call `list_styles`."
    ),
)
async def create_project(
    template_name: str = Field(description="From `list_templates`."),
    user_id: str = TokenClaim("id"),
) -> ProjectResult:

    templates = await to_thread(list_template_names, _settings.templates_dir)

    if template_name not in templates:
        raise ValueError(
            f"Template '{template_name}' not found. "
            "Pick one from `list_templates`."
        )

    template_file = _settings.templates_dir.joinpath(template_name)

    document = await to_thread(Document, template_file)
    drop_all_blocks(document)

    _store.set(user_id, Project(document=document))

    return ProjectResult(
        hint=(
            f"Empty project created from template '{template_name}'. "
            "Call `list_styles` to see the available styles, then "
            "`insert_paragraph` / `insert_table` to add content."
        ),
        block_count=0,
    )


@mcp.tool(
    name="open_project",
    description=(
        "Open a `.docx` the user attached in OpenWebUI, by its `file_id`. Use "
        "only when the user actually attached a file; if none was given, use "
        "`create_project` instead. Overwrites any existing project for the "
        "user. Then call `list_styles` to add blocks, or `list_blocks` to "
        "reorder or remove existing ones."
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
) -> ProjectResult:

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

    return ProjectResult(
        hint=(
            f"Project opened from attached file '{file_id}'. "
            "Call `list_blocks` to see the existing content, or `list_styles` "
            "before adding blocks with `insert_paragraph` / `insert_table`."
        ),
        block_count=count_blocks(document),
    )


@mcp.tool(
    name="list_styles",
    description=(
        "List the paragraph and table styles of the current project (name -> "
        "type, builtin). Then call `insert_paragraph` / `insert_table`."
    ),
)
async def list_styles(
    project: Project = Depends(_get_project),
) -> StylesResult:

    async with project.lock:
        styles = list_style_infos(project.document)

    return StylesResult(
        hint=(
            "Use a paragraph style with `insert_paragraph` and a table style "
            "with `insert_table`."
        ),
        styles=styles,
    )


@mcp.tool(
    name="insert_paragraph",
    description=(
        "Insert a paragraph. Use a paragraph style from `list_styles` (e.g. "
        "`Heading 1`, `Normal`) to control formatting. Without `block_index` "
        "the paragraph is appended; otherwise it is inserted at that "
        "zero-based position. After the requested batch of edits, call "
        "`finalize_project` once (not after each change)."
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
) -> ProjectResult:

    async with project.lock:

        try:
            project.document.styles.get_style_id(style, WD_STYLE_TYPE.PARAGRAPH)
        except (KeyError, ValueError):
            raise ValueError(
                f"Style '{style}' not found."
            ) from None

        project.document.add_paragraph(text, style=style)

        if block_index is not None:
            _move_block(
                project.document,
                count_blocks(project.document) - 1,
                block_index,
            )

        _store.touch(user_id, project)

        return ProjectResult(
            hint=_EDIT_HINT,
            block_count=count_blocks(project.document),
        )


@mcp.tool(
    name="insert_table",
    description=(
        "Insert a table with `rows` x `cols` cells. Optionally fill cells "
        "from `data` (row-major; extra rows/cols are ignored). Without "
        "`block_index` the table is appended; otherwise it is inserted at "
        "that zero-based position. After the requested batch of edits, call "
        "`finalize_project` once (not after each change)."
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
) -> ProjectResult:

    async with project.lock:

        try:
            project.document.styles.get_style_id(style, WD_STYLE_TYPE.TABLE)
        except (KeyError, ValueError):
            raise ValueError(
                f"Style '{style}' not found."
            ) from None

        table = project.document.add_table(rows=rows, cols=cols)

        if style is not None:
            table.style = style

        for r, row_data in enumerate(data[:rows]):
            for c, cell_text in enumerate(row_data[:cols]):
                table.cell(r, c).text = cell_text

        if block_index is not None:
            _move_block(
                project.document,
                count_blocks(project.document) - 1,
                block_index,
            )

        _store.touch(user_id, project)

        return ProjectResult(
            hint=_EDIT_HINT,
            block_count=count_blocks(project.document),
        )


@mcp.tool(
    name="insert_page_break",
    description=(
        "Insert a page break as a body block. Without `block_index` it is "
        "appended; otherwise it is inserted at that zero-based position. "
        "After the requested batch of edits, call `finalize_project` once "
        "(not after each change)."
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
) -> ProjectResult:

    async with project.lock:

        project.document.add_page_break()

        if block_index is not None:
            _move_block(
                project.document,
                count_blocks(project.document) - 1,
                block_index,
            )

        _store.touch(user_id, project)

        return ProjectResult(
            hint=_EDIT_HINT,
            block_count=count_blocks(project.document),
        )


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
) -> BlocksResult:

    async with project.lock:
        blocks = list_block_infos(project.document)

    hint = (
        "The list position is the zero-based block index for `move_block` "
        "and `remove_blocks`."
    ) if blocks else (
        "No blocks yet. Call `insert_paragraph` or `insert_table`."
    )

    return BlocksResult(hint=hint, blocks=blocks)


@mcp.tool(
    name="move_block",
    description=(
        "Move a body block (paragraph or table) to a new position by "
        "zero-based index. Negative `to_index` counts from the end. Changes "
        "stay in memory only. After the requested batch of edits, call "
        "`finalize_project` once (not after each change)."
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
) -> BlocksResult:

    async with project.lock:

        _move_block(project.document, from_index, to_index)

        _store.touch(user_id, project)

        return BlocksResult(
            hint=_EDIT_HINT,
            blocks=list_block_infos(project.document),
        )


@mcp.tool(
    name="remove_blocks",
    description=(
        "Remove body blocks (paragraphs and tables) by zero-based index "
        "(from `list_blocks`). Indices refer to positions before removal; "
        "duplicates are ignored. Changes stay in memory only. After the "
        "requested batch of edits, call `finalize_project` once (not after "
        "each change)."
    ),
)
async def remove_blocks(
    indices: list[int] = Field(
        description="Zero-based block indices."
    ),
    user_id: str = TokenClaim("id"),
    project: Project = Depends(_get_project),
) -> ProjectResult:

    async with project.lock:

        drop_blocks(project.document, indices)

        _store.touch(user_id, project)

        return ProjectResult(
            hint=_EDIT_HINT,
            block_count=count_blocks(project.document),
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
    project: Project = Depends(_get_project),
) -> FinalizeResult:

    async with project.lock:

        out_name = f"{file_name}.docx"

        buffer = BytesIO()
        await to_thread(project.document.save, buffer)

        _store.touch(user_id, project)

        block_count = count_blocks(project.document)

    uploaded = await upload_file(
        file_name=out_name,
        data=buffer.getvalue(),
        content_type=DOCX_MIME,
        token=token.token,
    )

    return FinalizeResult(
        hint=_FINALIZE_HINT,
        file_name=out_name,
        block_count=block_count,
        owui_url=uploaded.download_url,
    )
