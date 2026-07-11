from collections.abc import Callable
from io import BytesIO

from docx.document import Document as DocumentType
from docx.enum.style import WD_STYLE_TYPE
from docx.image.exceptions import (
    InvalidImageStreamError,
    UnexpectedEndOfFileError,
    UnrecognizedImageError,
)

from services.owui import download_file
from subservers._chart import render_chart
from subservers.docx import _ops

_MAX_TABLE_ROWS = 50
_MAX_TABLE_COLS = 10

SCRIPT_API = """
Run a Python script that builds or edits the current project. The script is
sandboxed: no imports, no file or network access. Available functions:

- add_paragraph(text: str, style: str | None = None,
                index: int | None = None) -> int
  Add a paragraph, optionally with a paragraph style from `styles` (see the
  `create_project` / `open_project` result); returns its zero-based block
  index. Without `index` it is appended.
- add_table(data: list[list[str]], style: str | None = None,
            index: int | None = None) -> int
  Add a table filled from row-major `data` (max 50x10), optionally with a
  table style from `styles`; returns its block index.
- add_page_break(index: int | None = None) -> int
- await add_image(file_id: str, width_cm: float | None = None,
                  index: int | None = None) -> int
  Insert an image the user attached in OpenWebUI (async: must be awaited).
  Without `width_cm` it is scaled to fit the page width. `file_id` is never
  invented.
- add_chart(kind: str, categories: list[str],
            series: dict[str, list[float]], title: str | None = None,
            width_cm: float | None = None, index: int | None = None) -> int
  Render a chart as an image block. `kind` is "bar", "line" or "pie";
  `series` maps each series name to one value per category. A pie chart
  takes exactly one series.
- edit_paragraph(index: int, text: str, style: str | None = None) -> None
  Replace the text (and optionally style) of an existing paragraph.
- list_blocks() -> list[dict]
  Body blocks in order (type, text); the list position is the block index.
- move_block(from_index: int, to_index: int) -> None
- remove_blocks(indices: list[int]) -> None

Plain Python works: loops, conditionals, f-strings, comprehensions,
functions, print. `print(...)` output and the last expression are returned
as `output`. If the script fails, fix it and call `run_script` again.

Example:

    add_paragraph("Bericht 2026", style="Heading 1")
    for name, value in [("Nord", "1,2 Mio"), ("Sued", "0,9 Mio")]:
        add_paragraph(name, style="Heading 2")
        add_paragraph(f"Umsatz: {value}")

After the requested batch of edits, call `finalize_project` once (not after
each change).
""".strip()


def script_functions(
    document: DocumentType,
    token: str,
) -> dict[str, Callable]:

    def _require_style(style: str | None, style_type: WD_STYLE_TYPE) -> None:
        try:
            document.styles.get_style_id(style, style_type)
        except (KeyError, ValueError):
            raise ValueError(
                f"Style '{style}' not found. Use a name from `styles`."
            ) from None

    def _place_last(index: int | None) -> int:
        # Move the block just appended to `index`; return its final position.
        if index is not None:
            _ops.move_block(
                document, _ops.count_blocks(document) - 1, index,
            )
            return index % _ops.count_blocks(document)
        return _ops.count_blocks(document) - 1

    def add_paragraph(
        text: str,
        style: str | None = None,
        index: int | None = None,
    ) -> int:
        _require_style(style, WD_STYLE_TYPE.PARAGRAPH)
        document.add_paragraph(text, style=style)
        return _place_last(index)

    def add_table(
        data: list[list[str]],
        style: str | None = None,
        index: int | None = None,
    ) -> int:

        rows = len(data)
        cols = max((len(row) for row in data), default=0)

        if not 0 < rows <= _MAX_TABLE_ROWS or not 0 < cols <= _MAX_TABLE_COLS:
            raise ValueError(
                f"Table must be between 1x1 and "
                f"{_MAX_TABLE_ROWS}x{_MAX_TABLE_COLS} cells."
            )

        _require_style(style, WD_STYLE_TYPE.TABLE)

        table = document.add_table(rows=rows, cols=cols)

        if style is not None:
            table.style = style

        for r, row in enumerate(data):
            for c, cell_text in enumerate(row):
                table.cell(r, c).text = str(cell_text)

        return _place_last(index)

    def add_page_break(index: int | None = None) -> int:
        document.add_page_break()
        return _place_last(index)

    async def add_image(
        file_id: str,
        width_cm: float | None = None,
        index: int | None = None,
    ) -> int:

        if width_cm is not None and not 0 < width_cm <= 30:
            raise ValueError("width_cm must be between 0 and 30.")

        file_content = await download_file(file_id=file_id, token=token)

        try:
            _ops.insert_picture(document, BytesIO(file_content), width_cm)
        except (
            InvalidImageStreamError,
            UnexpectedEndOfFileError,
            UnrecognizedImageError,
        ) as error:
            raise ValueError(
                f"File '{file_id}' is not a supported image."
            ) from error

        return _place_last(index)

    def add_chart(
        kind: str,
        categories: list[str],
        series: dict[str, list[float]],
        title: str | None = None,
        width_cm: float | None = None,
        index: int | None = None,
    ) -> int:

        if width_cm is not None and not 0 < width_cm <= 30:
            raise ValueError("width_cm must be between 0 and 30.")

        chart_png = render_chart(kind, categories, series, title)

        _ops.insert_picture(document, BytesIO(chart_png), width_cm)

        return _place_last(index)

    def edit_paragraph(
        index: int,
        text: str,
        style: str | None = None,
    ) -> None:

        _require_style(style, WD_STYLE_TYPE.PARAGRAPH)

        paragraph = _ops.get_paragraph(document, index)
        paragraph.text = text

        if style is not None:
            paragraph.style = style

    def list_blocks() -> list[dict]:
        return [
            info.model_dump()
            for info in _ops.list_block_infos(document)
        ]

    def move_block(from_index: int, to_index: int) -> None:
        _ops.move_block(document, from_index, to_index)

    def remove_blocks(indices: list[int]) -> None:
        _ops.remove_blocks(document, indices)

    return {
        "add_paragraph": add_paragraph,
        "add_table": add_table,
        "add_page_break": add_page_break,
        "add_image": add_image,
        "add_chart": add_chart,
        "edit_paragraph": edit_paragraph,
        "list_blocks": list_blocks,
        "move_block": move_block,
        "remove_blocks": remove_blocks,
    }
