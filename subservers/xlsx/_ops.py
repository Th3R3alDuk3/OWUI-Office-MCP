from io import BytesIO
from pathlib import Path

from fastmcp.utilities.logging import get_logger
from openpyxl import load_workbook
from openpyxl.drawing.image import Image
from openpyxl.utils.cell import (
    coordinate_from_string,
    coordinate_to_tuple,
    get_column_letter,
)
from openpyxl.utils.exceptions import CellCoordinatesException
from openpyxl.workbook import Workbook as WorkbookType
from openpyxl.worksheet.worksheet import Worksheet

from models.xlsx import SheetInfo

logger = get_logger(__name__)


_MAX_COLUMN_WIDTH = 60.0

# Display size of inserted pictures in pixels (16:9).
_PICTURE_WIDTH = 640
_PICTURE_HEIGHT = 360


def list_template_names(
    templates_dir: Path,
) -> list[str]:

    if not templates_dir.is_dir():
        raise RuntimeError(f"Templates directory not found: {templates_dir}")

    template_names: list[str] = []

    for file_path in sorted(templates_dir.glob("*.xlsx")):
        try:
            load_workbook(file_path, read_only=True).close()
        except Exception as error:
            logger.warning(f"Skipping template {file_path.name}: {error}")
        else:
            template_names.append(file_path.name)

    return template_names


def count_sheets(
    workbook: WorkbookType,
) -> int:
    return len(workbook.worksheets)


def _sheet_extent(
    worksheet: Worksheet,
) -> tuple[int, int]:

    rows, cols = worksheet.max_row, worksheet.max_column

    # openpyxl never reports below 1x1, so treat a lone empty A1 as empty.
    if rows == 1 and cols == 1 and worksheet["A1"].value is None:
        return 0, 0

    return rows, cols


def list_sheet_infos(
    workbook: WorkbookType,
) -> list[SheetInfo]:

    sheet_infos: list[SheetInfo] = []

    for worksheet in workbook.worksheets:

        rows, cols = _sheet_extent(worksheet)

        sheet_infos.append(
            SheetInfo(title=worksheet.title, rows=rows, cols=cols)
        )

    return sheet_infos


def _get_sheet(
    workbook: WorkbookType,
    sheet_title: str,
) -> Worksheet:
    try:
        return workbook[sheet_title]
    except KeyError:
        raise ValueError(
            f"Sheet '{sheet_title}' not found."
        ) from None


def read_sheet(
    workbook: WorkbookType,
    sheet_title: str,
) -> list[list[str]]:

    worksheet = _get_sheet(workbook, sheet_title)

    rows: list[list[str]] = []

    for row in worksheet.iter_rows(values_only=True):
        rows.append(["" if value is None else str(value) for value in row])

    return rows


def write_cell(
    workbook: WorkbookType,
    sheet_title: str,
    ref: str,
    value: bool | int | float | str | None,
    style: str | None,
) -> None:

    worksheet = _get_sheet(workbook, sheet_title)

    try:
        coordinate_from_string(ref)
    except CellCoordinatesException:
        raise ValueError(f"Invalid cell reference '{ref}'.") from None

    if style is not None and style not in workbook.named_styles:
        raise ValueError(f"Style '{style}' not found.")

    cell = worksheet[ref]
    cell.value = value

    if style is not None:
        cell.style = style


def write_rows(
    workbook: WorkbookType,
    sheet_title: str,
    rows: list[list[bool | int | float | str | None]],
    anchor: str,
    style: str | None,
) -> None:

    worksheet = _get_sheet(workbook, sheet_title)

    try:
        start_row, start_col = coordinate_to_tuple(anchor)
    except (CellCoordinatesException, ValueError):
        raise ValueError(f"Invalid anchor reference '{anchor}'.") from None

    if style is not None and style not in workbook.named_styles:
        raise ValueError(f"Style '{style}' not found.")

    for row_offset, row in enumerate(rows):
        for col_offset, value in enumerate(row):

            cell = worksheet.cell(
                row=start_row + row_offset,
                column=start_col + col_offset,
            )
            cell.value = value

            if style is not None:
                cell.style = style


def insert_picture(
    workbook: WorkbookType,
    sheet_title: str,
    image: BytesIO,
    anchor: str,
) -> None:

    worksheet = _get_sheet(workbook, sheet_title)

    try:
        coordinate_from_string(anchor)
    except CellCoordinatesException:
        raise ValueError(f"Invalid anchor reference '{anchor}'.") from None

    picture = Image(image)
    picture.width = _PICTURE_WIDTH
    picture.height = _PICTURE_HEIGHT

    worksheet.add_image(picture, anchor)


def clear_workbook(
    workbook: WorkbookType,
) -> None:

    keep = workbook.create_sheet()

    for worksheet in list(workbook.worksheets):
        if worksheet is not keep:
            workbook.remove(worksheet)

    keep.title = "Sheet"


def add_sheet(
    workbook: WorkbookType,
    title: str,
    index: int | None,
) -> None:

    if title in workbook.sheetnames:
        raise ValueError(f"Sheet '{title}' already exists.")

    workbook.create_sheet(title=title, index=index)


def remove_sheets(
    workbook: WorkbookType,
    titles: list[str],
) -> None:

    targets = [_get_sheet(workbook, title) for title in dict.fromkeys(titles)]

    if len(targets) >= len(workbook.worksheets):
        raise ValueError(
            "Cannot remove every sheet; a workbook needs at least one."
        )

    for worksheet in targets:
        workbook.remove(worksheet)


def move_sheet(
    workbook: WorkbookType,
    title: str,
    to_index: int,
) -> None:

    try:
        current = workbook.sheetnames.index(title)
    except ValueError:
        raise ValueError(
            f"Sheet '{title}' not found."
        ) from None

    count = len(workbook.sheetnames)

    if not -count <= to_index < count:
        raise ValueError(f"Target index {to_index} out of range.")

    workbook.move_sheet(title, offset=to_index % count - current)


def autofit_columns(
    workbook: WorkbookType,
) -> None:

    for worksheet in workbook.worksheets:
        for index, column in enumerate(worksheet.iter_cols(), start=1):

            lengths = [
                len(line)
                for cell in column
                if cell.value is not None
                for line in str(cell.value).splitlines()
            ]

            if lengths:
                letter = get_column_letter(index)
                worksheet.column_dimensions[letter].width = min(
                    max(lengths) + 2.0, _MAX_COLUMN_WIDTH,
                )
