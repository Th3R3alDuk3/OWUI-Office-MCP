from asyncio import Lock
from dataclasses import dataclass, field

from openpyxl.workbook import Workbook as WorkbookType
from pydantic import BaseModel, Field

from models._base import ToolResult, UploadResult


@dataclass
class Project:
    workbook: WorkbookType
    lock: Lock = field(default_factory=Lock)


class SheetInfo(BaseModel):
    title: str = Field(
        description="Worksheet title; pass it to the read/write tools.",
    )
    rows: int = Field(
        description="Used rows (0 when the sheet is empty).",
    )
    cols: int = Field(
        description="Used columns (0 when the sheet is empty).",
    )


class CellInput(BaseModel):
    ref: str = Field(
        description="A1-style cell reference, e.g. `B2`."
    )
    value: bool | int | float | str | None = Field(
        default=None,
        description=(
            "Cell value. Numbers stay numeric, a string starting with `=` "
            "becomes a formula, and `null` clears the cell."
        ),
    )
    style: str | None = Field(
        default=None,
        description="Named cell style from `list_styles`. None = default.",
    )


class ProjectResult(ToolResult):
    sheet_count: int = Field(
        description="Current number of worksheets in the project.",
    )


class SheetsResult(ToolResult):
    sheets: list[SheetInfo] = Field(
        description="Worksheets in workbook order.",
    )


class StylesResult(ToolResult):
    styles: list[str] = Field(
        description="Named cell styles for `write_rows` / `write_cells`.",
    )


class WriteResult(ToolResult):
    cells: int = Field(
        description="Number of cells written.",
    )


class ReadResult(ToolResult):
    rows: list[list[str]] = Field(
        description=(
            "Used range, row-major; empty cells as empty strings, formulas "
            "as their formula string."
        ),
    )


class FinalizeResult(UploadResult):
    sheet_count: int = Field(
        description="Number of worksheets in the uploaded file.",
    )
