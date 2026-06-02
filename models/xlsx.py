from asyncio import Lock
from dataclasses import dataclass, field

from openpyxl.workbook import Workbook as WorkbookType
from pydantic import BaseModel, Field


@dataclass
class Project:
    workbook: WorkbookType
    lock: Lock = field(default_factory=Lock)


class SheetInfo(BaseModel):
    title: str
    rows: int
    cols: int


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


class ProjectResponse(BaseModel):
    file_name: str
    sheet_count: int
    owui_url: str
