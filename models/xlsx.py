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
        description="Worksheet title; pass it to the `run_script` functions.",
    )
    rows: int = Field(
        description="Used rows (0 when the sheet is empty).",
    )
    cols: int = Field(
        description="Used columns (0 when the sheet is empty).",
    )


class ProjectResult(ToolResult):
    sheet_count: int = Field(
        description="Current number of worksheets in the project.",
    )


class StartResult(ProjectResult):
    sheets: list[SheetInfo] = Field(
        description="Worksheets in workbook order.",
    )
    styles: list[str] = Field(
        description=(
            "Named cell styles of the template. Everything `run_script` "
            "may use."
        ),
    )


class ScriptResult(ProjectResult):
    output: str = Field(
        description="Printed output and last expression of the script.",
    )


class FinalizeResult(UploadResult):
    sheet_count: int = Field(
        description="Number of worksheets in the uploaded file.",
    )
