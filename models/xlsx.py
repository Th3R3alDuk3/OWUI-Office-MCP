from asyncio import Lock
from dataclasses import dataclass, field

from openpyxl.workbook import Workbook as WorkbookType
from pydantic import BaseModel, Field


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


class TemplatesResult(BaseModel):
    hint: str = Field(
        description=(
            "Suggested next step — guidance for the agent, not part of the data."
        ),
    )
    templates: list[str] = Field(
        description="Template file names for `create_project`.",
    )


class ProjectResult(BaseModel):
    hint: str = Field(
        description=(
            "Suggested next step — guidance for the agent, not part of the data."
        ),
    )
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


class FinalizeResult(ProjectResult):
    file_name: str = Field(
        description="Name of the uploaded file.",
    )
    owui_url: str = Field(
        description="OpenWebUI download URL of the uploaded file.",
    )
