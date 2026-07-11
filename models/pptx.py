from asyncio import Lock
from dataclasses import dataclass, field

from pptx.presentation import Presentation as PresentationType
from pydantic import BaseModel, Field

from models._base import ToolResult, UploadResult


@dataclass
class Project:
    presentation: PresentationType
    lock: Lock = field(default_factory=Lock)


class PlaceholderInfo(BaseModel):
    idx: int = Field(
        description="Placeholder id; target it with `fill` in `run_script`.",
    )
    name: str = Field(
        description="Human-readable placeholder name.",
    )
    type: str = Field(
        description="Placeholder type, e.g. TITLE or BODY.",
    )


class LayoutInfo(BaseModel):
    placeholders: list[PlaceholderInfo] = Field(
        description="Text placeholders the layout offers.",
    )


class SlideInfo(BaseModel):
    layout: str = Field(
        description="Layout name the slide is based on.",
    )
    text: str = Field(
        description="All text currently on the slide.",
    )


class ProjectResult(ToolResult):
    slide_count: int = Field(
        description="Current number of slides in the project.",
    )


class StartResult(ProjectResult):
    layouts: dict[str, LayoutInfo] = Field(
        description=(
            "Layout name -> its placeholders, across all masters. Everything "
            "`run_script` may use."
        ),
    )


class ScriptResult(ProjectResult):
    output: str = Field(
        description="Printed output and last expression of the script.",
    )


class FinalizeResult(UploadResult):
    slide_count: int = Field(
        description="Number of slides in the uploaded file.",
    )
