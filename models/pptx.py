from asyncio import Lock
from dataclasses import dataclass, field

from pptx.presentation import Presentation as PresentationType
from pydantic import BaseModel, Field


@dataclass
class Project:
    presentation: PresentationType
    master_name: str | None = None
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


class MasterInfo(BaseModel):
    layouts: dict[str, LayoutInfo] = Field(
        description="Layout name -> its placeholders.",
    )


class SlideInfo(BaseModel):
    layout: str = Field(
        description="Layout name the slide is based on.",
    )
    text: str = Field(
        description="All text currently on the slide.",
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
    slide_count: int = Field(
        description="Current number of slides in the project.",
    )


class StartResult(ProjectResult):
    masters: dict[str, MasterInfo] = Field(
        description=(
            "Master name -> its layouts. `add_slide` only accepts layouts "
            "of the master selected via `set_master`."
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
