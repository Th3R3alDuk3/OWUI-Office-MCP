import re
from asyncio import Lock
from dataclasses import dataclass, field

from pptx.presentation import Presentation as PresentationType
from pydantic import BaseModel, Field, field_validator

from models._base import ToolResult, UploadResult

_BULLET_MARKER = re.compile(
    r"^[ \t]*(?:[-‐‑‒–—―−*+•∙‣⁃◦·▪▫■□●○◆◇❖]|\d+[.)])[ \t]+",
    re.MULTILINE,
)


@dataclass
class Project:
    presentation: PresentationType
    lock: Lock = field(default_factory=Lock)


class PlaceholderInfo(BaseModel):
    idx: int = Field(
        description="Placeholder id; target it in `insert_slide` / `edit_slide`.",
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


class PlaceholderText(BaseModel):
    idx: int = Field(
        description="Placeholder `idx` from `list_layouts`."
    )
    text: str = Field(
        description="Plain text for this placeholder."
    )

    @field_validator("text")
    @classmethod
    def validate_text(cls, value: str) -> str:
        return _BULLET_MARKER.sub("", value)


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


class MastersResult(ToolResult):
    masters: dict[int, str] = Field(
        description="Master index -> master name; pass the index to `list_layouts`.",
    )


class LayoutsResult(ToolResult):
    layouts: dict[str, LayoutInfo] = Field(
        description="Layout name -> its placeholders, for `insert_slide`.",
    )


class SlideResult(ToolResult):
    slide: SlideInfo = Field(
        description="The updated slide.",
    )


class SlidesResult(ToolResult):
    slides: list[SlideInfo] = Field(
        description=(
            "Slides in order; the list position is the zero-based slide index."
        ),
    )


class FinalizeResult(UploadResult):
    slide_count: int = Field(
        description="Number of slides in the uploaded file.",
    )
