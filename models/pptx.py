import re
from asyncio import Lock
from dataclasses import dataclass, field

from pptx.presentation import Presentation as PresentationType
from pydantic import BaseModel, Field, field_validator


_BULLET_MARKER = re.compile(
    r"^[ \t]*(?:[-‐‑‒–—―−*+•∙‣⁃◦·▪▫■□●○◆◇❖]|\d+[.)])[ \t]+",
    re.MULTILINE,
)


@dataclass
class Project:
    presentation: PresentationType
    lock: Lock = field(default_factory=Lock)


class PlaceholderInfo(BaseModel):
    idx: int
    name: str
    type: str


class LayoutInfo(BaseModel):
    placeholders: list[PlaceholderInfo]


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
    layout: str
    text: str


class ProjectResponse(BaseModel):
    file_name: str
    slide_count: int
    owui_url: str
