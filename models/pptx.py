from asyncio import Lock
from dataclasses import dataclass, field

from pptx.presentation import Presentation as PresentationType
from pydantic import BaseModel, Field


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
    idx: int = Field(description="Placeholder `idx` from `list_layouts`.")
    text: str = Field(description="Text for this placeholder.")


class SlideInfo(BaseModel):
    layout: str
    text: str


class DownloadProjectResponse(BaseModel):
    filename: str
    slide_count: int
    owui_url: str
