from asyncio import Lock
from dataclasses import dataclass, field

from pptx.presentation import Presentation as PresentationType
from pydantic import BaseModel


@dataclass
class Project:
    presentation: PresentationType
    lock: Lock = field(default_factory=Lock)


class DownloadProjectResponse(BaseModel):
    filename: str
    slide_count: int
    owui_url: str


class PlaceholderInfo(BaseModel):
    idx: int
    name: str
    type: str


class LayoutInfo(BaseModel):
    placeholders: list[PlaceholderInfo]
