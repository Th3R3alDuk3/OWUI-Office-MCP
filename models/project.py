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
