from asyncio import Lock
from dataclasses import dataclass, field

from pptx import Presentation
from pydantic import BaseModel


@dataclass
class Project:
    presentation: Presentation
    lock: Lock = field(default_factory=Lock)


class DownloadProjectResponse(BaseModel):
    filename: str
    slide_count: int
    owui_url: str
