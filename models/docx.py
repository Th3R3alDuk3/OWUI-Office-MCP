from asyncio import Lock
from dataclasses import dataclass, field

from docx.document import Document as DocumentType
from pydantic import BaseModel


@dataclass
class Project:
    document: DocumentType
    lock: Lock = field(default_factory=Lock)


class DownloadProjectResponse(BaseModel):
    filename: str
    block_count: int
    owui_url: str


class StyleInfo(BaseModel):
    type: str
    builtin: bool
