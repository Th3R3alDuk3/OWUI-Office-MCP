from asyncio import Lock
from dataclasses import dataclass, field

from docx.document import Document as DocumentType
from pydantic import BaseModel, Field

from models._base import ToolResult, UploadResult


@dataclass
class Project:
    document: DocumentType
    lock: Lock = field(default_factory=Lock)


class StyleInfo(BaseModel):
    type: str = Field(
        description="Style type: paragraph or table.",
    )
    builtin: bool = Field(
        description="Whether the style is a Word built-in.",
    )


class BlockInfo(BaseModel):
    type: str = Field(
        description="Block type: `paragraph` or `table`.",
    )
    text: str = Field(
        description="Paragraph text, or a size preview for tables.",
    )


class ProjectResult(ToolResult):
    block_count: int = Field(
        description="Current number of body blocks in the project.",
    )


class StylesResult(ToolResult):
    styles: dict[str, StyleInfo] = Field(
        description=(
            "Style name -> type and builtin flag, for `insert_paragraph` / "
            "`insert_table`."
        ),
    )


class BlocksResult(ToolResult):
    blocks: list[BlockInfo] = Field(
        description=(
            "Body blocks in order; the list position is the zero-based block index."
        ),
    )


class FinalizeResult(UploadResult):
    block_count: int = Field(
        description="Number of body blocks in the uploaded file.",
    )
