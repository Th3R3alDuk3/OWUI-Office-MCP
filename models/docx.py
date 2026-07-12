from asyncio import Lock
from dataclasses import dataclass, field

from docx.document import Document as DocumentType
from pydantic import BaseModel, Field


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
        description=(
            "Paragraph text (`[image]` for image paragraphs), or a size "
            "preview for tables."
        ),
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
    block_count: int = Field(
        description="Current number of body blocks in the project.",
    )


class StartResult(ProjectResult):
    styles: dict[str, StyleInfo] = Field(
        description=(
            "Style name -> type and builtin flag, across the whole template. "
            "Everything `run_script` may use."
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
