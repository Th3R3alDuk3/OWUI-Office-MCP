from asyncio import Lock
from dataclasses import dataclass, field

from docx.document import Document as DocumentType
from pydantic import BaseModel, Field


@dataclass
class Project:
    document: DocumentType
    lock: Lock = field(default_factory=Lock)


class StyleGroups(BaseModel):
    custom_paragraph: list[str] = Field(
        description=(
            "The template's own paragraph styles — prefer these; they "
            "carry the corporate design."
        ),
    )
    custom_table: list[str] = Field(
        description=(
            "The template's own table styles — prefer these; they carry "
            "the corporate design."
        ),
    )
    builtin_paragraph: list[str] = Field(
        description="Word built-in paragraph styles defined in the file.",
    )
    builtin_table: list[str] = Field(
        description="Word built-in table styles defined in the file.",
    )


class BlockInfo(BaseModel):
    type: str = Field(
        description="Block type: `paragraph` or `table`.",
    )
    text: str = Field(
        description=(
            "Paragraph text (`[image]` for image paragraphs, "
            "`[section break]` for section-break paragraphs), or a size "
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
    styles: StyleGroups = Field(
        description=(
            "Style names `run_script` may use, grouped by kind. Only these "
            "exist in the file."
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
