from typing import Literal

from pydantic import BaseModel, Field, JsonValue

from models.inventory import DocxInventory, PptxInventory, XlsxInventory

Format = Literal["pptx", "docx", "xlsx"]
DesignMode = Literal["template", "custom"]


class Template(BaseModel):
    name: str = Field(
        description="Pass it as `template_name` to `create_project`.",
    )
    format: Format = Field(
        description="Document format the template produces.",
    )


class PptxTemplate(Template):
    masters: dict[int, str] = Field(
        description=(
            "Slide master index -> name; pass the index as `master` to "
            "`create_project`."
        ),
    )


class ToolResult(BaseModel):
    hint: str = Field(
        description=(
            "Suggested next step — guidance for the agent, not part of the data."
        ),
    )


class TemplatesResult(ToolResult):
    templates: list[PptxTemplate | Template] = Field(
        description="Stored templates.",
    )


class StartResult(ToolResult):
    project_id: str = Field(
        description="Pass it to `run_commands` and `export_project`.",
    )
    format: Format = Field(
        description="The project's document format.",
    )
    inventory: PptxInventory | DocxInventory | XlsxInventory = Field(
        description="Everything the design offers; use only these.",
    )


class CommandsResult(ToolResult):
    results: list[JsonValue] = Field(
        description="OfficeCLI output per command, in command order.",
    )
    warnings: list[str] = Field(
        description="OfficeCLI advisories, e.g. ignored props.",
    )


class ReferenceResult(ToolResult):
    reference: JsonValue = Field(
        description=(
            "An element's paths and the props `run_commands` accepts, or a "
            "plain-text overview for a verb."
        ),
    )


class ExportResult(ToolResult):
    file_name: str = Field(
        description="Name of the uploaded file.",
    )
    owui_url: str = Field(
        description="OpenWebUI download URL of the uploaded file.",
    )
    warnings: list[str] = Field(
        description=(
            "Validation errors and layout issues the design did not already "
            "have."
        ),
    )
