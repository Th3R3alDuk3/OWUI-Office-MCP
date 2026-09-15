from typing import Annotated, Literal

from pydantic import BaseModel, ConfigDict, Field, JsonValue, StringConstraints

from models.inventory import DocxInventory, PptxInventory, XlsxInventory

Format = Literal["pptx", "docx", "xlsx"]
DesignMode = Literal["template", "custom"]
# OfficeCLI treats element types and prop names case-insensitively.
Lowercase = Annotated[str, StringConstraints(to_lower=True)]


class Command(BaseModel):
    model_config = ConfigDict(extra="forbid")

    command: Literal[
        "add", "set", "remove", "move", "swap", "import",
        "get", "query", "view", "validate",
    ]
    path: str | None = Field(
        default=None,
        description="Target of set, remove, move, get.",
    )
    parent: str | None = Field(
        default=None,
        description="Parent that add or import writes into.",
    )
    type: Lowercase | None = Field(
        default=None,
        description="Element type that add creates.",
    )
    from_: str | None = Field(
        default=None,
        alias="from",
        description="Existing element that add copies.",
    )
    index: int | None = Field(
        default=None,
        ge=0,
        description="0-based position for add and move.",
    )
    after: str | None = Field(
        default=None,
        description="Add or move after this element.",
    )
    before: str | None = Field(
        default=None,
        description="Add or move before this element.",
    )
    to: str | None = Field(
        default=None,
        description="Target parent of a move.",
    )
    path2: str | None = Field(
        default=None,
        description="Second element of a swap.",
    )
    props: dict[Lowercase, str | int | float | bool] | None = Field(
        default=None,
        description="Property name -> value.",
    )
    selector: str | None = Field(
        default=None,
        description="Filter of a query, e.g. `slide`.",
    )
    text: str | None = Field(
        default=None,
        description="Inline CSV of an import.",
    )
    mode: Literal["text", "annotated", "outline", "stats", "issues"] | None = Field(
        default=None,
        description="Mode of a view; `text` when omitted.",
    )
    depth: int | None = Field(
        default=None,
        ge=0,
        le=4,
        description="Child levels that get returns.",
    )


class Template(BaseModel):
    name: str = Field(
        description="Pass it as `template_name` to `create_project`.",
    )
    format: Format = Field(
        description="Document format the template produces.",
    )
    masters: dict[int, str] | None = Field(
        default=None,
        description=(
            "PPTX only: slide master index -> name; pass the index as "
            "`master` to `create_project`."
        ),
    )


class ToolResult(BaseModel):
    hint: str = Field(
        description=(
            "Suggested next step — guidance for the agent, not part of the data."
        ),
    )


class TemplatesResult(ToolResult):
    templates: list[Template] = Field(
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
