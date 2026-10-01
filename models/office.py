from json import loads
from typing import Annotated, Literal

from pydantic import (
    BaseModel,
    BeforeValidator,
    ConfigDict,
    Field,
    JsonValue,
    StringConstraints,
)

Format = Literal["pptx", "docx", "xlsx"]


class Command(BaseModel):
    # OfficeCLI validates the rest and answers with its own errors.
    model_config = ConfigDict(extra="allow")

    command: Annotated[str, StringConstraints(to_lower=True)] = Field(
        description=(
            "OfficeCLI verb, e.g. add, set, remove, move, swap, import, get, "
            "query, view, validate, raw. Its arguments are sibling fields as in "
            "an OfficeCLI batch: parent, path, type, props, selector, mode, ..."
        ),
    )


def _listed(
    value: object,
) -> object:

    # The array may arrive as a JSON string or a single command.
    if isinstance(value, str):
        value = loads(value)

    return [value] if isinstance(value, dict) else value


# A plain alias, so the schema inlines the array instead of a $ref.
Commands = Annotated[list[Command], BeforeValidator(_listed)]


class Layout(BaseModel):
    index: int = Field(
        description=(
            "Layout index in PowerPoint's order; pass it as the slide's "
            "`layout` prop."
        ),
    )
    name: str = Field(
        description="Layout name, for orientation only.",
    )
    slots: dict[str, dict[str, str]] = Field(
        description=(
            "Placeholder slots as `phType:idx`, e.g. `title:0` or `body:1`, "
            "with their `x`, `y`, `width` and `height` on the slide; add them "
            "as placeholders with that `phType` and `idx`, or use the position "
            "to fit a chart or picture into the slot."
        ),
    )


class SlideMaster(BaseModel):
    index: int = Field(
        description=(
            "Slide master index in PowerPoint's order; pass it as "
            "`slide_master`."
        ),
    )
    name: str = Field(
        description="Slide master name.",
    )
    layouts: list[Layout] = Field(
        description="The slide master's layouts.",
    )


class PptxInventory(BaseModel):
    slide_size: str = Field(
        description="Slide width x height; size charts and images from it.",
    )
    slide_masters: list[SlideMaster] = Field(
        description="The slide master the project is bound to, with its layouts.",
    )


class DocxStyles(BaseModel):
    custom_paragraph: list[str] = Field(
        description="The template's own paragraph styles (its corporate design).",
    )
    custom_table: list[str] = Field(
        description="The template's own table styles.",
    )
    builtin_paragraph: list[str] = Field(
        description="Word built-in paragraph styles defined in the file.",
    )
    builtin_table: list[str] = Field(
        description="Word built-in table styles defined in the file.",
    )


class DocxInventory(BaseModel):
    styles: DocxStyles = Field(
        description="Style IDs for the `style` prop; prefer the custom ones.",
    )


class Sheet(BaseModel):
    name: str = Field(
        description="Sheet name; paths start with `/<name>`.",
    )
    rows: int = Field(
        description="Rows holding content (0 when the sheet is empty).",
    )


class XlsxInventory(BaseModel):
    sheets: list[Sheet] = Field(
        description="Worksheets in workbook order.",
    )


type AnyInventory = PptxInventory | DocxInventory | XlsxInventory


class Template(BaseModel):
    name: str = Field(
        description="Pass it as `template_name` to `start_project`.",
    )
    format: Format = Field(
        description="Document format the template produces.",
    )
    slide_masters: dict[int, str] | None = Field(
        default=None,
        description=(
            "PPTX only: slide master index -> name; pass the index as "
            "`slide_master` to `start_project`."
        ),
    )


class OfficeResult(BaseModel):
    hint: str = Field(
        description="Your next step, not part of the data.",
    )


class TemplatesResult(OfficeResult):
    templates: list[Template] = Field(
        description="Stored templates.",
    )


class StartResult(OfficeResult):
    project_id: str = Field(
        description=(
            "Pass it to `run_commands`, `preview_project` and `export_project`."
        ),
    )
    format: Format = Field(
        description="The project's document format.",
    )
    inventory: AnyInventory = Field(
        description="Everything the design offers; use only these.",
    )


class ReferenceResult(OfficeResult):
    reference: JsonValue = Field(
        description=(
            "OfficeCLI's reference JSON for an element, or a plain-text "
            "overview for a verb."
        ),
    )


class CommandsResult(OfficeResult):
    results: list[JsonValue] = Field(
        description="OfficeCLI output per command, in command order.",
    )
    warnings: list[str] = Field(
        description=(
            "OfficeCLI advisories; unsupported props show up in `results`."
        ),
    )


class ExportResult(OfficeResult):
    file_name: str = Field(
        description="Name of the uploaded file.",
    )
    owui_url: str = Field(
        description="OpenWebUI download URL of the uploaded file.",
    )
