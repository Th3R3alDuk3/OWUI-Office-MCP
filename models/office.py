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


class Template(BaseModel):
    name: str = Field(
        description="Template file name.",
    )
    format: Format = Field(
        description="Document format the template produces.",
    )
    slide_masters: dict[int, str] | None = Field(
        description="PPTX only: slide master index -> name.",
    )


class TemplateList(BaseModel):
    templates: list[Template] = Field(
        description="Stored templates.",
    )


class Layout(BaseModel):
    index: int = Field(
        description=(
            "Layout index in PowerPoint's order, counted across all slide "
            "masters."
        ),
    )
    name: str = Field(
        description="Layout name.",
    )
    slots: dict[str, dict[str, str]] = Field(
        description=(
            "Placeholder slots by `phType:idx`, e.g. `title:0` or `body:1`, "
            "with their `x`, `y`, `width` and `height` on the slide."
        ),
    )


class SlideMaster(BaseModel):
    index: int = Field(
        description="Slide master index in PowerPoint's order.",
    )
    name: str = Field(
        description="Slide master name.",
    )
    layouts: list[Layout] = Field(
        description="The slide master's layouts.",
    )


class PptxInventory(BaseModel):
    slide_size: str = Field(
        description="Slide width x height.",
    )
    slide_masters: list[SlideMaster] = Field(
        description="The slide master the project is bound to, with its layouts.",
    )


class DocxInventory(BaseModel):
    custom_paragraph_styles: list[str] = Field(
        description="The file's own paragraph styles (its corporate design).",
    )
    custom_table_styles: list[str] = Field(
        description="The file's own table styles.",
    )
    builtin_paragraph_styles: list[str] = Field(
        description="Word built-in paragraph styles defined in the file.",
    )
    builtin_table_styles: list[str] = Field(
        description="Word built-in table styles defined in the file.",
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


class StartResult(BaseModel):
    hint: str = Field(
        description="Guidance for the next step.",
    )
    project_id: str = Field(
        description="ID of the new project.",
    )
    format: Format = Field(
        description="The project's document format.",
    )
    inventory: AnyInventory = Field(
        description=(
            "What the bound design offers: slide master and layouts, styles or "
            "sheets."
        ),
    )


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


def _as_list(
    value: object,
) -> object:

    # The array may arrive as a JSON string or a single command.
    if isinstance(value, str):
        value = loads(value)

    return [value] if isinstance(value, dict) else value


# A plain alias, so the schema inlines the array instead of a $ref.
Commands = Annotated[list[Command], BeforeValidator(_as_list)]


class CommandsResult(BaseModel):
    results: list[JsonValue] = Field(
        description="OfficeCLI output per command, in command order.",
    )
    warnings: list[str] = Field(
        description=(
            "OfficeCLI advisories; unsupported props show up in `results`."
        ),
    )


class ExportFile(BaseModel):
    id: str = Field(
        description="OpenWebUI file ID.",
    )
    name: str = Field(
        description="Name of the uploaded file.",
    )
    path: str = Field(
        description="Relative OpenWebUI download path.",
    )
