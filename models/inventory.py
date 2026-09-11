from pydantic import BaseModel, Field


class Theme(BaseModel):
    colors: dict[str, str] = Field(
        description="Theme color name (e.g. `accent1`) -> hex value.",
    )
    fonts: list[str] = Field(
        description="Theme heading and body fonts.",
    )


class Inventory(BaseModel):
    theme: Theme = Field(
        description=(
            "The design's palette and fonts; for `design_mode=\"custom\"`, "
            "prefer theme color names over hex values."
        ),
    )


class Layout(BaseModel):
    index: int = Field(
        description="Global layout index; pass it as the slide's `layout` prop.",
    )
    name: str = Field(
        description="Layout name, for orientation only.",
    )
    slots: list[str] = Field(
        description=(
            "Placeholder slots as `phType` or `phType:idx`, e.g. `title` or "
            "`body:1`; add them to the slide as placeholders with that `idx`."
        ),
    )


class Master(BaseModel):
    index: int = Field(
        description="Master index for the `master` parameter.",
    )
    name: str = Field(
        description="Master name.",
    )
    layouts: list[Layout] = Field(
        description="The master's layouts.",
    )


class PptxInventory(Inventory):
    slide_size: str = Field(
        description="Slide width x height; size charts and images from it.",
    )
    masters: list[Master] = Field(
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


class DocxInventory(Inventory):
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


class XlsxInventory(Inventory):
    sheets: list[Sheet] = Field(
        description="Worksheets in workbook order.",
    )
    styles: dict[str, int] = Field(
        description=(
            "Named cell style -> cell format index. Pass the name as `style` "
            "in a `set` command of its own."
        ),
    )
