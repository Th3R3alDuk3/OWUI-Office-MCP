from pathlib import Path
from xml.etree import ElementTree

from fastmcp.exceptions import ToolError

from models.inventory import DocxInventory, DocxStyles
from office import _officecli

FORMAT = "docx"
MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
CONTENT_TYPE = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml."
    "document.main+xml"
)
INVENTORY = DocxInventory

ELEMENT_TYPES = {"paragraph", "table", "pagebreak", "chart", "picture", "comment"}
CONTENT_PROPS = {
    "text", "style", "data", "charttype", "categories", "title", "legend",
    "datalabels", "src", "width", "height", "author", "initials",
}
APPEARANCE_PROPS = {
    "bold", "italic", "underline", "color", "size", "font", "highlight",
    "alignment", "shading",
}
PROTECTED_PATHS = ("/styles", "/theme", "/numbering")

_NAMESPACE = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

START_HINT = """
Build with `run_commands`, using style IDs from `inventory.styles`, e.g.:
- paragraph: {"command": "add", "parent": "/body", "type": "paragraph",
  "props": {"text": "Report 2026", "style": "Title"}}
- table: type `table`, props `data` "Region,Revenue;North,1.2" and `style`.
- page break: type `pagebreak`.
- chart (native Word chart): type `chart`, props `chartType` (bar, line,
  pie, ...), `categories` "Q1,Q2", `data` "Revenue:1,2;Costs:3,4".
- image: type `picture`, props `src` "file:<file_id>" and `width` "12cm".
- review comment: {"command": "add", "parent": "/body/p[2]",
  "type": "comment", "props": {"text": "...", "author": "..."}}
Colors, fonts and other appearance props need `design_mode="custom"` and an
explicit user request. Without `index`, `after` or `before`, blocks are
appended. Paths are 1-based, `index` is 0-based.
""".strip()


async def prepare(
    file: Path,
) -> None:

    (document,) = await _officecli.run(file, [
        {"command": "raw", "part": "/document"},
    ])

    body = ElementTree.fromstring(document["output"]).find("w:body", _NAMESPACE)

    if body is None:
        raise ToolError("The Word document has no body. Use another file.")

    sections = [
        block for block in body
        if block.find("w:pPr/w:sectPr", _NAMESPACE) is not None
    ]
    commands: list[dict] = []

    if any(
        block not in sections and block.tag != f"{{{_NAMESPACE['w']}}}sectPr"
        for block in body
    ):
        commands.append({
            "command": "raw-set", "part": "/document", "action": "remove",
            "xpath": "//w:body/*[not(self::w:sectPr) and not(w:pPr/w:sectPr)]",
        })

    # Keep section properties, not text, fields or images in their paragraphs.
    if any(
        child.tag != f"{{{_NAMESPACE['w']}}}pPr"
        for paragraph in sections for child in paragraph
    ):
        commands.append({
            "command": "raw-set", "part": "/document", "action": "remove",
            "xpath": "//w:body/w:p[w:pPr/w:sectPr]/*[not(self::w:pPr)]",
        })

    if commands:
        await _officecli.run(file, commands)


async def inventory(
    file: Path,
) -> DocxInventory:

    root, styles = await _officecli.run(file, [
        {"command": "get", "path": "/", "depth": 0},
        {"command": "query", "selector": "style"},
    ])

    groups = DocxStyles(
        custom_paragraph=[],
        custom_table=[],
        builtin_paragraph=[],
        builtin_table=[],
    )

    for style in styles["output"]["results"]:

        properties = style["format"]

        match properties.get("type"), bool(properties.get("customStyle")):
            case "paragraph", True:
                groups.custom_paragraph.append(properties["id"])
            case "table", True:
                groups.custom_table.append(properties["id"])
            case "paragraph", False:
                groups.builtin_paragraph.append(properties["id"])
            case "table", False:
                groups.builtin_table.append(properties["id"])

    return DocxInventory(
        theme=_officecli.theme(root),
        styles=groups,
    )


def bind(
    inventory: DocxInventory,
    master: int | None,
) -> DocxInventory:

    if master is not None:
        raise ToolError("`master` applies to pptx designs only.")

    return inventory


def adapt(
    command: dict,
    inventory: DocxInventory,
) -> dict:

    props = command.get("props", {})
    style = props.get("style")

    if style is not None and style not in {
        known for group in inventory.styles.model_dump().values() for known in group
    }:
        raise ToolError(
            f"Style '{style}' not found. Use a style ID from `styles`."
        )

    return command
