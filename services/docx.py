from pathlib import Path
from xml.etree import ElementTree

from fastmcp.exceptions import ToolError

from models.office import DocxInventory, DocxStyles
from services.officecli import run_batch

FORMAT = "docx"
MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
INVENTORY = DocxInventory


START_HINT = """
Use style IDs from `inventory.styles`. Build in steps: a few blocks per
batch, then `view outline`, then continue. New blocks go into `/body` and are
appended unless `index`, `after` or `before` say otherwise; paths are
1-based, `index` is 0-based. Tables take `data` "H1,H2;r1c1,r1c2", charts
`chartType`, `categories` "Q1,Q2" and `data` "Revenue:1,2;Costs:3,4", images
`src` "file:<file_id>" from attached files. Stay within the design: change
colors, fonts and other appearance only when the user asks. `get_reference`
documents elements and props.
""".strip()

_NAMESPACE = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}


async def prepare(
    file: Path,
) -> None:

    (document,) = await run_batch(file, [
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
        await run_batch(file, commands)


async def inventory(
    file: Path,
) -> DocxInventory:

    (styles,) = await run_batch(file, [
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

    return DocxInventory(styles=groups)
