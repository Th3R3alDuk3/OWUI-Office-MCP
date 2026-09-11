import re
from copy import deepcopy
from itertools import product
from pathlib import Path
from string import ascii_uppercase
from xml.etree import ElementTree

from fastmcp.exceptions import ToolError

from models.inventory import Sheet, XlsxInventory
from office import _officecli

FORMAT = "xlsx"
MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
CONTENT_TYPE = (
    "application/vnd.openxmlformats-officedocument.spreadsheetml."
    "sheet.main+xml"
)
INVENTORY = XlsxInventory

ELEMENT_TYPES = {"sheet", "chart", "picture", "comment"}
CONTENT_PROPS = {
    "value", "formula", "style", "name", "ref", "text", "author", "charttype",
    "datarange", "categories", "data", "title", "legend", "datalabels",
    "anchor", "src", "width", "height", "autofit", "startcell", "freeze",
}
APPEARANCE_PROPS = {
    "bold", "italic", "underline", "font.color", "font.size", "font.name",
    "fill", "numfmt", "alignment.horizontal", "border.all",
}
PROTECTED_PATHS = ()

START_HINT = """
The design's sheets, values and formulas are kept; `inventory.sheets` counts
their rows. Build with `run_commands`, e.g.:
- block of values: {"command": "import", "parent": "/Sheet1",
  "text": "Region,Revenue\\nNorth,1.2\\nSouth,0.9",
  "props": {"startCell": "A1"}} — CSV; numbers stay numeric, `=...` is a
  formula.
- one cell: {"command": "set", "path": "/Sheet1/B4",
  "props": {"formula": "SUM(B2:B3)"}} (or `value`).
- named style from `inventory.styles`, alone in its own `set` and after the
  cells hold values: {"command": "set", "path": "/Sheet1/A1:B1",
  "props": {"style": "Header"}}
- chart: {"command": "add", "parent": "/Sheet1", "type": "chart",
  "props": {"chartType": "bar", "dataRange": "A1:B3", "anchor": "D2:J18"}}
  — linked to the cells.
- image: type `picture`, props `src` "file:<file_id>"; cell note: type
  `comment`, props `ref` "B2" and `text`; sheet: {"command": "add",
  "parent": "/", "type": "sheet", "props": {"name": "Q2"}}.
- fit widths of new columns: {"command": "set", "path": "/Sheet1/col[B]",
  "props": {"autofit": "true"}}; authored widths stay as they are.
Direct formatting (bold, fill, number format, ...) needs
`design_mode="custom"` and an explicit user request; prefer named styles.
""".strip()

_NAMESPACE = {"x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}

_CELLS = re.compile(
    r"/(?P<sheet>[^/]+)/(?P<first_column>[A-Z]{1,3})(?P<first_row>[1-9][0-9]*)"
    r"(?::(?P<last_column>[A-Z]{1,3})(?P<last_row>[1-9][0-9]*))?",
    re.IGNORECASE,
)
# Column names in order: A … Z, AA … ZZ, AAA … ZZZ.
_COLUMNS = [
    "".join(letters)
    for size in (1, 2, 3)
    for letters in product(ascii_uppercase, repeat=size)
]


async def prepare(
    file: Path,
) -> None:

    # Sheets, formulas and validation are workbook logic; nothing is removed.
    pass


async def inventory(
    file: Path,
) -> XlsxInventory:

    root, sheets, stylesheet = await _officecli.run(file, [
        {"command": "get", "path": "/", "depth": 0},
        {"command": "query", "selector": "sheet"},
        {"command": "raw", "part": "/styles"},
    ])

    stylesheet = ElementTree.fromstring(stylesheet["output"])
    style_formats = stylesheet.findall("x:cellStyleXfs/x:xf", _NAMESPACE)
    cell_formats = stylesheet.findall("x:cellXfs/x:xf", _NAMESPACE)

    # Partial ID matches can reuse direct alignment or protection overrides.
    known: dict[str, int] = {}

    for index, cell_format in enumerate(cell_formats):
        known.setdefault(ElementTree.canonicalize(
            ElementTree.tostring(cell_format, encoding="unicode"),
            strip_text=True, rewrite_prefixes=True,
        ), index)

    styles: dict[str, int] = {}
    additions: list[str] = []

    for style in stylesheet.iterfind("x:cellStyles/x:cellStyle", _NAMESPACE):

        style_id = style.get("xfId", "0")
        style_name = style.get("name")

        if (
            not style_name
            or not style_id.isdigit()
            or int(style_id) >= len(style_formats)
        ):
            continue

        # A named style applies through a cell format that refers to it.
        cell_format = deepcopy(style_formats[int(style_id)])
        cell_format.set("xfId", style_id)
        xml = ElementTree.tostring(cell_format, encoding="unicode")
        key = ElementTree.canonicalize(
            xml, strip_text=True, rewrite_prefixes=True,
        )

        if key not in known:

            known[key] = len(cell_formats) + len(additions)
            additions.append(xml)

        styles[style_name] = known[key]

    # Unlike the other inventories, this one writes: missing cell formats.
    if additions:
        await _officecli.run(file, [
            *(
                {
                    "command": "raw-set", "part": "/styles",
                    "xpath": "//x:cellXfs", "action": "append", "xml": xml,
                }
                for xml in additions
            ),
            {
                "command": "raw-set", "part": "/styles",
                "xpath": "//x:cellXfs", "action": "setattr",
                "xml": f"count={len(cell_formats) + len(additions)}",
            },
        ])

    return XlsxInventory(
        theme=_officecli.theme(root),
        sheets=[
            Sheet(name=sheet["path"].removeprefix("/"), rows=sheet["childCount"])
            for sheet in sheets["output"]["results"]
        ],
        styles=styles,
    )


def bind(
    inventory: XlsxInventory,
    master: int | None,
) -> XlsxInventory:

    if master is not None:
        raise ToolError("`master` applies to pptx designs only.")

    return inventory


def adapt(
    command: dict,
    inventory: XlsxInventory,
) -> dict:

    props = command.get("props", {})

    if "style" not in props:
        return command

    cells = _CELLS.fullmatch(command.get("path", ""))

    if command["command"] != "set" or len(props) != 1 or cells is None:
        raise ToolError(
            "`style` goes alone into a `set` on a cell or range, e.g. "
            '{"command": "set", "path": "/Sheet1/A1:C1", '
            '"props": {"style": "Header"}}.'
        )

    if (cell_format := inventory.styles.get(props["style"])) is None:
        raise ToolError(
            f"Style '{props['style']}' not found. Use a name from `styles`."
        )

    if cells["last_column"] is None:
        xpath = f"//x:c[@r='{cells['first_column'].upper()}{cells['first_row']}']"
    else:
        rows = sorted((int(cells["first_row"]), int(cells["last_row"])))
        columns = sorted((
            _COLUMNS.index(cells["first_column"].upper()),
            _COLUMNS.index(cells["last_column"].upper()),
        ))
        names = " ".join(_COLUMNS[columns[0]:columns[1] + 1])
        xpath = (
            f"//x:row[@r>={rows[0]} and @r<={rows[1]}]/x:c[contains("
            f"' {names} ', concat(' ', translate(@r, '0123456789', ''), ' '))]"
        )

    # Named styles are no engine prop; the cell's `s` attribute applies one.
    return {
        "command": "raw-set",
        "part": f"/{cells['sheet']}",
        "xpath": xpath,
        "action": "setattr",
        "xml": f"s={cell_format}",
    }
