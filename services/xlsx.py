from pathlib import Path

from models.office import Sheet, XlsxInventory
from services.officecli import run_batch

FORMAT = "xlsx"
MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
INVENTORY = XlsxInventory


START_HINT = """
Sheets, values and formulas of the design are kept; `inventory.sheets`
counts their rows. Build in steps and read back with `view`. `import` with
`text` (CSV, a command field) and the props `startCell` and `header` fills a
block; `set` with the props `value` or `formula` one cell. Charts take
`chartType`, `dataRange` "A1:B3" and `anchor` "D2:J18"; images `src`
"file:<file_id>" from attached files. Stay within the design: change colors,
fonts and other appearance only when the user asks. `get_reference`
documents elements and props.
""".strip()


async def prepare(
    file: Path,
) -> None:

    # Sheets, formulas and validation are workbook logic; nothing is removed.
    pass


async def inventory(
    file: Path,
) -> XlsxInventory:

    (sheets,) = await run_batch(file, [
        {"command": "query", "selector": "sheet"},
    ])

    return XlsxInventory(sheets=[
        Sheet(name=sheet["path"].removeprefix("/"), rows=sheet["childCount"])
        for sheet in sheets["output"]["results"]
    ])
