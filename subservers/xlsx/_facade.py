from collections.abc import Callable
from io import BytesIO

from openpyxl.workbook import Workbook as WorkbookType

from subservers._chart import render_chart
from subservers.xlsx import _ops

SCRIPT_API = """
Run a Python script that builds or edits the current project. The script is
sandboxed: no imports, no file or network access. Available functions:

- write_rows(sheet: str, rows: list[list], anchor: str = "A1",
             style: str | None = None) -> None
  Fill a contiguous block of cells from row-major `rows`, starting at the
  `anchor` A1 reference. Numbers stay numeric, a string starting with `=`
  becomes a formula, and None clears the cell. An optional named cell style
  from `styles` (see the `create_project` / `open_project` result) formats
  every written cell.
- write_cell(sheet: str, ref: str, value=None,
             style: str | None = None) -> None
  Write one cell by A1 reference, optionally with a named style.
- add_chart(sheet: str, kind: str, categories: list[str],
            series: dict[str, list[float]], title: str | None = None,
            anchor: str = "A1") -> None
  Render a chart as an image with its top-left corner on the `anchor`
  cell. `kind` is "bar", "line" or "pie"; `series` maps each series name
  to one value per category. A pie chart takes exactly one series.
- read_sheet(sheet: str) -> list[list[str]]
  The sheet's used range as text rows (formulas as their formula string).
- add_sheet(title: str, index: int | None = None) -> None
- move_sheet(title: str, to_index: int) -> None
- remove_sheets(titles: list[str]) -> None
- list_sheets() -> list[dict]
  Worksheets in order (title, used rows/cols).

Plain Python works: loops, conditionals, f-strings, comprehensions,
functions, print. `print(...)` output and the last expression are returned
as `output`. If the script fails, fix it and call `run_script` again.

Example:

    data = [["Region", "Umsatz"], ["Nord", 1.2], ["Sued", 0.9]]
    write_rows("Sheet", data)
    write_cell("Sheet", f"B{len(data) + 1}", f"=SUM(B2:B{len(data)})")

Column widths are auto-fitted on `finalize_project`. After the requested
batch of edits, call `finalize_project` once (not after each change).
""".strip()


def script_functions(
    workbook: WorkbookType,
) -> dict[str, Callable]:

    def write_rows(
        sheet: str,
        rows: list[list[bool | int | float | str | None]],
        anchor: str = "A1",
        style: str | None = None,
    ) -> None:
        _ops.write_rows(workbook, sheet, rows, anchor, style)

    def write_cell(
        sheet: str,
        ref: str,
        value: bool | int | float | str | None = None,
        style: str | None = None,
    ) -> None:
        _ops.write_cell(workbook, sheet, ref, value, style)

    def add_chart(
        sheet: str,
        kind: str,
        categories: list[str],
        series: dict[str, list[float]],
        title: str | None = None,
        anchor: str = "A1",
    ) -> None:
        chart_png = render_chart(kind, categories, series, title)
        _ops.insert_picture(workbook, sheet, BytesIO(chart_png), anchor)

    def read_sheet(sheet: str) -> list[list[str]]:
        return _ops.read_sheet(workbook, sheet)

    def add_sheet(title: str, index: int | None = None) -> None:
        _ops.add_sheet(workbook, title, index)

    def move_sheet(title: str, to_index: int) -> None:
        _ops.move_sheet(workbook, title, to_index)

    def remove_sheets(titles: list[str]) -> None:
        _ops.remove_sheets(workbook, titles)

    def list_sheets() -> list[dict]:
        return [
            info.model_dump()
            for info in _ops.list_sheet_infos(workbook)
        ]

    return {
        "write_rows": write_rows,
        "write_cell": write_cell,
        "add_chart": add_chart,
        "read_sheet": read_sheet,
        "add_sheet": add_sheet,
        "move_sheet": move_sheet,
        "remove_sheets": remove_sheets,
        "list_sheets": list_sheets,
    }
