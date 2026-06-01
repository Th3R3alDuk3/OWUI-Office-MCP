# OWUI-Office-MCP

> Office documents via MCP for OpenWebUI. Lean, modern, extensible.

---

## 🚀 Setup

```bash
uv sync
cp .env.example .env
```

In `.env`:
- `JWT_SECRET` → OpenWebUI's `WEBUI_SECRET_KEY`
- `OWUI_BASE_URL` → e.g. `http://localhost:3000` (reachable from the MCP server)

Place templates in [templates/](templates/).

## ▶️ Run

```bash
uv run python main.py
```

Runs as `streamable-http` on `HOST:PORT` from `.env`.

## 🛠️ Tools

Each subserver is mounted under its file extension as a namespace (`pptx_*`, `docx_*`, `xlsx_*`). Per-user state per subserver (JWT claim `id`), sliding TTL, auto-sweep — no disk writes.

A project starts one of two ways: from a template (`create_project`) or from an existing file in OpenWebUI (`open_project`). Master/layout/style/sheet discovery then runs against that project, so both branches behave the same.

### `pptx` (11)

| Tool | |
|---|---|
| `list_templates` | available `.pptx` templates |
| `create_project` | empty project from a template |
| `open_project` | open an existing `.pptx` from OpenWebUI by `file_id` |
| `list_masters` | slide masters of the current project |
| `list_layouts` | layouts + placeholders of a master |
| `insert_slide` | insert a slide from a layout (optionally at an index, otherwise append) |
| `list_slides` | list slides (index, layout, text) |
| `edit_slide` | update placeholders of a slide |
| `move_slide` | move a slide to a new position by index |
| `remove_slides` | remove slides by index |
| `download_project` | upload the project to OpenWebUI |

### `docx` (11)

| Tool | |
|---|---|
| `list_templates` | available `.docx` templates |
| `create_project` | empty project from a template |
| `open_project` | open an existing `.docx` from OpenWebUI by `file_id` |
| `list_styles` | paragraph and table styles of the current project |
| `insert_paragraph` | insert a paragraph (optionally at an index, otherwise append) |
| `insert_table` | insert a table (optionally with cell data, optionally at an index) |
| `insert_page_break` | insert a page break as a body block (optionally at an index) |
| `list_blocks` | list body blocks (index, type, text preview) |
| `move_block` | move a body block (paragraph & table) by index |
| `remove_blocks` | remove body blocks (paragraph & table) by index |
| `download_project` | upload the project to OpenWebUI |

### `xlsx` (12)

| Tool | |
|---|---|
| `list_templates` | available `.xlsx` templates |
| `create_project` | empty project from a template |
| `open_project` | open an existing `.xlsx` from OpenWebUI by `file_id` |
| `list_sheets` | worksheets of the current project (title, rows, cols) |
| `list_styles` | named cell style names of the current project |
| `add_sheet` | add a worksheet (optionally at an index, otherwise append) |
| `write_rows` | fill a contiguous block (tables / bulk data) from a 2D array at an anchor (optionally styled) |
| `write_cells` | write scattered individual cells by A1 reference (optionally styled) |
| `read_sheet` | read a sheet's used range as rows of text |
| `move_sheet` | move a worksheet to a new position by index |
| `remove_sheets` | remove worksheets by title |
| `download_project` | upload the project to OpenWebUI |
