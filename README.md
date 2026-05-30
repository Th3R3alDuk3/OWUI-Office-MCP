# OWUI-Office-MCP

> Office documents via MCP for OpenWebUI. Lean, modern, extensible.

⚡ FastMCP &nbsp;·&nbsp; 📦 uv &nbsp;·&nbsp; 🔐 JWT (HS256) &nbsp;·&nbsp; 🌐 streamable-http

---

## 🎯 Status

| Name | Format | Status |
|---|---|---|
| PowerPoint | `.pptx` | ✅ |
| Word | `.docx` | ✅ |

## 🚀 Setup

```bash
uv sync
cp .env.example .env
```

In `.env`:
- `JWT_SECRET` → OpenWebUI's `WEBUI_SECRET_KEY`
- `OWUI_BASE_URL` → e.g. `http://localhost:3000` (reachable from the MCP server)

Place templates in 📁 [templates/](templates/).

## ▶️ Run

```bash
uv run python main.py
```

Runs as `streamable-http` on `HOST:PORT` from `.env`.

## 🛠️ Tools

Each subserver is mounted under its file extension as a namespace (`pptx_*`, `docx_*`). Per-user state per subserver (JWT claim `id`), sliding TTL, auto-sweep — no disk writes.

### `pptx` (10)

| Tool | |
|---|---|
| `list_templates` | available `.pptx` templates |
| `list_masters` | slide masters of a template |
| `list_layouts` | layouts + placeholders of a master |
| `create_project` | empty project from a template |
| `insert_slide` | insert a slide from a layout (optionally at an index, otherwise append) |
| `list_slides` | list slides (index, layout, text) |
| `edit_slide` | update placeholders of a slide |
| `move_slide` | move a slide to a new position by index |
| `remove_slides` | remove slides by index |
| `download_project` | upload the project to OpenWebUI |

### `docx` (10)

| Tool | |
|---|---|
| `list_templates` | available `.docx` templates |
| `list_styles` | paragraph and table styles of a template |
| `create_project` | empty project from a template |
| `insert_paragraph` | insert a paragraph (optionally at an index, otherwise append) |
| `insert_table` | insert a table (optionally with cell data, optionally at an index) |
| `insert_page_break` | insert a page break as a body block (optionally at an index) |
| `list_blocks` | list body blocks (index, type, text preview) |
| `move_block` | move a body block (paragraph & table) by index |
| `remove_blocks` | remove body blocks (paragraph & table) by index |
| `download_project` | upload the project to OpenWebUI |
