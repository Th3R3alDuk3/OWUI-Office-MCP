# OWUI-Office-MCP

> Office documents via MCP for OpenWebUI. Lean, modern, extensible.

---

## 🚀 Setup

```bash
uv sync
cp .env.example .env
```

Set the important values in `.env`:

- `JWT_SECRET` → OpenWebUI's `WEBUI_SECRET_KEY`
- `OWUI_BASE_URL` → OpenWebUI URL reachable from this server, e.g. `http://localhost:3000`
- `OWUI_VERIFY_TLS` → set `false` only for self-signed or plain-HTTP lab setups

Place templates in [templates/](templates/) — `list_templates` picks them up
by extension (`.pptx`, `.docx`, `.xlsx`).

## ▶️ Run

```bash
uv run python main.py
```

The server listens on `0.0.0.0:8000`. Point OpenWebUI's MCP/tools config at
`http://<host>:8000/mcp`. Requests must carry a JWT signed by `JWT_SECRET`
with the OpenWebUI user's `id` claim.

## 🐳 Docker (optional)

```bash
docker build -t owui-office-mcp .

docker run -d -p 8000:8000 \
--restart unless-stopped \
--env-file .env \
--name owui-office-mcp
owui-office-mcp
```

Config is read from your `.env` via `--env-file`; to expose a different port,
change the mapping, e.g. `-p 9000:8000`. `TEMPLATES_DIR=./templates` resolves
to the templates baked into the image; mount `-v ./templates:/app/templates`
to swap them without rebuilding. Prebuilt images are published to **ghcr.io**
on pushes to `main` and on version tags.

## 🛠️ Tools

Each subserver is mounted under its file extension as a namespace (`pptx_*`,
`docx_*`, `xlsx_*`). Per user and toolset there is one in-memory project
(keyed on the JWT's `id` claim, sliding TTL, auto-sweep — no disk writes).

All three toolsets share the same project lifecycle:

| Tool | |
|---|---|
| `list_templates` | available templates for the toolset |
| `create_project` | new, empty project from a template |
| `open_project` | open a file attached in OpenWebUI by `file_id` |
| `finalize_project` | serialize the project and upload it to OpenWebUI |

Discovery (`list_masters`, `list_styles`, `list_sheets`, …) runs against the
current project, so both starting points behave the same. All `insert_*` and
`move_*` tools use zero-based indices; inserts append when no index is given.
Every tool result carries a `hint` with the suggested next step, guiding
agents through the flow.

### `pptx`

| Tool | |
|---|---|
| `list_masters` | slide masters of the current project |
| `list_layouts` | layouts + placeholders of a master |
| `insert_slide` | insert a slide from a layout, optionally filling placeholders |
| `list_slides` | slides in order (layout, text) |
| `edit_slide` | update placeholders of a slide |
| `move_slide` | move a slide by index |
| `remove_slides` | remove slides by index |

### `docx`

| Tool | |
|---|---|
| `list_styles` | paragraph and table styles |
| `insert_paragraph` | insert a styled paragraph |
| `insert_table` | insert a table, optionally with cell data |
| `insert_page_break` | insert a page break |
| `list_blocks` | body blocks in order (type, text preview) |
| `move_block` | move a body block by index |
| `remove_blocks` | remove body blocks by index |

### `xlsx`

| Tool | |
|---|---|
| `list_sheets` | worksheets (title, used extent) |
| `list_styles` | named cell styles |
| `insert_sheet` | add a worksheet |
| `write_rows` | fill a contiguous block from a 2D array at an anchor |
| `write_cells` | write individual cells by A1 reference |
| `read_sheet` | read a sheet's used range as text rows |
| `move_sheet` | move a worksheet by index |
| `remove_sheets` | remove worksheets by title |

## 🔒 Limits

- Projects live in memory only: a restart drops open projects, and idle
  projects expire after `PROJECT_TTL` seconds.
- Requests are rate-limited per user (token bucket: `RATE_LIMIT_RPS`
  sustained, `RATE_LIMIT_BURST` burst), keyed on the JWT's `id` claim.
- The server speaks plain HTTP — put it behind a reverse proxy for TLS.
