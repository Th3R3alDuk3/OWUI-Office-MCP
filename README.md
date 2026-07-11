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
--name owui-office-mcp \
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

All three toolsets expose the same five tools:

| Tool | |
|---|---|
| `list_templates` | available templates for the toolset |
| `create_project` | new, empty project from a template |
| `open_project` | open a file attached in OpenWebUI by `file_id` |
| `run_script` | build/edit the project with one sandboxed Python script |
| `finalize_project` | serialize the project and upload it to OpenWebUI |

Building and editing happens in `run_script`: the model writes plain Python
(loops, f-strings, data transforms) against a small facade of functions
documented in the tool description — one call per edit batch instead of
many round trips. Scripts run in
[Monty](https://github.com/pydantic/monty), a sandboxed interpreter with
time/memory limits: no imports, no file or network access, and the facade
only accepts the template's layouts and styles, so template governance
stays technically enforced. `create_project` / `open_project` return the
available layouts/styles directly, so no discovery round trips are needed.
Every tool result carries a `hint` with the suggested next step, guiding
agents through the flow.

The script functions per toolset (zero-based indices everywhere):

| Toolset | Functions |
|---|---|
| `pptx` | `add_slide`, `fill` (tab-nested bullets), `set_notes`, `add_image` (attached OpenWebUI image), `add_chart`, `list_slides`, `move_slide`, `remove_slides` |
| `docx` | `add_paragraph`, `add_table`, `add_page_break`, `add_image`, `add_chart`, `edit_paragraph`, `list_blocks`, `move_block`, `remove_blocks` |
| `xlsx` | `write_rows`, `write_cell` (values, formulas, named styles), `add_chart`, `read_sheet`, `add_sheet`, `move_sheet`, `remove_sheets`, `list_sheets` |

`add_chart` renders a bar, line, or pie chart server-side (matplotlib) and
inserts it as a PNG image. `xlsx_finalize_project` auto-fits every column's
width to its content before uploading, so sheets open without cut-off text.

## 🔒 Limits

- `run_script` code runs in the Monty sandbox: no imports, no file or
  network access, capped execution time and memory.
- Projects live in memory only: a restart drops open projects, and idle
  projects expire after `PROJECT_TTL` seconds.
- Requests are rate-limited per user (token bucket: `RATE_LIMIT_RPS`
  sustained, `RATE_LIMIT_BURST` burst), keyed on the JWT's `id` claim.
- The server speaks plain HTTP — put it behind a reverse proxy for TLS.
