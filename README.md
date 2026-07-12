# OWUI-Office-MCP

[![Docker](https://github.com/Th3R3alDuk3/OWUI-PPTX-MCP/actions/workflows/docker.yml/badge.svg)](https://github.com/Th3R3alDuk3/OWUI-PPTX-MCP/actions/workflows/docker.yml)
[![Version](https://img.shields.io/github/v/tag/Th3R3alDuk3/OWUI-PPTX-MCP?label=version)](https://github.com/Th3R3alDuk3/OWUI-PPTX-MCP/tags)
[![Python](https://img.shields.io/badge/python-3.12%2B-blue)](pyproject.toml)
[![License](https://img.shields.io/github/license/Th3R3alDuk3/OWUI-PPTX-MCP)](LICENSE)

> Office documents via MCP for OpenWebUI. Lean, modern, extensible.

PowerPoint, Word, and Excel toolsets for OpenWebUI over the Model Context
Protocol. The model starts a per-user project from a governed template or a
file attached in the chat, builds it with sandboxed Python scripts, and
uploads the result to OpenWebUI as a download link.

---

## ✨ Highlights

- **Code execution** — one Python script per edit batch instead of one
  tool round trip per change (see below)
- **Sandboxed, not `exec()`** — model code runs in the Monty sandbox,
  never a full Python interpreter
- **Stateful projects** — create from a template or open an attached
  `.pptx` / `.docx` / `.xlsx`, keep editing across the chat, finalize on
  demand
- **Template governance** — only the template's layouts and named styles
  are accepted, so corporate design is enforced technically, not by prompt
- **Charts and images** — native, editable charts in slides and workbooks,
  rendered chart images in documents, plus images the user attached in the
  chat
- **Review workflows** — annotate an opened file without changing its
  content: real Word comments (`docx`), cell notes (`xlsx`), speaker-note
  feedback (`pptx`)
- **Multi-user by design** — JWT auth against OpenWebUI's secret, one
  in-memory project per user and toolset, per-user rate limiting, TTL sweep
- **Agent-friendly** — every tool result carries a `hint` with the
  suggested next step, and `create_project` / `open_project` return the
  available layouts and styles up front

## ⚡ Code execution

Editing a document through classic MCP tools means one round trip per
change — a 20-slide deck becomes 20+ tool calls. OWUI-Office-MCP follows
the [code execution with MCP](https://www.anthropic.com/engineering/code-execution-with-mcp)
pattern instead: the model writes one plain Python script per edit batch
and sends it to `run_script`.

```python
revenue = {"Q1": 215, "Q2": 250, "Q3": 280}

i = add_slide("Title and Content")
fill(i, 0, "Revenue 2026")
fill(i, 1, "\n".join(f"{q}: {v} kEUR" for q, v in revenue.items()))
add_chart(i, "bar", list(revenue), {"Revenue": list(revenue.values())})
```

Loops, conditionals, and data transforms collapse a whole edit batch into
one call; `print(...)` output and the script's last expression come back
in the result, so the model can read and write in the same script.

Scripts run in [Monty](https://github.com/pydantic/monty), a sandboxed
interpreter with time and memory limits — no imports, no file or network
access. A script sees nothing but the toolset's facade functions, and the
facade only accepts the template's layouts and styles, so template
governance stays enforced by the runtime.

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

## 🏃 Run

```bash
uv run python main.py
```

The server listens on `0.0.0.0:8000`. Point OpenWebUI's MCP/tools config at
`http://<host>:8000/mcp`. Requests must carry a JWT signed by `JWT_SECRET`
with the OpenWebUI user's `id` claim.

## 🐳 Docker (optional)

Prebuilt images are published to **ghcr.io** on pushes to `main` (`latest`)
and on version tags (`X.Y.Z`):

```bash
docker run -d -p 8000:8000 \
--restart unless-stopped \
--env-file .env \
--name owui-office-mcp \
ghcr.io/th3r3alduk3/owui-pptx-mcp:latest
```

Or build the image locally: `docker build -t owui-office-mcp .`

Config is read from your `.env` via `--env-file`; to expose a different port,
change the mapping, e.g. `-p 9000:8000`. `TEMPLATES_DIR=./templates` resolves
to the templates baked into the image; mount `-v ./templates:/app/templates`
to swap them without rebuilding.

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

The script functions per toolset (zero-based indices everywhere):

| Toolset | Functions |
|---|---|
| `pptx` | `add_slide`, `fill` (tab-nested bullets), `set_notes`, `add_image` (attached OpenWebUI image), `add_chart`, `add_comment`, `list_slides`, `move_slide`, `remove_slides` |
| `docx` | `add_paragraph`, `add_table`, `add_page_break`, `add_image`, `add_chart`, `edit_paragraph`, `add_comment`, `list_blocks`, `move_block`, `remove_blocks` |
| `xlsx` | `write_rows`, `write_cell` (values, formulas, named styles), `add_image`, `add_chart`, `add_comment`, `read_sheet`, `list_sheets`, `add_sheet`, `move_sheet`, `remove_sheets` |

Format specifics:

- `add_chart` inserts a native, editable chart in `pptx` and `xlsx` — in
  `xlsx` linked to its data cells, so it updates when the sheet changes.
  `docx` gets a server-rendered PNG, as python-docx has no chart support.
- `add_comment` is a real Word comment in `docx`, a cell note in `xlsx`,
  and an appended speaker-notes paragraph in `pptx`.
- `xlsx_finalize_project` auto-fits every column's width to its content
  before uploading, so sheets open without cut-off text.

## ⚠️ Limits

- Projects live in memory only: a restart drops open projects, and idle
  projects expire after `PROJECT_TTL` seconds.
- Opening an attached `.xlsx` re-serializes the workbook through openpyxl:
  values, formulas, styles, standard charts, and images survive; VBA
  macros, slicers, form controls, and modern chart types (treemap, funnel)
  are dropped. `.docx` and `.pptx` are edited in place and keep everything.
- Requests are rate-limited per user (token bucket: `RATE_LIMIT_RPS`
  sustained, `RATE_LIMIT_BURST` burst), keyed on the JWT's `id` claim.
- The server speaks plain HTTP — put it behind a reverse proxy for TLS.
