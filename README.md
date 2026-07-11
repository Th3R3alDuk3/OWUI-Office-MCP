# OWUI-Office-MCP

> Office documents via MCP for OpenWebUI. Lean, modern, extensible.

PowerPoint, Word, and Excel toolsets for OpenWebUI over the Model Context
Protocol. The model starts a per-user project from a governed template or a
file attached in the chat, builds it with sandboxed Python scripts, and
uploads the result to OpenWebUI as a download link.

---

## ✨ Highlights

- **Code execution mode** — one Python script per edit batch instead of one
  tool round trip per change (see below)
- **Sandboxed, not `exec()`** — model code runs in the Monty sandbox,
  never a full Python interpreter
- **Stateful projects** — create from a template or open an attached
  `.pptx` / `.docx` / `.xlsx`, keep editing across the chat, finalize on
  demand; later edit requests continue on the same project
- **Template governance** — only the template's layouts and named styles
  are accepted, so corporate design is enforced technically, not by prompt
- **Charts and images** — native, editable charts in slides and workbooks,
  server-rendered chart images in documents, plus images the user attached
  in the chat
- **Review workflows** — annotate an opened file without changing its
  content: real Word comments (`docx`), cell notes (`xlsx`), speaker-note
  feedback (`pptx`)
- **Multi-user by design** — JWT auth against OpenWebUI's secret, one
  in-memory project per user and toolset, per-user rate limiting, TTL sweep
- **Agent-friendly** — every tool result carries a `hint` with the
  suggested next step, and `create_project` / `open_project` return the
  available layouts and styles up front, so no discovery round trips

## ⚡ Code execution mode

Editing a document through classic MCP tools means one round trip per
change — a 20-slide deck becomes 20+ tool calls, each with its own token
overhead. OWUI-Office-MCP instead follows the
[code execution with MCP](https://www.anthropic.com/engineering/code-execution-with-mcp)
pattern: the model writes one plain Python script per edit batch and sends
it to `run_script`.

```python
quarters = {"Q1": [120, 95], "Q2": [140, 110], "Q3": [160, 120]}

i = add_slide("Title and Content")
fill(i, 0, "Revenue 2026")
fill(i, 1, "\n".join(f"{q}: {sum(v)} kEUR" for q, v in quarters.items()))
add_chart(i, "bar", list(quarters),
          {"North": [v[0] for v in quarters.values()],
           "South": [v[1] for v in quarters.values()]})
```

Loops, conditionals, f-strings, data transforms — a whole edit batch
collapses into a single call. `print(...)` output and the script's last
expression come back in the result, so the model can read and write in the
same script (`read_sheet`, `list_slides`, `list_blocks`).

Model-written code never touches a full Python interpreter: scripts run in
[Monty](https://github.com/pydantic/monty), a sandboxed interpreter with
time and memory limits — no imports, no file or network access. A script
sees nothing but the toolset's facade functions, and the facade only
accepts the template's layouts and styles, so template governance stays
enforced by the runtime.

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

The script functions per toolset (zero-based indices everywhere):

| Toolset | Functions |
|---|---|
| `pptx` | `add_slide`, `fill` (tab-nested bullets), `set_notes`, `add_image` (attached OpenWebUI image), `add_chart`, `add_comment`, `list_slides`, `move_slide`, `remove_slides` |
| `docx` | `add_paragraph`, `add_table`, `add_page_break`, `add_image`, `add_chart`, `edit_paragraph`, `add_comment`, `list_blocks`, `move_block`, `remove_blocks` |
| `xlsx` | `write_rows`, `write_cell` (values, formulas, named styles), `add_image`, `add_chart`, `add_comment`, `read_sheet`, `list_sheets`, `add_sheet`, `move_sheet`, `remove_sheets` |

In `pptx` and `xlsx`, `add_chart` inserts a native, editable chart — in
`xlsx` linked to its data cells, so the chart updates when the sheet
changes. In `docx` it renders the chart server-side (matplotlib) and
inserts it as a PNG image, as python-docx has no chart support.
`xlsx_finalize_project` auto-fits every column's width to its content
before uploading, so sheets open without cut-off text.

`add_comment` covers review workflows on opened files without touching
their content: a real Word comment on a `docx` paragraph, a cell note in
`xlsx`, and an appended speaker-notes paragraph in `pptx` (PowerPoint
comments have no library support).

## 🔒 Limits

- Projects live in memory only: a restart drops open projects, and idle
  projects expire after `PROJECT_TTL` seconds.
- Opening an attached `.xlsx` re-serializes the workbook through openpyxl:
  values, formulas, styles, standard charts, and images survive, but
  anything outside openpyxl's model — VBA macros, slicers, form controls,
  modern chart types like treemap or funnel — is dropped in the finalized
  copy (`.docx` and `.pptx` are edited in place and keep everything
  untouched).
- Requests are rate-limited per user (token bucket: `RATE_LIMIT_RPS`
  sustained, `RATE_LIMIT_BURST` burst), keyed on the JWT's `id` claim.
- The server speaks plain HTTP — put it behind a reverse proxy for TLS.
