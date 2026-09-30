# OWUI-Office-MCP

[![App image](https://github.com/Th3R3alDuk3/OWUI-Office-MCP/actions/workflows/app.yml/badge.svg)](https://github.com/Th3R3alDuk3/OWUI-Office-MCP/actions/workflows/app.yml)
[![Version](https://img.shields.io/github/v/tag/Th3R3alDuk3/OWUI-Office-MCP?label=version)](https://github.com/Th3R3alDuk3/OWUI-Office-MCP/tags)
[![Python](https://img.shields.io/badge/python-3.14%2B-blue)](pyproject.toml)
[![License](https://img.shields.io/github/license/Th3R3alDuk3/OWUI-Office-MCP)](LICENSE)

> Office documents for OpenWebUI via MCP.

Creates and edits PowerPoint, Word and Excel files from a stored template or a
file the user attached, faithful to its design, and returns the result as an
OpenWebUI download link. The document engine is
[OfficeCLI](https://github.com/iOfficeAI/OfficeCLI); every run of it is
confined with Landlock through [landrun](https://github.com/Zouuup/landrun).
Drafts live in [Valkey](https://valkey.io) with a TTL.

## 🚀 Setup

Requires Docker with Compose on Linux. Full confinement of OfficeCLI needs
Linux 6.7 or newer (Landlock ABI 4); on older kernels the server starts with
a warning and OfficeCLI keeps network access. gVisor has no Landlock.

1. Configure:

   ```bash
   cp .env.example .env
   ```

   - `JWT_SECRET`: OpenWebUI's `WEBUI_SECRET_KEY`, so the server accepts the
     users' own tokens.
   - `OWUI_BASE_URL`: OpenWebUI as reached from this server;
     `OWUI_PUBLIC_URL` if users reach it under another URL.
   - `OWUI_VERIFY_TLS`: keep `true`; `false` only for self-signed or plain
     HTTP. Corporate CAs work via `SSL_CERT_FILE`.
   - `PREVIEW_TO_MODEL`: `true` hands previews to the model as well as to
     the user. Set `false` when a model without vision is in use: OpenWebUI
     would send it the image and the chat fails.
   - `VALKEY_URL`: already points at the Valkey of the compose file.

2. Fetch the pinned OfficeCLI and landrun once, then start the stack:

   ```bash
   bin/download.sh
   docker compose up -d --build
   ```

   This starts the server on port 8000, a Valkey bounded to 1 GB without
   persistence, and [Valkey Admin](https://github.com/valkey-io/valkey-admin)
   on `http://127.0.0.1:8080` to browse drafts (no login, so localhost only).

3. Connect OpenWebUI: in the admin settings add a tool server with
   **Add Connection**, type **MCP** (Streamable HTTP), URL
   `http://<host>:8000/mcp`, auth **Session**, so OpenWebUI forwards each
   user's own JWT. In the model's **Advanced Params** set **Function Calling**
   to **Native**. Start a new chat after the server's tools change.

Requests need a JWT signed with `JWT_SECRET` that names the user in the `id`
claim. The transport is stateless; projects are keyed on that claim, never on
an MCP session. Use a TLS reverse proxy outside a trusted network.

### Without Docker

For testing, the server runs directly with [uv](https://docs.astral.sh/uv/).
It expects `officecli` and `landrun` in `/usr/local/bin`, a Valkey at
`VALKEY_URL` and, for previews, a `chromium` on the `PATH`:

```bash
bin/download.sh && sudo install -m 755 bin/officecli bin/landrun /usr/local/bin/
docker run -d -p 6379:6379 valkey/valkey:9-alpine
uv sync
uv run python main.py
```

### Templates

[templates/](templates/) holds the stored designs (`.pptx`, `.docx`, `.xlsx`).
They are prepared once at startup: sample slides, sections and body text are
removed; masters, layouts and styles stay. A PPTX template may carry several
masters. Mount `-v ./templates:/app/templates` to swap templates without
rebuilding, and restart the server afterwards.

### Prebuilt image

CI publishes `ghcr.io/th3r3alduk3/owui-office-mcp` on pushes to `main`
(`latest`) and `dev` (`dev`) and on version tags (`X.Y.Z`). To run it instead
of building, replace `build:` in [compose.yaml](compose.yaml) with that
`image:`.

The image carries OfficeCLI and landrun, pinned by version and checksum in
[bin/download.sh](bin/download.sh), plus a headless Chromium and
Office-compatible fonts for previews. Building offline: the download script
is the only step that reaches GitHub; base image, apt and PyPI go through the
mirrors Docker, apt and uv are configured with, and the `[tool.uv]` block in
[pyproject.toml](pyproject.toml) selects a private index. Running needs no
network beyond OpenWebUI and Valkey.

## 🛠️ Tools

| Tool | Description |
|---|---|
| `list_templates` | Stored templates with their format and, for PPTX, their masters |
| `start_project` | Project from a stored template or an attached file: its design only, or the file as it is |
| `get_reference` | OfficeCLI reference for an element, a verb or both |
| `run_commands` | One atomic batch of OfficeCLI commands; reads return their output |
| `preview_project` | The document as thumbnails, or one slide or page, as an image |
| `export_project` | Upload the document to OpenWebUI as a download; again after later edits |

```json
{
  "project_id": "3f9a1c2e5b7d9f01",
  "commands": [
    {"command": "add", "parent": "/", "type": "slide", "props": {"layout": "3"}},
    {"command": "add", "parent": "/slide[1]", "type": "placeholder",
     "props": {"phType": "title", "text": "Revenue 2026"}},
    {"command": "add", "parent": "/slide[1]", "type": "chart",
     "props": {"chartType": "bar", "categories": "Q1,Q2,Q3", "data": "Revenue:215,250,280"}}
  ]
}
```

- `project_id`: from `start_project`, which also returns the design's
  inventory: the bound master's layouts and placeholder slots (`pptx`), style
  IDs split into custom and built-in (`docx`), sheets (`xlsx`).
- `keep_content`: edit an attached file as it is; without it only its design
  is used and its slides or body text are dropped.
- `master`: a `pptx` project is bound to one slide master. With several,
  `start_project` needs it and lists the masters when it is missing.
- `commands`: the OfficeCLI batch shape, any command except file-level ones
  (`create`, `open`, `merge`, ...). OfficeCLI validates and answers with its
  own errors; `get_reference` documents elements and props. Images come from
  attached files as `file:<file_id>`. A JSON string or a single object is
  accepted as well.

Every result carries a `hint` with the next step. A project expires after
`PROJECT_TTL` seconds without access; the exported file lives on in OpenWebUI.
Named Excel cell styles have no OfficeCLI prop yet
([iOfficeAI/OfficeCLI#455](https://github.com/iOfficeAI/OfficeCLI/issues/455));
until then, cells are formatted with direct props.

## ⚙️ Limits & security

[.env.example](.env.example) lists all settings: admission and rate limits per
user, OfficeCLI processes, memory and timeouts, and the project TTL.

- **Users:** every request carries the user's OpenWebUI JWT. Projects are keyed
  on the user and attachments are fetched with the user's own token, so one
  user never reaches another user's projects or files.
- **Isolation:** every OfficeCLI run starts through landrun and is confined by
  Landlock to its scratch directory: it reads `/usr`, `/proc`, `/dev` and the
  font configuration, nothing else, and has no network. Only a fixed
  environment reaches it. Verified in the image: a TCP connect fails with
  `EACCES`, and the server's environment under `/proc` is unreadable.
  `OFFICECLI_MAX_MEMORY` bounds a run's managed heap, so a pathological
  document kills its own run.
- **Commands:** file-level verbs are denied; everything else OfficeCLI
  validates itself. A PPTX project stays within its bound master: layouts are
  resolved there before the run, slides and placeholders are checked after
  it. Batches hold up to 500 commands and 1 MB; `run_commands` answers are
  cut at 50 KB.
- **Files:** uploads stop past 100 MB, images past 20 MB. The upload's content
  type selects the format, so macro and template variants are refused;
  OfficeCLI rejects decompression bombs and broken packages.
- **Drafts:** projects are one Valkey hash each with a TTL, edited under a
  per-project lock and written back only when the whole batch succeeded.
  Valkey runs bounded, evicts the least recently used drafts first and keeps
  nothing on disk; give it a password in `VALKEY_URL` wherever it is reachable
  beyond the compose network.
- **Errors:** tool errors say what to do next and carry no exception text;
  anything unexpected is masked.

### Sizing

One OfficeCLI run takes about a second; a PPTX edit takes two (the batch and
the master check), and a preview renders through Chromium. Admission
counters, the rate limit and the process semaphore are per server process.
The container runs non-root with a read-only root filesystem and a 4 GB
memory limit for Python, OfficeCLI runs and Chromium together; raise it or
lower `OFFICECLI_MAX_MEMORY` when `OFFICECLI_MAX_PROCESSES` grows. There is
no per-user quota; Valkey's memory limit bounds all drafts together.

Several replicas can share one Valkey; on Kubernetes that is a Deployment
without a volume. For the full confinement the nodes need Linux 6.7 or newer
and a container runtime that allows the Landlock syscalls (Docker 23+,
containerd 1.7+; gVisor does not).

## 🧩 Layout

- [`main.py`](main.py) - FastMCP server: auth, rate and response limits, lifespan, tools
- [`config.py`](config.py) - settings, read from `.env`
- [`models/office.py`](models/office.py) - the command shape, inventories and results
- [`services/owui.py`](services/owui.py) - OpenWebUI downloads and uploads
- [`services/officecli.py`](services/officecli.py) - OfficeCLI runs under landrun
- [`services/pptx.py`](services/pptx.py), [`docx.py`](services/docx.py), [`xlsx.py`](services/xlsx.py) - one module per format: preparation and inventory; pptx also the master binding
- [`services/project.py`](services/project.py) - Valkey project store, locks, templates, lifespan
- [`tools/`](tools/) - the six MCP tools and the command guard
- [`templates/`](templates/) - stored designs
- [`docker/`](docker/) - the server image; [`compose.yaml`](compose.yaml) runs it with Valkey
