# OWUI-Office-MCP

[![Docker](https://github.com/Th3R3alDuk3/OWUI-Office-MCP/actions/workflows/docker.yml/badge.svg)](https://github.com/Th3R3alDuk3/OWUI-Office-MCP/actions/workflows/docker.yml)
[![Version](https://img.shields.io/github/v/tag/Th3R3alDuk3/OWUI-Office-MCP?label=version)](https://github.com/Th3R3alDuk3/OWUI-Office-MCP/tags)
[![Python](https://img.shields.io/badge/python-3.14%2B-blue)](pyproject.toml)
[![License](https://img.shields.io/github/license/Th3R3alDuk3/OWUI-Office-MCP)](LICENSE)

> Office documents via MCP for OpenWebUI. Lean, modern, template-faithful.

PowerPoint, Word and Excel for OpenWebUI over the Model Context Protocol.
The model starts a project from a stored template or a file the user
attached, builds it with batches of [OfficeCLI](https://github.com/iOfficeAI/OfficeCLI)
commands, and uploads the result to OpenWebUI as a download link.

---

## ✨ Highlights

- **One engine, three formats** — OfficeCLI edits `.pptx`, `.docx` and
  `.xlsx` in place: native charts and real comments everywhere, and the
  parts it does not touch stay as they are
- **Corporate design from templates** — stored on the server or attached by
  the user as their own slide master; several masters per file are fine.
  Template mode only accepts the design's layouts, placeholders and styles;
  `design_mode="custom"` unlocks colors and fonts on explicit request
- **Persistent projects** — a user holds several projects on disk; they
  survive restarts, expire after `PROJECT_TTL` and stay within quotas
- **Atomic batches** — one `run_commands` call per edit step; if one
  command fails, nothing is applied
- **Multi-user by design** — JWT auth against OpenWebUI's secret, per-user
  rate limiting, bounded admission and OfficeCLI concurrency, stateless
  transport
- **Agent-friendly** — every result carries a `hint`, start results carry
  the design's inventory, `get_reference` documents elements and props on
  demand, and `preview_project` lets a vision model check its own work

## 🚀 Quick start

Requires [uv](https://docs.astral.sh/uv/) and the
[OfficeCLI](https://github.com/iOfficeAI/OfficeCLI/releases) binary
(`officecli-linux-x64`, version pinned in the [Dockerfile](Dockerfile)) as
`/usr/local/bin/officecli`. The Docker image ships it. OfficeCLI runs
confined by Landlock, so the host needs Linux 6.7 or newer and, for Docker,
version 23 or newer; the server refuses to start otherwise. Previews need
Chromium under `/usr`; the image ships its headless shell.

```bash
uv sync
cp .env.example .env
uv run python main.py
```

Set at least these in `.env`:

- `JWT_SECRET` → OpenWebUI's `WEBUI_SECRET_KEY`. Both sides need the same value.
- `OWUI_BASE_URL` → OpenWebUI URL reachable from this server
- `OWUI_VERIFY_TLS` → `false` only for self-signed or plain-HTTP setups;
  corporate CAs work via `SSL_CERT_FILE` / `SSL_CERT_DIR`

Templates live in [templates/](templates/) and are prepared at startup by
extension (`.pptx`, `.docx`, `.xlsx`); restart after changing them. Projects
and prepared templates live in `data/`.

The server listens on `0.0.0.0:8000`. Requests must carry a JWT signed with
`JWT_SECRET` that names the user in the `id` claim. The transport is
stateless — projects are keyed on that claim, never on an MCP session. The
server speaks plain HTTP, so put a TLS reverse proxy in front of it outside a
trusted network.

### Connect OpenWebUI

1. In the admin settings, add a tool server with **Add Connection**: type
   **MCP** (Streamable HTTP), URL `http://<host>:8000/mcp`, auth
   **Session**, so OpenWebUI forwards each user's own JWT.
2. Put these **Headers** on the same connection:

   ```json
   {"X-OpenWebUI-Chat-Id": "{{CHAT_ID}}", "X-OpenWebUI-Message-Id": "{{MESSAGE_ID}}"}
   ```

   OpenWebUI fills in the current chat and message, and `export_project`
   attaches the finished file to the reply as a chip that opens OpenWebUI's
   own PPTX, DOCX or XLSX preview. Without the headers the reply only carries
   the download link.
3. In the model's **Advanced Params**, set **Function Calling** to
   **Native** and use a model that reads images. Only then do the slide
   images of `preview_project` reach the model; the user sees them either
   way.
4. Start a new chat after the server's tools change.

## 🛠️ Tools

| Tool | Description |
|---|---|
| `list_templates` | Stored templates with their format and, for `pptx`, their masters |
| `create_project` | New, empty project from a stored template or from an attached file's design |
| `open_project` | Edit a file attached in OpenWebUI as it is |
| `run_commands` | Apply one atomic batch of OfficeCLI commands; reads return their output |
| `get_reference` | OfficeCLI reference for an element, a verb or both, limited to the accepted props |
| `preview_project` | One slide or page as an image for the model to check visually |
| `export_project` | Check the document and upload it to OpenWebUI; again after later edits |

A `pptx` project is bound to one slide master: `list_templates` shows the
masters of stored templates; with several masters, `create_project` and
`open_project` need `master` and list the masters when it is missing. Both
return the design's inventory: the bound master's layouts and placeholder
slots (`pptx`), style IDs split into custom and built-in (`docx`), sheets and
named cell styles (`xlsx`), plus the theme palette and fonts. `run_commands`
takes the OfficeCLI batch shape:

```jsonc
[
  {"command": "add", "parent": "/", "type": "slide", "props": {"layout": "3"}},
  {"command": "add", "parent": "/slide[1]", "type": "placeholder",
   "props": {"phType": "title", "text": "Revenue 2026"}},
  {"command": "add", "parent": "/slide[1]", "type": "chart",
   "props": {"chartType": "bar", "categories": "Q1,Q2,Q3", "data": "Revenue:215,250,280"}}
]
```

Only a supported subset of OfficeCLI is exposed: `add`, `set`, `remove`,
`move`, `swap`, `import`, `get`, `query`, `view`, `validate` with a
per-format list of element types and props. Images come from attached files
as `file:<file_id>`. Masters, layouts and style definitions are read-only.
Excel named styles are applied through a `style` prop that the server
resolves to the workbook's cell format.

## ⚙️ Configuration

`.env.example` lists every setting.

| Variable | Purpose |
|---|---|
| `JWT_SECRET` | OpenWebUI's `WEBUI_SECRET_KEY`; required, not empty |
| `JWT_ALGORITHM` | Signing algorithm of the JWT; `HS256` for OpenWebUI |
| `OWUI_BASE_URL` | OpenWebUI URL, reachable from this server; required |
| `OWUI_VERIFY_TLS` | Verify OpenWebUI's TLS certificate |
| `MAX_CONCURRENT_REQUESTS` | Server-wide cap on document requests |
| `MAX_CONCURRENT_REQUESTS_PER_USER` | Per-user cap on document requests |
| `RATE_LIMIT_RPS` / `RATE_LIMIT_BURST` | Per-user request rate |
| `OFFICECLI_MAX_PROCESSES` | OfficeCLI processes running at once |
| `OFFICECLI_TIMEOUT` | Seconds one OfficeCLI run may take |
| `OFFICECLI_QUEUE_TIMEOUT` | Seconds a request waits for its project and a process |
| `DATA_MAX_SIZE` | MB for all data: projects, assets and prepared templates |
| `DATA_MAX_SIZE_PER_USER` | MB per user |
| `PROJECT_MAX_PER_USER` | Projects a user may hold |
| `PROJECT_TTL` | Seconds of inactivity until a project expires |
| `PROJECT_SWEEP_INTERVAL` | Seconds between expiry sweeps |

## 🔒 Security model

**Users.** Every request carries the user's OpenWebUI JWT. Projects live under
the user's key, and attachments are fetched with the user's own token, so one
user never reaches another user's projects or files.

**Isolation.** Every OfficeCLI run is confined by Landlock
([py-landlock](https://github.com/SebastienWae/py-landlock)) to its project
directory and a temporary directory, without network. Other users' projects,
the server's code and its environment are out of reach, whatever the command.

**Commands.** The model only gets a subset of OfficeCLI. Template mode accepts
the design's layouts, placeholders and styles; colors and fonts need
`design_mode="custom"`. Masters, layouts and style definitions are read-only.
Images come from attachments as `file:<file_id>`.

**Files.** Uploads stop past 100 MB, images past 20 MB. The package manifest
must name exactly one `.pptx`, `.docx` or `.xlsx` document, so macro and
template formats are refused; OfficeCLI itself rejects decompression bombs
and broken packages. Batches hold up to 500 commands.

**Errors.** Tool errors say what to do next and carry no exception text;
anything unexpected is masked.

**Known limits.** One server process owns `data/`; replicas need sticky
routing and separate data directories. Quotas are checked before each write,
so one write may overshoot by its own size. Excess requests get a retryable
busy error, and each OfficeCLI call costs about a second of startup.

## 📦 Docker deployment

Prebuilt images are published to **ghcr.io** on pushes to `main` (`latest`)
and on version tags (`X.Y.Z`); OfficeCLI is baked in, pinned by version and
checksum, together with a headless Chromium and Office-compatible fonts for
previews.

```bash
docker run -d -p 8000:8000 \
--restart unless-stopped \
--env-file .env \
--read-only --tmpfs /tmp \
-v ./data:/app/data \
--name owui-office-mcp \
ghcr.io/th3r3alduk3/owui-office-mcp:latest
```

Or build the image locally: `docker build -t owui-office-mcp .`

Mount `./data` to keep projects across restarts. Mount
`-v ./templates:/app/templates` to swap templates without rebuilding; restart
the container to prepare them. The container runs as a non-root user; only
`/app/data` and `/tmp` need to be writable.
