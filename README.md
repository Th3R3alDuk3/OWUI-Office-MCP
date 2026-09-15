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
- **Drafts with a TTL** — projects live in [Valkey](https://valkey.io),
  expire after `PROJECT_TTL` of inactivity and are bounded by its memory
  limit; nothing is written to disk. Finished files live in OpenWebUI
- **Atomic batches** — one `run_commands` call per edit step; if one
  command fails, nothing is applied
- **Confined engine** — every OfficeCLI run is sandboxed with Landlock
  through [landrun](https://github.com/Zouuup/landrun): no network, nothing
  beyond its own scratch directory
- **Multi-user by design** — JWT auth against OpenWebUI's secret, per-user
  rate limiting, bounded admission and OfficeCLI concurrency, stateless
  transport; any number of replicas can share one Valkey
- **Agent-friendly** — every result carries a `hint`, start results carry
  the design's inventory, `get_reference` documents elements and props on
  demand, and `preview_project` lets a vision model check its own work

## 🚀 Quick start

Requires Docker with Compose on Linux 6.7 or newer: OfficeCLI runs confined
by Landlock, and the server refuses to start where the kernel cannot
restrict the network (Landlock ABI 4). gVisor does not implement Landlock.

```bash
cp .env.example .env
docker compose up -d
```

Set at least these in `.env`:

- `JWT_SECRET` → OpenWebUI's `WEBUI_SECRET_KEY`. Both sides need the same value.
- `OWUI_BASE_URL` → OpenWebUI URL reachable from this server
- `OWUI_VERIFY_TLS` → `false` only for self-signed or plain-HTTP setups;
  corporate CAs work via `SSL_CERT_FILE` / `SSL_CERT_DIR`

`VALKEY_URL` already points at the Valkey of the compose file. Templates
live in [templates/](templates/) and are prepared at startup by extension
(`.pptx`, `.docx`, `.xlsx`); restart the server after changing them.

The server listens on `0.0.0.0:8000`. Requests must carry a JWT signed with
`JWT_SECRET` that names the user in the `id` claim. The transport is
stateless — projects are keyed on that claim, never on an MCP session. The
server speaks plain HTTP, so put a TLS reverse proxy in front of it outside a
trusted network.

### Connect OpenWebUI

1. In the admin settings, add a tool server with **Add Connection**: type
   **MCP** (Streamable HTTP), URL `http://<host>:8000/mcp`, auth
   **Session**, so OpenWebUI forwards each user's own JWT.
2. In the model's **Advanced Params**, set **Function Calling** to
   **Native** and use a model that reads images. Only then do the slide
   images of `preview_project` reach the model; the user sees them either
   way.
3. Start a new chat after the server's tools change.

## 🛠️ Tools

| Tool | Description |
|---|---|
| `list_templates` | Stored templates with their format and, for `pptx`, their masters |
| `create_project` | New, empty project from a stored template or from an attached file's design |
| `open_project` | Edit a file attached in OpenWebUI as it is |
| `get_reference` | OfficeCLI reference for an element, a verb or both, limited to the accepted props |
| `run_commands` | Apply one atomic batch of OfficeCLI commands; reads return their output |
| `preview_project` | The whole document as thumbnails, or one slide or page, as an image for the model to check visually |
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
per-format list of element types and props. Elements: slides, placeholders,
paragraphs, notes, comments, charts, pictures, tables and text boxes
(`pptx`); paragraphs, tables, page breaks, charts, pictures, comments,
hyperlinks and a table of contents (`docx`); sheets, charts, pictures,
comments, Excel tables, autofilters and conditional formatting (`xlsx`).
Images come from attached files as `file:<file_id>`. Masters, layouts and
style definitions are read-only. Excel named styles are applied through a
`style` prop that the server resolves to the workbook's cell format.

## ⚙️ Configuration

`.env.example` lists every setting.

| Variable | Purpose |
|---|---|
| `JWT_SECRET` | OpenWebUI's `WEBUI_SECRET_KEY`; required, not empty |
| `JWT_ALGORITHM` | Signing algorithm of the JWT; `HS256` for OpenWebUI |
| `OWUI_BASE_URL` | OpenWebUI URL, reachable from this server; required |
| `OWUI_PUBLIC_URL` | OpenWebUI URL for download links, if it differs from `OWUI_BASE_URL` |
| `OWUI_VERIFY_TLS` | Verify OpenWebUI's TLS certificate |
| `VALKEY_URL` | `valkey://[:password@]host:6379[/db]` of the project store |
| `MAX_CONCURRENT_REQUESTS` | Cap on document requests per server process |
| `MAX_CONCURRENT_REQUESTS_PER_USER` | Per-user cap on document requests |
| `RATE_LIMIT_RPS` / `RATE_LIMIT_BURST` | Per-user request rate |
| `OFFICECLI_MAX_PROCESSES` | OfficeCLI processes running at once |
| `OFFICECLI_MAX_MEMORY` | MB of managed heap one OfficeCLI run may use |
| `OFFICECLI_TIMEOUT` | Seconds one OfficeCLI run may take |
| `OFFICECLI_QUEUE_TIMEOUT` | Seconds a request waits for its project and a process |
| `PROJECT_MAX_PER_USER` | Projects a user may hold |
| `PROJECT_TTL` | Seconds of inactivity until a project expires |

## 🔒 Security model

**Users.** Every request carries the user's OpenWebUI JWT. Projects are keyed
on the user, and attachments are fetched with the user's own token, so one
user never reaches another user's projects or files.

**Isolation.** Every OfficeCLI run starts through landrun and is confined by
Landlock to its scratch directory and a temporary directory: it can read
`/usr`, `/proc`, `/dev` and the font configuration, nothing else, and it has
no network. Only a fixed environment reaches it. Verified in the image: a
TCP connect fails with `EACCES`, and the server's own environment under
`/proc` is unreadable although the child shares the server's UID. A run's
managed heap is bounded by `OFFICECLI_MAX_MEMORY`, so a pathological
document kills its own run.

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

**Known limits.** Admission counters and rate limits are per server process;
with several replicas they are approximate, and the per-user project limit
can be exceeded by one when a user creates projects concurrently. There is
no per-user byte quota, so one user's projects can evict other users' drafts
within Valkey's memory limit. Drafts are gone when Valkey restarts without
persistence or evicts them under memory pressure; exported files are safe in
OpenWebUI. The server's Python and Chromium memory is bounded only by the
container limit in `compose.yaml`, each OfficeCLI run's managed heap by
`OFFICECLI_MAX_MEMORY`. Each OfficeCLI call costs about a second of startup.

## 📦 Docker deployment

[compose.yaml](compose.yaml) runs the server, a Valkey and
[Valkey Admin](https://github.com/valkey-io/valkey-admin) on
`http://127.0.0.1:8080` to browse keys, TTLs and memory (no login, so it is
bound to localhost). Its connection list is kept in the browser: add one
with **Add Connection**: endpoint type Node, host `valkey`, port `6379`,
TLS off, password empty. Prebuilt images
are published to **ghcr.io** on pushes to `main` (`latest`) and `dev` (`dev`)
and on version tags (`X.Y.Z`); OfficeCLI and landrun are baked in, pinned by
version and checksum in `bin/download.sh`, together with a headless
Chromium and Office-compatible fonts for previews.

```bash
docker compose up -d           # pulls the image
bin/download.sh           # fetches the pinned binaries once
docker compose up -d --build   # builds it from this checkout
```

The script is the only step that reaches GitHub. The base image, apt and
PyPI go through whatever mirrors Docker, apt and uv are configured with; an
offline build needs those mirrors, or a `docker save` of an image built
elsewhere. Running needs no network beyond OpenWebUI and Valkey.

Valkey runs with a memory limit, evicts the least recently used drafts
first and keeps nothing on disk; change that in `compose.yaml` (a larger
`--maxmemory`, or `--appendonly yes` with a volume to survive restarts). It
is reachable only inside the compose network; put a password in
`VALKEY_URL` wherever that is not the case.

Mount `-v ./templates:/app/templates` to swap templates without rebuilding;
restart the server to prepare them. The container runs as a non-root user
with a read-only root filesystem, only `/tmp` writable, and a memory limit of
4 GB (`mem_limit`) for Python, OfficeCLI runs and Chromium together; scale it
with `OFFICECLI_MAX_PROCESSES` and `MAX_CONCURRENT_REQUESTS`.

Several server replicas can share one Valkey. On Kubernetes that is a
Deployment without a volume; the nodes need Linux 6.7 or newer, and the
container runtime must allow the Landlock syscalls (Docker 23+ and
containerd 1.7+ do by default; gVisor does not implement them).
