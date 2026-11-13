# OWUI-Office-MCP

[![App image](https://github.com/Th3R3alDuk3/OWUI-Office-MCP/actions/workflows/app.yml/badge.svg)](https://github.com/Th3R3alDuk3/OWUI-Office-MCP/actions/workflows/app.yml)
[![Python](https://img.shields.io/badge/python-3.14%2B-blue)](pyproject.toml)
[![OpenWebUI](https://img.shields.io/badge/OpenWebUI-%E2%89%A5%200.12.5-blue)](https://github.com/open-webui/open-webui)
[![License](https://img.shields.io/github/license/Th3R3alDuk3/OWUI-Office-MCP)](LICENSE)

> Office documents for OpenWebUI via MCP.

Creates and edits PowerPoint, Word and Excel files from a stored template or
an attached file, faithful to its design, and returns the result as an
OpenWebUI file with ID and download path. Built on
[OfficeCLI](https://github.com/iOfficeAI/OfficeCLI), every run confined with
[landrun](https://github.com/Zouuup/landrun); drafts live as the user's files
in OpenWebUI.

## 🚀 Setup

Requires Docker with Compose on Linux; full confinement of OfficeCLI needs
Linux 6.7 or newer.

1. Configure:

   ```bash
   cp .env.example .env
   ```

   - `JWT_ISSUER`: the Keycloak realm, `https://<keycloak>/realms/<realm>`,
     exactly as its tokens name it in `iss`.
   - `JWT_JWKS_URI`: the realm's signing keys,
     `<JWT_ISSUER>/protocol/openid-connect/certs`, reachable from this server.
   - `OWUI_BASE_URL`: OpenWebUI URL reachable from this server.
   - Corporate or self-signed CAs for Keycloak and OpenWebUI: mount a PEM
     bundle with every CA to trust and set `SSL_CERT_FILE` to it.

2. Fetch the pinned OfficeCLI and landrun, then start the stack:

   ```bash
   bin/download.sh
   docker compose up -d --build
   ```

   The server listens on port 8000.

3. Connect OpenWebUI 0.12.5 or newer, signing users in through the same
   Keycloak realm: add an **MCP** tool server at
   `http://<host>:8000/mcp` with auth **OAuth**, assign it to models and
   set their **Function Calling** to **Native**. Previews reach the model as
   images; turn **Vision** off for text-only models, which then get an
   `[image]` note instead. If OpenWebUI sets
   `AIOHTTP_CLIENT_TIMEOUT_TOOL_SERVER` or `AIOHTTP_CLIENT_TIMEOUT`, keep it
   at 600 s or more, or a long edit is reported as failed although it
   applied.

Requests need an access token of that realm signed with RS256; its `sub`
claim names the user. For OpenWebUI's files the server exchanges it for the
user's OpenWebUI token, so OpenWebUI needs
`ENABLE_OAUTH_TOKEN_EXCHANGE=true` and `OAUTH_TOKEN_EXCHANGE_TRUSTED_CLIENT_IDS`
set to its client ID; Keycloak then has to add that client to the token's
audience (audience mapper). Use a TLS reverse proxy outside a trusted network.

### Without Docker

For testing, the server runs directly with [uv](https://docs.astral.sh/uv/).
It expects `officecli` and `landrun` in `/usr/local/bin` and, for previews, a
`chromium` on the `PATH`:

```bash
bin/download.sh && sudo install -m 755 bin/officecli bin/landrun /usr/local/bin/
uv sync
uv run python main.py
```

### Templates

[templates/](templates/) holds the stored designs (`.pptx`, `.docx`, `.xlsx`).
At startup their sample slides and body text are removed; slide masters,
layouts and styles stay, and workbooks stay as they are. Mount
`-v ./templates:/app/templates` to swap templates without rebuilding, and
restart the server afterwards.

### Prebuilt image

CI publishes `ghcr.io/th3r3alduk3/owui-office-mcp` on pushes to `main`
(`latest`) and `dev` (`dev`) and on version tags. To run it instead of
building, replace `build:` in [compose.yaml](compose.yaml) with that `image:`.

The image adds a headless Chromium and fonts for previews. Behind mirrors,
[bin/download.sh](bin/download.sh) is the only build step reaching GitHub;
the `[tool.uv]` block in [pyproject.toml](pyproject.toml) selects the package
index.

## 🛠️ Tools

| Tool | Description |
|---|---|
| `list_templates` | Stored templates with their format and, for PPTX, their slide masters |
| `start_project` | Project from a stored template or an attached file: its design only, or the file as it is |
| `get_reference` | OfficeCLI reference for an element, a verb or both |
| `run_commands` | One atomic batch of OfficeCLI commands; reads return their output |
| `preview_project` | The document as thumbnails, or one slide or page, as an image |
| `export_project` | Upload the document to OpenWebUI as a download; again after later edits |

```json
{
  "project_id": "3f9a1c2e5b7d9f01",
  "commands": [
    {"command": "add", "parent": "/", "type": "slide",
     "props": {"layout": "3", "title": "Revenue 2026"}},
    {"command": "add", "parent": "/slide[1]", "type": "chart",
     "props": {"chartType": "bar", "categories": "Q1,Q2,Q3", "data": "Revenue:215,250,280"}}
  ]
}
```

- `project_id`: from `start_project`, which also returns the design's
  inventory: layouts and slots, styles or sheets.
- `keep_content`: edit an attached file as it is; without it only its design
  is used and its slides or body text are dropped. Workbooks keep their
  content either way.
- `slide_master`: a `pptx` project is bound to one slide master; with
  several, `start_project` needs it and lists them when it is missing.
- `commands`: the OfficeCLI batch shape. OfficeCLI validates and answers
  with its own errors; `get_reference` documents elements and props. Images
  come from attached files as `file:<file_id>`.

`start_project` returns a `hint` with the next step. A project is a file of
the user in OpenWebUI (`project-<project_id>.pptx`, `.docx` or `.xlsx`) that
stays until deleted there, like the exported files.

## ⚙️ Limits & security

[.env.example](.env.example) lists all settings: admission and rate limits per
user, OfficeCLI processes, memory and timeouts.

- **Users:** every request carries the user's Keycloak access token;
  projects and exports belong to the OpenWebUI account it is exchanged for,
  so only that user and OpenWebUI admins reach them.
- **Isolation:** every OfficeCLI run starts through landrun and is confined by
  Landlock to its scratch directory, read-only system paths and no TCP; UDP
  stays open until Landlock ABI 10 (Linux 7.2) reaches landrun.
  `OFFICECLI_MAX_MEMORY` bounds its heap, so a pathological document kills
  its own run.
- **Commands:** OfficeCLI validates the batch itself; its batch verbs work
  inside the document, never on files. A PPTX project's slides and
  placeholders stay within its bound slide master. Batches hold up to 500
  commands, 1 MB and 10 attached files; `run_commands` answers are cut at
  100 KB.
- **Files:** attachments and documents stop past 100 MB, images past 20 MB;
  macro and template variants are refused.
- **Images:** libvips decodes attached images in the server process, with
  its unfuzzed loaders blocked. Except for animations, they lose their EXIF,
  XMP and IPTC metadata, and those past 1 MB are scaled to 2560 px.
- **Drafts:** a batch is written back only when it succeeded as a whole, as
  a new file of the project's name; the previous one is then deleted. Of two
  edits of one project at once, the first saved wins; the other is discarded
  with a request to repeat it. OpenWebUI stores
  projects and exports unprocessed, without extraction or embeddings.
- **Errors:** tool errors say what to do next and carry no exception text.
- **Sizing:** the container's 8 GB cover four OfficeCLI runs with 1 GB heap
  and Chromium each, plus 2 GB for Python and the tmpfs scratch, which hold
  `MAX_CONCURRENT_REQUESTS` documents at once; scale it with
  `OFFICECLI_MAX_PROCESSES`. Previews of large decks need the 1 GB heap.
  Limits are per server process.

## 🧩 Layout

- [`main.py`](main.py) - FastMCP server: auth, rate and response limits, lifespan, tools
- [`config.py`](config.py) - settings, read from `.env`
- [`models/office.py`](models/office.py) - the command shape, inventories and results
- [`services/owui.py`](services/owui.py) - OpenWebUI token exchange and files
- [`services/officecli.py`](services/officecli.py) - OfficeCLI runs under landrun
- [`services/image.py`](services/image.py) - attached images: shrunk and stripped of metadata
- [`services/pptx.py`](services/pptx.py), [`docx.py`](services/docx.py), [`xlsx.py`](services/xlsx.py) - one module per format: preparation and inventory; pptx also the slide master binding
- [`services/project.py`](services/project.py) - projects as OpenWebUI files, templates, lifespan
- [`tools/`](tools/) - the six MCP tools and the batch builder
- [`templates/`](templates/) - stored designs
- [`docker/`](docker/) - the server image; [`compose.yaml`](compose.yaml) runs it
