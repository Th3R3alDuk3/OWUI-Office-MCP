# pptx-mcp

Minimaler MCP-Server für PowerPoint-Templates. FastMCP + python-pptx. Frontend: OpenWebUI.

## Setup

```bash
uv sync
cp .env.example .env
# JWT_SECRET = OpenWebUIs WEBUI_SECRET_KEY
# OWUI_BASE_URL = z.B. http://localhost:3000 (vom MCP-Server aus erreichbar)
```

`.pptx` Templates in [templates/](templates/) ablegen.

## Run

```bash
uv run python main.py
```

Läuft als `streamable-http` auf `HOST:PORT` aus `.env`. JWT-Auth über OpenWebUI's Shared Secret (HS256).

## Tools

| Tool | Beschreibung |
|---|---|
| `powerpoint_list_templates` | Listet beim Start analysierte Templates (Name, Pfad, Slide-Count, Layouts + Placeholders) |
| `powerpoint_create_project` | Legt ein leeres In-Memory-Projekt für den User an (Slides aus dem Template werden entfernt, Masters/Layouts bleiben) |
| `powerpoint_append_slide` | Hängt eine Folie aus einem Layout an; optional Placeholder per `idx` füllen |
| `powerpoint_remove_slides` | Entfernt Folien per 0-basiertem Index (Liste, Duplikate ignoriert) |
| `powerpoint_save_project` | Serialisiert das Projekt im Speicher und lädt es per User-JWT zu OpenWebUI hoch (`/api/v1/files/`) — kein Disk-Write |

Projekt-State wird pro User gehalten — Key kommt aus dem JWT-Claim `id` (OpenWebUI User-UUID). State überlebt damit einzelne Requests innerhalb derselben User-Session.

## Struktur

```
main.py              FastMCP entry, Auth, mountet Subserver
config.py            .env via pydantic-settings
models/              Pydantic-Returns
services/            (placeholder)
subservers/
  powerpoint/        Lifespan-Loading + Tools (list/create/append/remove/save)
```

Details: siehe [AGENTS.md](AGENTS.md).
