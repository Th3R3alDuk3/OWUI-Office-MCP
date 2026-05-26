# AGENTS.md

## Projekt

Neuimplementierung eines PowerPoint MCP Servers (Referenz / Anti-Vorbild: https://github.com/GongRzhe/Office-PowerPoint-MCP-Server).

Ziel: sauber, minimal, erweiterbar. **KISS** ist Pflicht — keine Abstraktionen auf Vorrat, keine Features die nicht im Scope stehen, keine Fallbacks für Szenarien die nicht eintreten können.

## Tech Stack

- **FastMCP** als MCP-Framework (Server, Mounting, Tools).
- **python-pptx** für die Analyse / Verarbeitung von `.pptx` Dateien.
- **pydantic-settings** zum Lesen der `.env` (kein selbstgebautes Config-Parsing).
- Python 3.11+.

## Werkzeuge & Philosophie

- **KISS, aber modern**: Stack soll schlank bleiben, aber aktuelle, gepflegte Packages nutzen. Nichts selbst bauen, was ein etabliertes Package gut macht.
- **Bekannte Packages bevorzugen** (FastMCP, python-pptx, pydantic / pydantic-settings, httpx, structlog/loguru o.ä. wenn Logging gebraucht wird). Keine Eigenimplementierungen für gelöste Probleme (JWT-Decode, Env-Parsing, HTTP, …).
- **`uv` ist das Standardwerkzeug** für alles rund um Python in diesem Projekt:
  - Projekt-Init: `uv init`
  - Dependencies hinzufügen: `uv add <package>` (kein manuelles Editieren von `pyproject.toml` für Deps).
  - Dev-Deps: `uv add --dev <package>`.
  - Lock / Sync: `uv lock`, `uv sync`.
  - Skripte / Server starten: `uv run python main.py` bzw. `uv run fastmcp …`.
  - Python-Version pinnen via `uv python pin`.
- **Keine** `requirements.txt`, kein `pip install`, kein `venv` manuell — `uv` managed alles inkl. `.venv` und `uv.lock`.
- `pyproject.toml` + `uv.lock` werden committet, `.venv/` nicht.

## Frontend & Auth

- **Frontend**: [OpenWebUI](https://github.com/open-webui/open-webui). Der MCP-Server wird über die OpenWebUI "Integrations" angebunden.
- **Transport**: FastMCP läuft als **`streamable-http`** (`mcp.run(transport="streamable-http", host=…, port=…)`). Kein stdio, kein SSE-legacy.
- `HOST` und `PORT` kommen aus `.env` via `config.py`.
- OpenWebUI liefert pro Request einen **Access Token (JWT)** im `Authorization: Bearer …` Header.
- OpenWebUI signiert JWTs **symmetrisch** mit dem `WEBUI_SECRET_KEY` (HS256). Daher reicht serverseitig:
  - `JWT_SECRET` — derselbe Shared Secret wie OpenWebUIs `WEBUI_SECRET_KEY`.
  - `JWT_ALGORITHM` — default `HS256`.
- Validierung via FastMCPs `JWTVerifier` (`fastmcp.server.auth.providers.jwt`). Der Shared Secret wird als `public_key`-Parameter übergeben (laut Docstring: *"shared secret for symmetric algorithms"*).
- Kein Issuer / Audience Check — OpenWebUI setzt diese Claims standardmäßig nicht.
- Ungültige / fehlende Tokens → Request wird von FastMCP abgelehnt; keine eigene Auth-Schicht bauen.
- Auth wird auf dem **Haupt-Server** in `main.py` konfiguriert und gilt damit auch für alle gemounteten Subserver.

## Aktueller Scope

1. FastMCP Server startet via `main.py` (`streamable-http`, JWT-Auth).
2. Konfiguration kommt aus `.env` (via `config.py`).
3. Beim Start lädt der **powerpoint** Subserver alle Templates aus `TEMPLATES_DIR` und analysiert sie mit `python-pptx`.
4. Pro User (JWT-Claim `id`) wird ein In-Memory-Projekt gehalten; Tools mutieren dieses Projekt:
   - `list_templates` — beim Start analysierte Templates zurückgeben.
   - `create_project` — leeres Projekt aus einem Template anlegen (überschreibt existierendes).
   - `append_slide` — Folie aus einem Layout anhängen, Placeholder per `idx` füllen.
   - `remove_slides` — Folien per Indexliste löschen.
   - `save_project` — Projekt im Speicher serialisieren und per User-JWT an OpenWebUI (`POST {OPENWEBUI_BASE_URL}/api/v1/files/`) hochladen. Kein lokaler Disk-Write. Antwort enthält `openwebui_file_id`.

State-Lifecycle: kein Auto-Cleanup. `_projects[user_id]` lebt bis Server-Restart oder bis derselbe User `create_project` erneut aufruft.

## Projektstruktur

```
PPTX-MCP/
├── .env                  # Lokale Konfiguration (nicht committen)
├── .env.example          # Vorlage mit allen erwarteten Variablen
├── AGENTS.md
├── main.py               # FastMCP Entry, mountet Subserver
├── config.py             # Lädt .env, stellt Settings-Objekt bereit
├── models/               # Pydantic-Modelle für Tool-Returns
│   └── __init__.py
├── services/             # Wiederverwendbare Service-Klassen (später)
│   └── __init__.py
└── subservers/           # Mountbare FastMCP-Subserver
    ├── __init__.py
    └── powerpoint/
        ├── __init__.py
        ├── server.py     # FastMCP-Instanz + Tools des Subservers
        └── _utils.py     # Template-Loading + python-pptx Helpers
```

## Modulverantwortlichkeiten

### `main.py`
- Erstellt die Haupt-`FastMCP`-Instanz.
- Mountet die Subserver aus `subservers/` (aktuell nur `powerpoint`).
- Startet den Server (`mcp.run()` o.ä.).
- Keine Business-Logik.

### `config.py`
- Lädt `.env`.
- Exportiert ein Settings-Objekt (Pydantic `BaseSettings`) mit:
  - `HOST: str`, `PORT: int` — HTTP-Bind.
  - `TEMPLATES_DIR: Path` — Ordner mit `.pptx` Templates.
  - `JWT_SECRET: str`, `JWT_ALGORITHM: str` (default `HS256`) — Shared Secret für OpenWebUI-Tokens.
  - `OWUI_BASE_URL: str` — Base-URL der OpenWebUI-Instanz für File-Upload (z.B. `http://localhost:3000`).
- Settings sind via `@lru_cache` ein Singleton (`get_settings()`).

### `subservers/powerpoint/`
- `server.py`:
  - Eigene `FastMCP`-Instanz mit Lifespan, der Templates per `analyze_templates` einliest und in das Module-Global `_templates` legt.
  - Module-Global `_projects: dict[str, _Project]` — Key ist die User-ID aus dem JWT.
  - `_user_key()`: liest den aktuellen `AccessToken` via `fastmcp.server.dependencies.get_access_token()` und gibt `claims["id"]` zurück (OpenWebUI-User-UUID). Wirft, wenn keiner da ist.
  - Tools: `list_templates`, `create_project`, `append_slide`, `remove_slides`, `save_project`. Kein `ctx: Context` Param — User-Identifikation läuft komplett über `_user_key()`.
- `_utils.py`:
  - `analyze_templates`: Templates aus dem Template-Ordner einlesen, Metadaten extrahieren (Pfad, Slide-Count, Layouts als `dict[str, LayoutInfo]` mit Placeholders). Rückgabe `dict[str, TemplateInfo]` (Key: `path.stem`).
  - `drop_slide(pptx, index)`: einzelne Slide aus `sldIdLst` entfernen inkl. `drop_rel` auf die Relation.
  - `drop_all_slides(pptx)`: schleift `drop_slide` bis leer (Masters/Layouts bleiben).
- `_files.py`:
  - `upload_to_owui(filename, data, token, base_url)`: lädt Bytes per `httpx.AsyncClient` an `{base_url}/api/v1/files/` hoch (Multipart-Field `file`, `Authorization: Bearer <token>`). Antwort wird in ein `OwuiFile`-Pydantic-Modell (siehe `models/owui.py`) geparst.

### `models/`
- Pydantic-Modelle für Tool-Returns (z.B. `TemplateInfo`, `TemplateList`).
- Klein halten; nur was tatsächlich zurückgegeben wird.

### `services/`
- Aktuell leer (nur `__init__.py`). Platzhalter für spätere wiederverwendbare Logik (z.B. Slide-Builder, Export-Service). Nicht vorab füllen.

## Konventionen

- Templates werden **einmal beim Start** geladen. Keine Watcher, kein Lazy-Loading, kein Caching-Layer.
- Fehler beim Laden eines einzelnen Templates: kurz loggen und überspringen, Server startet trotzdem. Fehlt der Template-Ordner komplett: hart fehlschlagen.
- Keine Kommentare die beschreiben *was* der Code tut — nur *warum*, falls nicht offensichtlich.
- Keine zusätzlichen Tools, Endpoints, CLI-Flags oder Helper-Skripte ohne expliziten Auftrag.

### Logging

- Nur marginal nutzen. Python-stdlib `logging` reicht, kein extra Package.
- Loggen erlaubt für: Server-Start, fehlgeschlagene Template-Loads, Auth-Konfig-Fehler beim Boot.
- Erwartbare/uninteressante Exceptions mit `contextlib.suppress(...)` statt `try/except: pass` oder Log-Spam.
- Kein Debug-Logging in normalem Code-Pfad.

### Type Hints

- **Pflicht** auf allen Funktions-Signaturen (Args + Return) und Attributen.
- Modernes Python: `list[str]`, `dict[str, X]`, `X | None` statt `typing.List`, `Optional`.
- Returns von Tools sind Pydantic-Modelle aus `models/`, keine nackten dicts.
- Kein `Any`, außer es ist wirklich unvermeidbar.

### Was es **nicht** gibt

- Kein Linter, kein Formatter (kein ruff/black/mypy als Pflicht-Gate).
- Keine Tests in dieser Iteration.
- Kein CI.

### `.gitignore`

Mindestens:
```
.venv/
.env
__pycache__/
*.pyc
.python-version
```

`pyproject.toml` und `uv.lock` werden committet, `.env.example` auch.
