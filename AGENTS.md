# AGENTS.md

## Projekt

Neuimplementierung eines Office MCP Servers für OpenWebUI (Referenz / Anti-Vorbild: https://github.com/GongRzhe/Office-PowerPoint-MCP-Server). Erstes unterstütztes Format: PowerPoint (`.pptx`). Word (`.docx`) und Excel (`.xlsx`) folgen — jeweils als eigener Subserver unter `subservers/` (analog zu `subservers/powerpoint/`), gemounted in `main.py` mit eigenem Namespace.

Ziel: sauber, minimal, erweiterbar. **KISS** ist Pflicht — keine Abstraktionen auf Vorrat, keine Features die nicht im Scope stehen, keine Fallbacks für Szenarien die nicht eintreten können.

## Tech Stack

- **FastMCP** als MCP-Framework (Server, Mounting, Tools).
- **python-pptx** für die Analyse / Verarbeitung von `.pptx` Dateien.
- **pydantic-settings** zum Lesen der `.env` (kein selbstgebautes Config-Parsing).
- Python 3.12+.

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
3. Templates aus `TEMPLATES_DIR` werden **lazy** pro Tool-Call mit `python-pptx` eingelesen — kein Preload, kein Cache.
4. Pro User (JWT-Claim `id`) wird ein In-Memory-Projekt gehalten; Tools mutieren dieses Projekt:
   - `list_templates` — verfügbare `.pptx` Dateien im Template-Ordner listen.
   - `list_masters` — Slide Masters eines Templates listen (`dict[int, str]`, Key ist der Index).
   - `list_layouts` — Layouts eines Masters mit Placeholder-Infos (`idx`, `name`, `type`).
   - `create_project` — leeres Projekt aus einem Template anlegen (überschreibt existierendes).
   - `append_slide` — Folie aus einem Layout anhängen, Placeholder per `idx` füllen.
   - `edit_slide` — nur die Placeholder-Inhalte einer bestehenden Folie (per Index) aktualisieren; nicht aufgeführte Placeholder bleiben unverändert.
   - `remove_slides` — Folien per Indexliste löschen.
   - `download_project` — Projekt im Speicher serialisieren und per User-JWT an OpenWebUI (`POST {OWUI_BASE_URL}/api/v1/files/`) hochladen. Kein lokaler Disk-Write. Antwort enthält `owui_url`.

State-Lifecycle: `_projects` ist eine `cachetools.TTLCache` (Sliding-TTL via Re-Insert nach jedem Tool-Call, harter `maxsize=10_000`-Cap gegen unbegrenztes Wachstum). Lazy Eviction beim Zugriff plus aktiver Background-Sweep via `_sweep_projects` → `cache.expire()`. TTL und Sweep-Intervall kommen aus `.env` (`PROJECT_TTL_SECONDS`, `PROJECT_SWEEP_INTERVAL_SECONDS`, Defaults 3600 / 300).

Concurrency: jedes `Project` hat ein `asyncio.Lock`-Feld. Alle mutierenden Tools (`append_slide`, `edit_slide`, `remove_slides`, `download_project`) wrappen ihren Mutationsblock in `async with project.lock:`, sodass python-pptx-Aufrufe auf demselben Projekt serialisiert sind. In `download_project` wird der Upload-Call **außerhalb** des Locks ausgeführt — der `BytesIO`-Snapshot entsteht innerhalb des Locks, der I/O-bound Upload blockiert danach keine weiteren Mutationen. OWUI-Fehler (`httpx.HTTPStatusError`, `httpx.RequestError`) werden in eine `RuntimeError` mit lesbarer Meldung umgewandelt.

Auth: kein manuelles `get_access_token()` in Tools. Stattdessen FastMCP-DI via Default-Args — `user_id: str = TokenClaim("id")`, `project: Project = Depends(_get_project)`, ggf. `token: AccessToken = CurrentAccessToken()`. Der Helper `_get_project` löst `_projects[user_id]` auf und raised `ValueError` wenn kein Projekt existiert.

## Projektstruktur

```
OWUI-Office-MCP/
├── .env                  # Lokale Konfiguration (nicht committen)
├── .env.example          # Vorlage mit allen erwarteten Variablen
├── AGENTS.md
├── main.py               # FastMCP Entry, mountet Subserver
├── config.py             # Lädt .env, stellt Settings-Objekt bereit
├── models/               # Pydantic-Modelle für Tool-Returns
│   ├── __init__.py
│   ├── owui.py           # OpenWebUI Upload-Response
│   ├── project.py        # Project, DownloadProjectResponse
│   └── template.py       # LayoutInfo, PlaceholderInfo
├── services/             # Subserver-übergreifende Service-Funktionen
│   ├── __init__.py
│   └── owui.py           # Generischer OpenWebUI File-Upload
└── subservers/           # Mountbare FastMCP-Subserver
    ├── __init__.py
    └── powerpoint/
        ├── __init__.py
        ├── server.py     # FastMCP-Instanz + Tools des Subservers
        └── _utils.py     # python-pptx Helpers (Templates/Slides)
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
  - `PROJECT_TTL_SECONDS: int` (default 3600) — Sliding-TTL für In-Memory-Projekte pro User.
  - `PROJECT_SWEEP_INTERVAL_SECONDS: int` (default 300) — Intervall für den aktiven Sweep, der abgelaufene Einträge aus dem Cache räumt.
- Settings sind via `@lru_cache` ein Singleton (`get_settings()`).

### `subservers/powerpoint/`
- `server.py`:
  - Eigene `FastMCP`-Instanz mit Lifespan, der ausschließlich den TTL-Sweep-Task (`_ttl_task`) verwaltet. Kein Template-Preload — Templates werden lazy pro Tool-Call eingelesen.
  - Module-Global `_projects: TTLCache[str, Project]` — Key ist die User-ID aus dem JWT.
  - Konstante `PPTX_MIME` (PPTX-Content-Type), wird beim Upload an `services.owui.upload_file` weitergereicht.
  - Tools (alle `async`): `list_templates`, `list_masters`, `list_layouts`, `create_project`, `append_slide`, `edit_slide`, `remove_slides`, `download_project`. Auth-Werte kommen via FastMCP-DI als Default-Args rein (`TokenClaim("id")`, `Depends(_get_project)`, `CurrentAccessToken()`) — keine manuelle Token-Behandlung im Tool-Body.
  - Helper `_get_project(user_id: str = TokenClaim("id")) -> Project` — DI-Factory, liefert das Projekt aus `_projects` oder raised `ValueError` mit dem Hinweis auf `create_project`.
- `_utils.py`:
  - `list_template_names(templates_dir)`: `.pptx` Dateien im Template-Ordner listen (defekte überspringen + warn).
  - `list_master_names(templates_dir, template_name)`: Slide Masters → `dict[int, str]`. Name kommt aus `master.name`, fällt zurück auf den Theme-Namen via `RELATIONSHIP_TYPE.THEME`.
  - `list_layout_infos(templates_dir, template_name, master_index)`: Layouts eines Masters als `dict[str, LayoutInfo]` inkl. Placeholders (`idx`, `name`, `type`).
  - `drop_slide(presentation, index)`: einzelne Slide aus `sldIdLst` entfernen inkl. `drop_rel` auf die Relation.
  - `drop_all_slides(presentation)`: schleift `drop_slide` bis leer (Masters/Layouts bleiben).

### `models/`
- Pydantic-Modelle für Tool-Returns (z.B. `TemplateInfo`, `TemplateList`).
- Klein halten; nur was tatsächlich zurückgegeben wird.

### `services/`
- Wiederverwendbare Service-Funktionen, subserver-übergreifend nutzbar.
- `owui.py`: `upload_file(filename, data, content_type, token, base_url)` — generischer File-Upload an OpenWebUI (`POST {base_url}/api/v1/files/`, Multipart-Field `file`, Bearer-Auth). Antwort wird in `OWUIFile` (siehe `models/owui.py`) geparst. Keine Format-Annahmen — der Content-Type wird vom Aufrufer übergeben.

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
