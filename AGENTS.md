# AGENTS.md

## Projekt

Neuimplementierung eines Office MCP Servers für OpenWebUI (Referenz / Anti-Vorbild: https://github.com/GongRzhe/Office-PowerPoint-MCP-Server). Unterstützte Formate: PowerPoint (`.pptx`) und Word (`.docx`). Excel (`.xlsx`) folgt — jeder Office-Typ ist ein eigener Subserver unter `subservers/<endung>/`, gemounted in `main.py` mit Dateiendung als Namespace.

Ziel: sauber, minimal, erweiterbar. **KISS** ist Pflicht — keine Abstraktionen auf Vorrat, keine Features die nicht im Scope stehen, keine Fallbacks für Szenarien die nicht eintreten können.

## Tech Stack

- **FastMCP** als MCP-Framework (Server, Mounting, Tools).
- **python-pptx** / **python-docx** für die Analyse / Verarbeitung von `.pptx` bzw. `.docx` Dateien.
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
3. Templates aus `TEMPLATES_DIR` werden **lazy** pro Tool-Call mit `python-pptx` bzw. `python-docx` eingelesen — kein Preload, kein Cache. Beide Subserver teilen sich `TEMPLATES_DIR` und filtern nach Dateiendung.
4. Pro User (JWT-Claim `id`) und Subserver wird ein eigenes In-Memory-Projekt gehalten; Tools mutieren das jeweilige Projekt.

   **`pptx`** (Namespace `pptx`):
   - `list_templates` — verfügbare `.pptx` Dateien im Template-Ordner listen.
   - `list_masters` — Slide Masters eines Templates listen (`dict[int, str]`, Key ist der Index).
   - `list_layouts` — Layouts eines Masters mit Placeholder-Infos (`idx`, `name`, `type`).
   - `create_project` — leeres Projekt aus einem Template anlegen (überschreibt existierendes).
   - `insert_slide` — Folie aus einem Layout einfügen (optional an `slide_index`, sonst anhängen), Placeholder per `idx` füllen.
   - `list_slides` — Folien in Reihenfolge auflisten (Index, Layout, Text) als Grundlage für edit/move/remove per Index.
   - `edit_slide` — nur die Placeholder-Inhalte einer bestehenden Folie (per Index) aktualisieren; nicht aufgeführte Placeholder bleiben unverändert.
   - `move_slide` — Folie per Index an eine neue Position verschieben.
   - `remove_slides` — Folien per Indexliste löschen.
   - `download_project` — Projekt im Speicher serialisieren und per User-JWT an OpenWebUI (`POST {OWUI_BASE_URL}/api/v1/files/`) hochladen. Kein lokaler Disk-Write. Antwort enthält `owui_url`.

   **`docx`** (Namespace `docx`):
   - `list_templates` — verfügbare `.docx` Dateien im Template-Ordner listen.
   - `list_styles` — Paragraph- und Table-Styles eines Templates als `dict[str, StyleInfo]` (type, builtin). Andere Style-Typen werden gefiltert.
   - `create_project` — leeres Projekt aus einem Template anlegen (Body geleert, `<w:sectPr>` bleibt erhalten).
   - `insert_paragraph` — Paragraph mit optionalem Style einfügen (optional an `block_index`, sonst anhängen; Headings via Style-Name, z.B. `Heading 1`).
   - `insert_table` — Tabelle (`rows` × `cols`, Grenzen via `Field`) mit optionalem Style und optionalen Zelldaten einfügen (optional an `block_index`).
   - `insert_page_break` — Seitenumbruch als Body-Block einfügen (optional an `block_index`).
   - `list_blocks` — Body-Blöcke in Reihenfolge auflisten (Index, Typ, Text-Preview) als Grundlage für edit/move/remove per Index.
   - `edit_paragraph` — Text eines Paragraphen per Block-Index ersetzen; der Index umfasst alle Body-Blöcke (Paragraphs **und** Tables), Tabellen-Blöcke werden abgelehnt.
   - `move_block` — Body-Block (Paragraph oder Table) per Index an eine neue Position verschieben.
   - `remove_blocks` — Body-Blöcke (Paragraphs **und** Tables, gemeinsamer Index) per Indexliste löschen.
   - `download_project` — analog zu pptx; Antwort enthält `block_count` + `owui_url`.

State-Lifecycle: `_projects` ist eine `cachetools.TTLCache` (Sliding-TTL via Re-Insert nach jedem Tool-Call, harter `maxsize=1_000`-Cap gegen unbegrenztes Wachstum). Lazy Eviction beim Zugriff plus aktiver Background-Sweep via `_ttl_task` → `cache.expire()`. TTL und Sweep-Intervall kommen aus `.env` (`PROJECT_TTL_SECONDS`, `PROJECT_SWEEP_INTERVAL_SECONDS`, Defaults 3600 / 300).

Concurrency: jedes `Project` hat ein `asyncio.Lock`-Feld. Alle mutierenden Tools wrappen ihren Mutationsblock in `async with project.lock:`, sodass python-pptx/-docx-Aufrufe auf demselben Projekt serialisiert sind. In `download_project` wird der Upload-Call **außerhalb** des Locks ausgeführt — der `BytesIO`-Snapshot entsteht innerhalb des Locks, der I/O-bound Upload blockiert danach keine weiteren Mutationen. OWUI-Fehler (`httpx.HTTPStatusError`, `httpx.RequestError`) werden in eine `RuntimeError` mit lesbarer Meldung umgewandelt.

Auth: kein manuelles `get_access_token()` in Tools. Stattdessen FastMCP-DI via Default-Args — `user_id: str = TokenClaim("id")`, `project: Project = Depends(_get_project)`, ggf. `token: AccessToken = CurrentAccessToken()`. Der Helper `_get_project` löst `_projects[user_id]` auf und raised `ValueError` wenn kein Projekt existiert.

## Projektstruktur

```
OWUI-Office-MCP/
├── .env                  # Lokale Konfiguration (nicht committen)
├── .env.example          # Vorlage mit allen erwarteten Variablen
├── AGENTS.md
├── main.py               # FastMCP Entry, mountet Subserver
├── config.py             # Lädt .env, stellt Settings-Objekt bereit
├── models/               # Pydantic-Modelle pro Subserver (+ shared)
│   ├── __init__.py
│   ├── owui.py           # OpenWebUI Upload-Response (shared)
│   ├── pptx.py           # Project, DownloadProjectResponse, LayoutInfo, PlaceholderInfo, SlideInfo
│   └── docx.py           # Project, DownloadProjectResponse, StyleInfo, BlockInfo
├── services/             # Subserver-übergreifende Service-Funktionen
│   ├── __init__.py
│   └── owui.py           # Generischer OpenWebUI File-Upload
└── subservers/           # Mountbare FastMCP-Subserver (Namespace = Dateiendung)
    ├── __init__.py
    ├── pptx/
    │   ├── __init__.py
    │   ├── server.py     # FastMCP-Instanz + Tools des Subservers
    │   └── _utils.py     # python-pptx Helpers (Templates/Slides)
    └── docx/
        ├── __init__.py
        ├── server.py     # FastMCP-Instanz + Tools des Subservers
        └── _utils.py     # python-docx Helpers (Templates/Blocks)
```

## Modulverantwortlichkeiten

### `main.py`
- Erstellt die Haupt-`FastMCP`-Instanz.
- Mountet die Subserver aus `subservers/` (`pptx`, `docx`) jeweils unter ihrem Dateiendungs-Namespace.
- Kombiniert die Subserver-Lifespans via `combine_lifespans`.
- Startet den Server (`mcp.run()` o.ä.).
- Keine Business-Logik.

### `config.py`
- Lädt `.env`.
- Exportiert ein Settings-Objekt (Pydantic `BaseSettings`) mit:
  - `HOST: str`, `PORT: int` — HTTP-Bind.
  - `TEMPLATES_DIR: Path` — Ordner mit `.pptx`/`.docx` Templates (von beiden Subservern geteilt, gefiltert nach Endung).
  - `JWT_SECRET: str`, `JWT_ALGORITHM: str` (default `HS256`) — Shared Secret für OpenWebUI-Tokens.
  - `OWUI_BASE_URL: str` — Base-URL der OpenWebUI-Instanz für File-Upload (z.B. `http://localhost:3000`).
  - `PROJECT_TTL_SECONDS: int` (default 3600) — Sliding-TTL für In-Memory-Projekte pro User.
  - `PROJECT_SWEEP_INTERVAL_SECONDS: int` (default 300) — Intervall für den aktiven Sweep, der abgelaufene Einträge aus dem Cache räumt.
- Settings sind via `@lru_cache` ein Singleton (`get_settings()`).

### `subservers/pptx/`
- `server.py`:
  - Eigene `FastMCP`-Instanz mit Lifespan, der ausschließlich den TTL-Sweep-Task (`_ttl_task`) verwaltet. Kein Template-Preload — Templates werden lazy pro Tool-Call eingelesen.
  - Module-Global `_projects: TTLCache[str, Project]` — Key ist die User-ID aus dem JWT.
  - Konstante `PPTX_MIME` (PPTX-Content-Type), wird beim Upload an `services.owui.upload_file` weitergereicht.
  - Tools (alle `async`): `list_templates`, `list_masters`, `list_layouts`, `create_project`, `insert_slide`, `list_slides`, `edit_slide`, `move_slide`, `remove_slides`, `download_project`. Auth-Werte kommen via FastMCP-DI als Default-Args rein (`TokenClaim("id")`, `Depends(_get_project)`, `CurrentAccessToken()`) — keine manuelle Token-Behandlung im Tool-Body.
  - Helper `_get_project(user_id: str = TokenClaim("id")) -> Project` — DI-Factory, liefert das Projekt aus `_projects` oder raised `ValueError` mit dem Hinweis auf `create_project`.
- `_utils.py`:
  - `list_template_names(templates_dir)`: `.pptx` Dateien im Template-Ordner listen (defekte überspringen + warn).
  - `list_master_names(templates_dir, template_name)`: Slide Masters → `dict[int, str]`. Name kommt aus `master.name`, fällt zurück auf den Theme-Namen via `RELATIONSHIP_TYPE.THEME`.
  - `list_layout_infos(templates_dir, template_name, master_index)`: Layouts eines Masters als `dict[str, LayoutInfo]` inkl. Placeholders (`idx`, `name`, `type`).
  - `count_slides(presentation)`: Anzahl Slides.
  - `list_slide_infos(presentation)`: Slides als `list[SlideInfo]` (Layout + Text).
  - `drop_slides(presentation, indices)`: mehrere Slides per Indexliste in einem Durchlauf aus `sldIdLst` entfernen (inkl. `drop_rel`); nutzt den privaten `_detach_slide`.
  - `drop_all_slides(presentation)`: alle Slides in einem Durchlauf entfernen (Masters/Layouts bleiben).
  - `move_slide(presentation, from_index, to_index)`: Slide in `sldIdLst` umsortieren (negativer `to_index` zählt vom Ende).

### `subservers/docx/`
- `server.py`:
  - Aufbau identisch zu `pptx/server.py`: eigene `FastMCP`-Instanz, eigener `_projects: TTLCache[str, Project]`, eigener Lifespan-TTL-Sweep, eigener `_get_project`-Helper.
  - Konstante `DOCX_MIME` (Word-Content-Type) für den OWUI-Upload.
  - Tools (alle `async`): `list_templates`, `list_styles`, `create_project`, `insert_paragraph`, `insert_table`, `insert_page_break`, `list_blocks`, `edit_paragraph`, `move_block`, `remove_blocks`, `download_project`. DI-Konvention wie beim pptx-Subserver. `insert_paragraph` akzeptiert optional einen Paragraph-Style (Headings über z.B. `Heading 1`) — kein separater Heading-Endpoint. `list_blocks`, `move_block` und `remove_blocks` arbeiten auf einer einheitlichen Block-Liste (Paragraphs **und** Tables, ohne `<w:sectPr>`) → Indizes über beide Block-Typen hinweg.
- `_utils.py`:
  - `list_template_names(templates_dir)`: `.docx` Dateien im Template-Ordner listen (defekte überspringen + warn).
  - `list_style_infos(templates_dir, template_name)`: Paragraph- und Table-Styles als `dict[str, StyleInfo]`. Andere Style-Typen (Character/List) werden gefiltert — sind für die aktuellen Tools nicht relevant.
  - `count_blocks(document)`: Anzahl Body-Blöcke (ohne `<w:sectPr>`).
  - `content_blocks(document)`: Body-Blöcke als python-docx-Objekte (`Paragraph`/`Table`) in Reihenfolge; Basis für `list_blocks` und `edit_paragraph`.
  - `list_block_infos(document)`: Body-Blöcke als `list[BlockInfo]` (Typ + Text-Preview).
  - `move_block(document, from_index, to_index)`: Body-Block umsortieren (negativer `to_index` zählt vom Ende).
  - `drop_blocks(document, indices)` / `drop_all_blocks(document)`: ausgewählte bzw. alle Body-Kinder entfernen, `<w:sectPr>` bleibt erhalten (Word braucht das Section-Properties-Element für ein valides Dokument).

### `models/`
- Pydantic-Modelle für Tool-Returns (z.B. `LayoutInfo`, `PlaceholderInfo`, `SlideInfo`, `StyleInfo`, `BlockInfo`, `DownloadProjectResponse`).
- Klein halten; nur was tatsächlich zurückgegeben wird.

### `services/`
- Wiederverwendbare Service-Funktionen, subserver-übergreifend nutzbar.
- `owui.py`: `upload_file(filename, data, content_type, token, base_url)` — generischer File-Upload an OpenWebUI (`POST {base_url}/api/v1/files/`, Multipart-Field `file`, Bearer-Auth). Antwort wird in `OWUIFile` (siehe `models/owui.py`) geparst. Keine Format-Annahmen — der Content-Type wird vom Aufrufer übergeben.

## Konventionen

- Templates werden **lazy pro Tool-Call** eingelesen (von `python-pptx`/`python-docx`). Kein Preload beim Start, kein Watcher, kein Caching-Layer.
- Fehler beim Laden eines einzelnen Templates (beim Listen via `list_templates`): kurz loggen und überspringen. Fehlt der Template-Ordner komplett: hart fehlschlagen (`RuntimeError`).
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
- Komplexe Tool-Returns sind Pydantic-Modelle aus `models/`; einfache `list[str]`/`dict[...]`-Returns sind okay, wenn sie die Tool-Antwort direkt und klar ausdrücken.
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
