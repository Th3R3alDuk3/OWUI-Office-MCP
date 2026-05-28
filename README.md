# 🧰 OWUI-Office-MCP

> Office-Dokumente via MCP für OpenWebUI. Schlank, modern, erweiterbar.

⚡ FastMCP &nbsp;·&nbsp; 📦 uv &nbsp;·&nbsp; 🔐 JWT (HS256) &nbsp;·&nbsp; 🌐 streamable-http

---

## 🎯 Status

| Format | Status |
|---|---|
| 🎬 PowerPoint `.pptx` | ✅ |
| 📄 Word `.docx` | ✅ |
| 📊 Excel `.xlsx` | 🚧 |

## 🚀 Setup

```bash
uv sync
cp .env.example .env
```

In der `.env`:
- 🔑 `JWT_SECRET` → OpenWebUIs `WEBUI_SECRET_KEY`
- 🌐 `OWUI_BASE_URL` → z.B. `http://localhost:3000` (vom MCP-Server aus erreichbar)

Templates in 📁 [templates/](templates/) ablegen.

## ▶️ Run

```bash
uv run python main.py
```

Läuft als `streamable-http` auf `HOST:PORT` aus der `.env`.

## 🛠️ Tools

Jeder Subserver ist unter seiner Dateiendung als Namespace gemounted (`pptx.*`, `docx.*`). Per-User State pro Subserver (JWT-Claim `id`), Sliding-TTL, Auto-Sweep — keine Disk-Writes.

### 🎬 `pptx`

| Tool | |
|---|---|
| 📋 `list_templates` | verfügbare `.pptx` Templates |
| 🎨 `list_masters` | Slide Masters eines Templates |
| 📐 `list_layouts` | Layouts + Placeholders eines Masters |
| 📁 `create_project` | leeres Projekt aus einem Template |
| ➕ `insert_slide` | Folie aus einem Layout einfügen (optional an Index, sonst anhängen) |
| 🔍 `list_slides` | Folien auflisten (Index, Layout, Text) |
| ✏️ `edit_slide` | Placeholder einer Folie ändern |
| ↕️ `move_slide` | Folie per Index an neue Position verschieben |
| 🗑️ `remove_slides` | Folien per Index entfernen |
| 💾 `download_project` | Projekt zu OpenWebUI hochladen |

### 📄 `docx`

| Tool | |
|---|---|
| 📋 `list_templates` | verfügbare `.docx` Templates |
| 🎨 `list_styles` | Paragraph- und Table-Styles eines Templates |
| 📁 `create_project` | leeres Projekt aus einem Template |
| ➕ `insert_paragraph` | Paragraph einfügen (optional an Index, sonst anhängen) |
| 📊 `insert_table` | Tabelle einfügen (optional mit Zelldaten, optional an Index) |
| 📃 `insert_page_break` | Seitenumbruch als Body-Block einfügen (optional an Index) |
| 🔍 `list_blocks` | Body-Blöcke auflisten (Index, Typ, Text-Preview) |
| ✏️ `edit_paragraph` | Text eines Paragraphs per Index ändern (Style bleibt) |
| ↕️ `move_block` | Body-Block (Paragraph & Table) per Index verschieben |
| 🗑️ `remove_blocks` | Body-Blöcke (Paragraph & Table) per Index entfernen |
| 💾 `download_project` | Projekt zu OpenWebUI hochladen |

## 📚 Mehr

Architektur & Konventionen → 👉 [AGENTS.md](AGENTS.md)
