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
| ➕ `append_slide` | Folie aus einem Layout anhängen |
| ✏️ `edit_slide` | Placeholder einer Folie ändern |
| 🗑️ `remove_slides` | Folien per Index entfernen |
| 💾 `download_project` | Projekt zu OpenWebUI hochladen |

### 📄 `docx`

| Tool | |
|---|---|
| 📋 `list_templates` | verfügbare `.docx` Templates |
| 🎨 `list_styles` | Paragraph- und Table-Styles eines Templates |
| 📁 `create_project` | leeres Projekt aus einem Template |
| ➕ `append_paragraph` | Paragraph anhängen (Heading via Style-Name) |
| 📊 `append_table` | Tabelle anhängen (optional mit Zelldaten) |
| 🗑️ `remove_blocks` | Body-Blöcke (Paragraph & Table) per Index entfernen |
| 💾 `download_project` | Projekt zu OpenWebUI hochladen |

## 📚 Mehr

Architektur & Konventionen → 👉 [AGENTS.md](AGENTS.md)
