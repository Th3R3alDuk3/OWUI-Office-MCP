# 🧰 OWUI-Office-MCP

> Office-Dokumente via MCP für OpenWebUI. Schlank, modern, erweiterbar.

⚡ FastMCP &nbsp;·&nbsp; 📦 uv &nbsp;·&nbsp; 🔐 JWT (HS256) &nbsp;·&nbsp; 🌐 streamable-http

---

## 🎯 Status

| Format | Status |
|---|---|
| 🎬 PowerPoint `.pptx` | ✅ |
| 📄 Word `.docx` | 🚧 |
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

## 🛠️ Tools (PowerPoint)

| Tool | |
|---|---|
| 📋 `list_templates` | verfügbare Templates |
| 🎨 `list_masters` | Slide Masters eines Templates |
| 📐 `list_layouts` | Layouts + Placeholders eines Masters |
| 📁 `create_project` | leeres Projekt aus einem Template |
| ➕ `append_slide` | Folie aus einem Layout anhängen |
| ✏️ `edit_slide` | Placeholder einer Folie ändern |
| 🗑️ `remove_slides` | Folien per Index entfernen |
| 💾 `download_project` | Projekt zu OpenWebUI hochladen |

Per-User State (JWT-Claim `id`), Sliding-TTL, Auto-Sweep — keine Disk-Writes.

## 📚 Mehr

Architektur & Konventionen → 👉 [AGENTS.md](AGENTS.md)
