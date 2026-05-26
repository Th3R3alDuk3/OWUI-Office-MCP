import io
from collections.abc import AsyncIterator
from contextlib import asynccontextmanager, suppress
from dataclasses import dataclass
from typing import Any

from fastmcp import FastMCP
from fastmcp.server.dependencies import get_access_token
from pptx import Presentation
from pptx.presentation import Presentation as PresentationType
from pydantic import Field

from config import get_settings
from models.project import ProjectInfo, SavedProjectInfo, SlideInfo
from models.template import TemplateInfo
from subservers.powerpoint._files import upload_to_owui
from subservers.powerpoint._utils import (
    analyze_templates,
    drop_all_slides,
    drop_slide,
)


@dataclass
class _Project:
    template_name: str
    pptx: PresentationType


_projects: dict[str, _Project] = {}
_templates: dict[str, TemplateInfo] = {}


def _user_key() -> str:
    token = get_access_token()
    if token is None:
        raise ValueError("Authentication required.")
    user_id = token.claims.get("id")
    if not user_id:
        raise ValueError("JWT is missing the 'id' claim.")
    return str(user_id)


def _bearer_token() -> str:
    token = get_access_token()
    if token is None:
        raise ValueError("Authentication required.")
    return token.token


@asynccontextmanager
async def lifespan(server: FastMCP) -> AsyncIterator[dict[str, Any]]:
    global _templates
    _templates = analyze_templates(get_settings().templates_dir)
    yield {"templates": _templates}


mcp = FastMCP(name="powerpoint", lifespan=lifespan)


@mcp.tool(
    name="list_templates",
    description=(
        "List all PowerPoint templates (.pptx) that were loaded and analyzed from "
        "the configured templates directory at server startup. Returns a mapping "
        "of template name to its info (absolute path, slide count, and all slide "
        "layouts including their placeholders: index, name, type). Use this tool "
        "to discover which templates are available and which layouts/placeholders "
        "can be filled when generating new slides."
    ),
)
async def list_templates() -> dict[str, TemplateInfo]:
    return _templates


@mcp.tool(
    name="create_project",
    description=(
        "Create a new in-memory PowerPoint project for the current chat session, "
        "based on one of the available templates. The template is loaded fresh "
        "into memory and stored under the active session id, so subsequent tool "
        "calls in this chat can modify it. Pass the template `name` exactly as "
        "returned by `list_templates`. If a project already exists for this "
        "session it is overwritten."
    ),
)
async def create_project(
    template_name: str = Field(
        description="Name of the template to use, as returned by `list_templates`."
    ),
) -> ProjectInfo:
    template = _templates.get(template_name)
    if template is None:
        raise ValueError(
            f"Template '{template_name}' not found."
            f" Available: {list(_templates)}"
        )

    user_id = _user_key()
    pptx = Presentation(template.path)
    drop_all_slides(pptx)
    _projects[user_id] = _Project(template_name=template_name, pptx=pptx)

    return ProjectInfo(
        user_id=user_id,
        template_name=template_name,
        slide_count=len(pptx.slides),
    )


@mcp.tool(
    name="append_slide",
    description=(
        "Append a new slide to the current session's project. The slide is "
        "created from one of the slide layouts defined in the project's "
        "template (see `list_templates` for layout names and their "
        "placeholders). Optionally fill text placeholders by passing a "
        "mapping from placeholder index (`idx`) to the text value to "
        "insert; placeholders not in the mapping are left empty. Returns "
        "the zero-based index of the appended slide within the project."
    ),
)
async def append_slide(
    layout_name: str = Field(
        description="Name of the slide layout to use for the new slide."
    ),
    placeholders: dict[int, str] = Field(
        default_factory=dict,
        description=(
            "Optional mapping from placeholder index (`idx`) to the text value to insert."
        ),
    ),
) -> SlideInfo:

    project = _projects.get(_user_key())
    if project is None:
        raise ValueError(
            "No project for this user."
            " Call `create_project` first."
        )

    layout = project.pptx.slide_layouts.get_by_name(layout_name)
    if layout is None:
        available = [lo.name for lo in project.pptx.slide_layouts]
        raise ValueError(
            f"Layout '{layout_name}' not found in template "
            f"'{project.template_name}'. Available: {available}"
        )

    slide = project.pptx.slides.add_slide(layout)

    for idx, text in placeholders.items():
        with suppress(KeyError):
            slide.placeholders[idx].text = text

    return SlideInfo(
        index=len(project.pptx.slides) - 1,
        layout_name=layout_name,
    )


@mcp.tool(
    name="remove_slides",
    description=(
        "Remove one or more slides from the current session's project by "
        "their zero-based indices. Indices refer to the slide positions "
        "before removal; the call validates all indices first, then "
        "removes them in descending order so earlier indices stay stable. "
        "Returns the updated project info with the new slide count."
    ),
)
async def remove_slides(
    indices: list[int] = Field(
        description=(
            "Zero-based slide indices to remove. Duplicates are ignored."
        ),
    ),
) -> ProjectInfo:

    user_id = _user_key()
    project = _projects.get(user_id)
    if project is None:
        raise ValueError(
            "No project for this user."
            " Call `create_project` first."
        )

    count = len(project.pptx.slides)
    out_of_range = [i for i in indices if not 0 <= i < count]
    if out_of_range:
        raise ValueError(
            f"Slide indices out of range {out_of_range}."
            f" Project has {count} slide(s) (valid: 0..{count - 1})."
        )

    for i in sorted(set(indices), reverse=True):
        drop_slide(project.pptx, i)

    return ProjectInfo(
        user_id=user_id,
        template_name=project.template_name,
        slide_count=len(project.pptx.slides),
    )


@mcp.tool(
    name="save_project",
    description=(
        "Serialize the current user's project to a `.pptx` byte stream and "
        "upload it to OpenWebUI using the caller's JWT. Pass `filename` "
        "without extension (`.pptx` is appended automatically); if omitted, "
        "the template name is used. The project stays active after saving, "
        "so further `append_slide` calls and re-saving are possible. Returns "
        "the chosen filename, the slide count, and the OpenWebUI file id of "
        "the uploaded file."
    ),
)
async def save_project(
    filename: str = Field(
        default="",
        description=(
            "File name without extension. Defaults to the template name "
            "when empty. Existing files with the same name are overwritten."
        ),
    ),
) -> SavedProjectInfo:
    project = _projects.get(_user_key())
    if project is None:
        raise ValueError(
            "No project for this user."
            " Call `create_project` first."
        )

    stem = filename.strip() or project.template_name
    out_name = f"{stem}.pptx"

    buf = io.BytesIO()
    project.pptx.save(buf)

    base_url = get_settings().owui_base_url
    uploaded = await upload_to_owui(
        filename=out_name,
        data=buf.getvalue(),
        token=_bearer_token(),
        base_url=base_url,
    )

    file_id = uploaded.id
    return SavedProjectInfo(
        filename=out_name,
        slide_count=len(project.pptx.slides),
        owui_file_id=file_id,
        owui_url=f"{base_url.rstrip('/')}/api/v1/files/{file_id}/content",
    )
