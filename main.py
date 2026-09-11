from fastmcp import FastMCP
from fastmcp.server.auth.providers.jwt import JWTVerifier
from fastmcp.server.dependencies import get_access_token
from fastmcp.server.middleware.rate_limiting import RateLimitingMiddleware

from config import get_settings
from office import TOOLS, office_lifespan

settings = get_settings()

INSTRUCTIONS = """
OWUI-Office-MCP creates and edits PowerPoint, Word and Excel files with
OfficeCLI commands.

A user may hold several projects; each tool result names the `project_id`.
A project is started one of two ways:
- `create_project` starts a new, empty document from a design: a stored
  template from `list_templates` (the default when the user attached
  nothing), or a file the user attached as the design, e.g. their own slide
  master — PPTX/DOCX example content is cleared while template structure is
  retained. XLSX keeps sheets, existing values, formulas and formatting.
- `open_project` edits a file the user attached, as it is; use it to keep
  sample slides, cover pages or other existing content.
A `file_id` is an attached OpenWebUI file — never a template name from
`list_templates`, and never invented.

Workflow: start a project, build or edit it with `run_commands` — one batch
of OfficeCLI commands per edit step, applied atomically — check the result
visually with `preview_project`, then call `export_project` once the user's
request is fully applied, not after every individual change. The project
stays editable: apply a later request to the same project and export it
again. Stay within the design: its layouts, styles, theme colors and
fonts. Only when the user explicitly asks for other colors, fonts or
formatting, pass `design_mode="custom"`. `get_reference` documents elements
and props. Every tool result includes a `hint` field with the suggested
next step — follow it unless the user's request says otherwise.
""".strip()

mcp = FastMCP(
    name="OWUI-Office-MCP",
    instructions=INSTRUCTIONS,
    auth=JWTVerifier(
        public_key=settings.jwt_secret,
        algorithm=settings.jwt_algorithm,
    ),
    lifespan=office_lifespan,
    middleware=[
        RateLimitingMiddleware(
            max_requests_per_second=settings.rate_limit_rps,
            burst_capacity=settings.rate_limit_burst,
            # OpenWebUI JWTs carry the user in the `id` claim.
            get_client_id=lambda context: (
                token.claims.get("id", "anonymous")
                if (token := get_access_token()) else "anonymous"
            ),
        ),
    ],
    tools=TOOLS,
    mask_error_details=True,
)


if __name__ == "__main__":
    mcp.run(
        host="0.0.0.0",
        port=8000,
        transport="http",
        # Projects are keyed on the JWT user, never on an MCP session.
        stateless_http=True,
    )
