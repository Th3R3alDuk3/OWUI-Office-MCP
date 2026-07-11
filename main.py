from fastmcp import FastMCP
from fastmcp.server.auth.providers.jwt import JWTVerifier
from fastmcp.server.dependencies import get_access_token
from fastmcp.server.middleware.rate_limiting import RateLimitingMiddleware

from config import get_settings
from subservers.docx.server import mcp as docx_mcp
from subservers.pptx.server import mcp as pptx_mcp
from subservers.xlsx.server import mcp as xlsx_mcp

settings = get_settings()

ROOT_INSTRUCTIONS = """
OWUI-Office-MCP exposes three stateful Office toolsets:
- Use the `pptx_*` tools for PowerPoint presentations.
- Use the `docx_*` tools for Word documents.
- Use the `xlsx_*` tools for Excel workbooks.

Each user has one in-memory project per toolset. A project is started one of
two ways, each overwriting any existing project for that user in the same
toolset:
- `create_project` starts a new, empty project from a template. This is the
  default: use it whenever the user did not attach a file.
- `open_project` loads a file the user attached in OpenWebUI by its `file_id`.
  Use it only when the user actually attached a file; if none was given, use
  `create_project` instead. A `file_id` is an attached OpenWebUI file — never
  a template name from `list_templates`, and never invented.

Workflow: start a project with `create_project` or `open_project`, build or
edit it with `run_script` — one Python script per edit batch, written
against the sandboxed functions documented in the `run_script` tool — then
call the matching `finalize_project` tool exactly once after the user's
requested batch of edits is complete. Do not call `finalize_project` after
every individual change when multiple changes belong to one request. If the
user later asks for another edit, apply that edit batch to the active
project and call `finalize_project` once again. Every tool result includes
a `hint` field with the suggested next step — follow it unless the user's
request says otherwise.

Mounted tool names are prefixed, for example `pptx_finalize_project` and
`docx_finalize_project`. Tool descriptions may mention local names like
`finalize_project`; use the corresponding prefixed tool exposed by this root
server.
""".strip()


auth = JWTVerifier(
    public_key=settings.jwt_secret,
    algorithm=settings.jwt_algorithm,
)

mcp = FastMCP(
    name="OWUI-Office-MCP",
    instructions=ROOT_INSTRUCTIONS,
    auth=auth,
)

mcp.add_middleware(RateLimitingMiddleware(
    max_requests_per_second=settings.rate_limit_rps,
    burst_capacity=settings.rate_limit_burst,
    # OpenWebUI JWTs carry the user in the `id` claim.
    get_client_id=lambda context: (
        token.claims.get("id", "anonymous")
        if (token := get_access_token()) else "anonymous"
    ),
))

mcp.mount(pptx_mcp, namespace="pptx")
mcp.mount(docx_mcp, namespace="docx")
mcp.mount(xlsx_mcp, namespace="xlsx")


if __name__ == "__main__":
    mcp.run(
        host="0.0.0.0",
        port=8000,
        transport="http",
    )
