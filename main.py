from fastmcp import FastMCP
from fastmcp.server.auth.providers.jwt import JWTVerifier
from fastmcp.utilities.lifespan import combine_lifespans

from config import get_settings
from subservers.docx.server import lifespan as docx_lifespan
from subservers.docx.server import mcp as docx_mcp
from subservers.pptx.server import lifespan as pptx_lifespan
from subservers.pptx.server import mcp as pptx_mcp


settings = get_settings()

ROOT_INSTRUCTIONS = """
OWUI-Office-MCP exposes two stateful Office toolsets:
- Use the `pptx_*` tools for PowerPoint presentations.
- Use the `docx_*` tools for Word documents.

Each user has one in-memory project per toolset. `create_project` starts a new
project from a template and overwrites any existing project for that user in
the same toolset.

Workflow: create a project, build or edit it with insert/edit/move/remove
tools, then call the matching `download_project` tool exactly once after the
user's requested batch of edits is complete. Do not call `download_project`
after every individual change when multiple changes belong to one request. If
the user later asks for another edit, apply that edit batch to the active
project and call `download_project` once again.

Mounted tool names are prefixed, for example `pptx_download_project` and
`docx_download_project`. Tool descriptions may mention local names like
`download_project`; use the corresponding prefixed tool exposed by this root
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
    lifespan=combine_lifespans(pptx_lifespan, docx_lifespan),
)

mcp.mount(pptx_mcp, namespace="pptx")
mcp.mount(docx_mcp, namespace="docx")


if __name__ == "__main__":
    mcp.run(
        host=settings.host, port=settings.port,
        transport="streamable-http",
    )
