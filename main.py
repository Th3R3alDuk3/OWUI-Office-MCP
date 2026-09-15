from fastmcp import FastMCP
from fastmcp.server.auth.providers.jwt import JWTVerifier
from fastmcp.server.dependencies import get_access_token
from fastmcp.server.middleware.rate_limiting import RateLimitingMiddleware

from config import get_settings
from tools import TOOLS, office_lifespan

settings = get_settings()

INSTRUCTIONS = """
OWUI-Office-MCP creates and edits PowerPoint, Word and Excel files with
OfficeCLI commands, faithful to a design: a stored template or a file the
user attached.

Workflow: `list_templates` unless the user attached a design, then
`create_project` or `open_project`, then `run_commands` with one atomic
batch per edit step, `preview_project` to check the result, and
`export_project` once the user's request is fully applied. Every result
carries a `hint` with the next step; follow it unless the user's request
says otherwise.
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
