from fastmcp import FastMCP
from fastmcp.server.auth.providers.jwt import JWTVerifier
from fastmcp.utilities.lifespan import combine_lifespans

from config import get_settings
from subservers.powerpoint.server import lifespan as powerpoint_lifespan
from subservers.powerpoint.server import mcp as powerpoint_mcp


settings = get_settings()

auth = JWTVerifier(
    public_key=settings.jwt_secret,
    algorithm=settings.jwt_algorithm,
)

mcp = FastMCP(
    name="OWUI-Office-MCP",
    auth=auth,
    lifespan=combine_lifespans(powerpoint_lifespan),
)

mcp.mount(powerpoint_mcp, namespace="powerpoint")


if __name__ == "__main__":
    mcp.run(
        host=settings.host, port=settings.port,
        transport="streamable-http",
    )
