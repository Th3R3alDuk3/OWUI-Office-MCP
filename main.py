from fastmcp import FastMCP
from fastmcp.server.auth.providers.jwt import JWTVerifier
from fastmcp.utilities.lifespan import combine_lifespans

from config import get_settings
from subservers.docx.server import lifespan as docx_lifespan
from subservers.docx.server import mcp as docx_mcp
from subservers.pptx.server import lifespan as pptx_lifespan
from subservers.pptx.server import mcp as pptx_mcp


settings = get_settings()

auth = JWTVerifier(
    public_key=settings.jwt_secret,
    algorithm=settings.jwt_algorithm,
)

mcp = FastMCP(
    name="OWUI-Office-MCP",
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
