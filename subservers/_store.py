from asyncio import CancelledError, create_task, sleep
from collections.abc import AsyncIterator
from contextlib import asynccontextmanager, suppress

from cachetools import TTLCache
from fastmcp import FastMCP
from fastmcp.dependencies import TokenClaim
from fastmcp.exceptions import ToolError


class ProjectStore[ProjectT]:

    def __init__(
        self,
        max_size: int,
        ttl: float,
        sweep_interval: float,
    ) -> None:
        self._projects: TTLCache[str, ProjectT] = TTLCache(max_size, ttl)
        self._sweep_interval = sweep_interval

    def set(
        self,
        user_id: str,
        project: ProjectT,
    ) -> None:
        self._projects[user_id] = project

    def require(
        self,
        user_id: str = TokenClaim("id"),
    ) -> ProjectT:

        project = self._projects.get(user_id)

        if project is None:
            # ToolError passes dependency resolution unchanged; other
            # exceptions get swallowed into a generic "failed to resolve"
            # message.
            raise ToolError(
                "No project for this user. "
                "Call `create_project` or `open_project` first."
            )

        return project

    def touch(
        self,
        user_id: str,
        project: ProjectT,
    ) -> None:
        # Re-inserting resets the sliding TTL, but only for the live project.
        if self._projects.get(user_id) is project:
            self._projects[user_id] = project

    async def _sweep_loop(self) -> None:
        while True:
            await sleep(self._sweep_interval)
            self._projects.expire()

    @asynccontextmanager
    async def lifespan(
        self,
        server: FastMCP,
    ) -> AsyncIterator[None]:

        task = create_task(self._sweep_loop())

        try:
            yield
        finally:

            task.cancel()

            with suppress(CancelledError):
                await task
