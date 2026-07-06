from asyncio import CancelledError, create_task, sleep
from collections.abc import AsyncIterator
from contextlib import asynccontextmanager, suppress

from cachetools import TTLCache
from fastmcp import FastMCP


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

    def get(
        self,
        user_id: str,
    ) -> ProjectT | None:
        return self._projects.get(user_id)

    def touch(
        self,
        user_id: str,
        project: ProjectT,
    ) -> None:
        # Re-inserting resets the sliding TTL, but only for the live project.
        if self._projects.get(user_id) is project:
            self._projects[user_id] = project

    async def _sweep_task(self) -> None:
        while True:
            await sleep(self._sweep_interval)
            self._projects.expire()

    @asynccontextmanager
    async def lifespan(
        self,
        server: FastMCP,
    ) -> AsyncIterator[None]:

        task = create_task(self._sweep_task())

        try:
            yield
        finally:

            task.cancel()

            with suppress(CancelledError):
                await task
