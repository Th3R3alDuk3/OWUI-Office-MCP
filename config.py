from functools import lru_cache
from pathlib import Path

from pydantic_settings import BaseSettings, SettingsConfigDict


class Settings(BaseSettings):
    model_config = SettingsConfigDict(
        env_file=".env",
        env_file_encoding="utf-8",
        extra="ignore",
    )

    host: str
    port: int

    templates_dir: Path

    jwt_secret: str
    jwt_algorithm: str

    owui_base_url: str

    project_ttl_seconds: int
    project_sweep_interval_seconds: int


@lru_cache(maxsize=1)
def get_settings() -> Settings:
    return Settings()
