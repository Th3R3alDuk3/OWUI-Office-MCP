from functools import lru_cache
from pathlib import Path

from pydantic_settings import BaseSettings, SettingsConfigDict


class Settings(BaseSettings):
    model_config = SettingsConfigDict(
        env_file=".env",
        env_file_encoding="utf-8",
        extra="ignore",
    )

    host: str = "127.0.0.1"
    port: int = 8000

    templates_dir: Path

    jwt_secret: str
    jwt_algorithm: str = "HS256"

    owui_base_url: str

    project_ttl_seconds: int = 3600
    project_sweep_interval_seconds: int = 300


@lru_cache(maxsize=1)
def get_settings() -> Settings:
    return Settings()
