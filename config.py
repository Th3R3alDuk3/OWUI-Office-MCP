from functools import cache
from pathlib import Path

from pydantic_settings import BaseSettings, SettingsConfigDict


class Settings(BaseSettings):
    model_config = SettingsConfigDict(
        env_file=".env",
        env_file_encoding="utf-8",
        extra="ignore",
    )

    jwt_secret: str
    jwt_algorithm: str

    owui_base_url: str
    owui_verify_tls: bool

    rate_limit_rps: float
    rate_limit_burst: int

    templates_dir: Path

    # seconds
    project_ttl: int
    project_sweep_interval: int


@cache
def get_settings() -> Settings:
    return Settings()
