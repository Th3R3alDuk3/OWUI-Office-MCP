from functools import cache

from pydantic import Field, PositiveFloat, PositiveInt
from pydantic_settings import BaseSettings, SettingsConfigDict


class Settings(BaseSettings):
    model_config = SettingsConfigDict(
        env_file=".env",
        env_file_encoding="utf-8",
        extra="ignore",
    )

    jwt_secret: str = Field(min_length=1)
    jwt_algorithm: str

    owui_base_url: str
    owui_public_url: str | None = None
    owui_verify_tls: bool

    valkey_url: str

    max_concurrent_requests: PositiveInt
    max_concurrent_requests_per_user: PositiveInt

    rate_limit_rps: PositiveFloat
    rate_limit_burst: PositiveInt

    officecli_max_processes: PositiveInt
    # MB; managed heap of one OfficeCLI run
    officecli_max_memory: PositiveInt
    # seconds; one OfficeCLI run
    officecli_timeout: PositiveFloat
    # seconds; waiting for the project lock and an OfficeCLI process
    officecli_queue_timeout: PositiveFloat

    # seconds; idle time until a project expires
    project_ttl: PositiveInt


@cache
def get_settings() -> Settings:
    return Settings()  # pyright: ignore[reportCallIssue]
