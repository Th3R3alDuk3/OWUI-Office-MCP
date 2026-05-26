from pydantic import BaseModel, ConfigDict


class OWUIFileMeta(BaseModel):
    model_config = ConfigDict(extra="allow")

    name: str | None = None
    content_type: str | None = None
    size: int | None = None


class OWUIFile(BaseModel):
    model_config = ConfigDict(extra="allow")

    id: str
    user_id: str | None = None
    filename: str | None = None
    hash: str | None = None
    meta: OWUIFileMeta | None = None
    created_at: int | None = None
    updated_at: int | None = None
