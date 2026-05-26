from pydantic import BaseModel


class PlaceholderInfo(BaseModel):
    idx: int
    name: str
    type: str


class LayoutInfo(BaseModel):
    placeholders: list[PlaceholderInfo]


class TemplateInfo(BaseModel):
    path: str
    slide_count: int
    layouts: dict[str, LayoutInfo]
