from pydantic import BaseModel


class ProjectInfo(BaseModel):
    user_id: str
    template_name: str
    slide_count: int


class SlideInfo(BaseModel):
    index: int
    layout_name: str


class SavedProjectInfo(BaseModel):
    filename: str
    slide_count: int
    owui_file_id: str
    owui_url: str
