from pydantic import BaseModel, Field


class ToolResult(BaseModel):
    hint: str = Field(
        description=(
            "Suggested next step — guidance for the agent, not part of the data."
        ),
    )


class TemplatesResult(ToolResult):
    templates: list[str] = Field(
        description="Template file names for `create_project`.",
    )


class UploadResult(ToolResult):
    file_name: str = Field(
        description="Name of the uploaded file.",
    )
    owui_url: str = Field(
        description="OpenWebUI download URL of the uploaded file.",
    )
