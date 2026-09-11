from typing import Annotated, Literal

from pydantic import BaseModel, ConfigDict, Field, StringConstraints, model_validator

# OfficeCLI treats element types and prop names case-insensitively.
Lowercase = Annotated[str, StringConstraints(to_lower=True)]


class Command(BaseModel):
    model_config = ConfigDict(extra="forbid")

    command: Literal[
        "add", "set", "remove", "move", "swap", "import",
        "get", "query", "view", "validate",
    ]
    path: str | None = Field(
        default=None,
        description="Target of set, remove, move, get.",
    )
    parent: str | None = Field(
        default=None,
        description="Parent that add or import writes into.",
    )
    type: Lowercase | None = Field(
        default=None,
        description="Element type that add creates.",
    )
    from_: str | None = Field(
        default=None,
        alias="from",
        description="Existing element that add copies.",
    )
    index: int | None = Field(
        default=None,
        ge=0,
        description="0-based position for add and move.",
    )
    after: str | None = Field(
        default=None,
        description="Add or move after this element.",
    )
    before: str | None = Field(
        default=None,
        description="Add or move before this element.",
    )
    to: str | None = Field(
        default=None,
        description="Target parent of a move.",
    )
    path2: str | None = Field(
        default=None,
        description="Second element of a swap.",
    )
    props: dict[Lowercase, str | int | float | bool] | None = Field(
        default=None,
        description="Property name -> value.",
    )
    selector: str | None = Field(
        default=None,
        description="Filter of a query, e.g. `slide`.",
    )
    text: str | None = Field(
        default=None,
        description="Inline CSV of an import.",
    )
    mode: Literal["text", "annotated", "outline", "stats", "issues"] | None = Field(
        default=None,
        description="Required by view.",
    )
    depth: int | None = Field(
        default=None,
        ge=0,
        le=4,
        description="Child levels that get returns.",
    )

    @model_validator(mode="after")
    def validate_mode(self) -> Command:

        if self.command == "view" and self.mode is None:
            raise ValueError("`view` needs a `mode`.")

        return self
