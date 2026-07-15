from collections.abc import Callable
from io import BytesIO

from pptx.shapes.base import BaseShape
from pptx.shapes.placeholder import PicturePlaceholder
from pptx.slide import Slide

from models.pptx import Project
from services.owui import download_file
from subservers.pptx import _ops

SCRIPT_API = """
Run a Python script that builds or edits the current project. The script is
sandboxed: no imports, no file or network access. Available functions:

- set_master(master: str) -> None
  Select the slide master to build with, by name (see `masters` in the
  `create_project` / `open_project` result); required before `add_slide`.
  Pick the master the user asked for; in an opened file, prefer the
  master its slides already use. If several masters could fit and the
  user named none, ask the user which one to use instead of guessing.
- add_slide(layout: str, index: int | None = None) -> int
  Add a slide by layout name from the selected master (see the
  `list_layouts` tool; call it before writing the script); returns the
  slide's zero-based index.
- fill(slide_index: int, placeholder_idx: int, text: str) -> None
  Set placeholder text. Each line becomes a paragraph; prefix lines with
  tabs to nest them as sub-bullets (one tab per level).
- set_notes(slide_index: int, text: str) -> None
  Set the speaker notes of a slide.
- await add_image(slide_index: int, file_id: str,
                  placeholder_idx: int | None = None) -> None
  Insert an image the user attached in OpenWebUI (async: must be awaited).
  With `placeholder_idx` it fills that PICTURE placeholder, else it is
  centered on the slide. `file_id` is never invented.
- add_chart(slide_index: int, kind: str, categories: list[str],
            series: dict[str, list[float]], title: str | None = None,
            placeholder_idx: int | None = None) -> None
  Insert a native, editable chart. `kind` is "bar", "line" or "pie";
  `series` maps each series name to one value per category. A pie chart
  takes exactly one series. With `placeholder_idx` the chart fills that
  placeholder's area, else it is centered on the slide.
- add_comment(slide_index: int, text: str) -> None
  Attach a review comment by appending it to the slide's speaker notes,
  keeping the existing notes. The slide's content stays untouched — use
  it to give feedback on an opened presentation.
- list_slides() -> list[dict]
  Slides in order (layout, text); the list position is the slide index.
- move_slide(from_index: int, to_index: int) -> None
- remove_slides(indices: list[int]) -> None

Plain Python works: loops, conditionals, f-strings, comprehensions,
functions, print. `print(...)` output and the last expression are returned
as `output`. If the script fails, fix it and call `run_script` again.

Example:

    set_master("Office Theme")
    for title, body in [("Q1", "- Umsatz\\n- Kosten"), ("Q2", "- Ausblick")]:
        i = add_slide("Title and Content")
        fill(i, 0, title)
        fill(i, 1, body)

After the requested batch of edits, call `finalize_project` once (not after
each change).
""".strip()


def script_functions(
    project: Project,
    token: str,
) -> dict[str, Callable]:

    presentation = project.presentation

    def _slide(slide_index: int) -> Slide:
        try:
            return presentation.slides[slide_index]
        except IndexError:
            raise ValueError(
                f"Slide index {slide_index} out of range."
            ) from None

    def _placeholder(slide: Slide, placeholder_idx: int) -> BaseShape:
        # `list_layouts` serializes the idx keys as strings; accept both.
        try:
            return slide.placeholders[int(placeholder_idx)]
        except (KeyError, TypeError, ValueError):
            raise ValueError(
                f"Placeholder idx {placeholder_idx} not found on this "
                "slide's layout."
            ) from None

    def set_master(master: str) -> None:

        if master not in _ops.list_masters(presentation):
            raise ValueError(
                f"Master '{master}' not found. Use a name from `masters`."
            )

        project.master_name = master

    def add_slide(layout: str, index: int | None = None) -> int:

        if project.master_name is None:
            raise ValueError(
                "No master selected. Call set_master(...) first, "
                "using a name from `masters`."
            )

        master = _ops.list_masters(presentation)[project.master_name]

        layout_obj = master.slide_layouts.get_by_name(layout)
        if layout_obj is None:
            raise ValueError(
                f"Layout '{layout}' not found in master "
                f"'{project.master_name}'. Use a name from `list_layouts`."
            )

        presentation.slides.add_slide(layout_obj)

        if index is not None:
            _ops.move_slide(
                presentation, _ops.count_slides(presentation) - 1, index,
            )
            return index % _ops.count_slides(presentation)

        return _ops.count_slides(presentation) - 1

    def fill(slide_index: int, placeholder_idx: int, text: str) -> None:
        placeholder = _placeholder(_slide(slide_index), placeholder_idx)
        _ops.fill_placeholder(placeholder, text)

    def set_notes(slide_index: int, text: str) -> None:
        slide = _slide(slide_index)
        slide.notes_slide.notes_text_frame.text = text

    async def add_image(
        slide_index: int,
        file_id: str,
        placeholder_idx: int | None = None,
    ) -> None:

        slide = _slide(slide_index)

        if placeholder_idx is not None:
            placeholder = _placeholder(slide, placeholder_idx)

            if not isinstance(placeholder, PicturePlaceholder):
                raise ValueError(
                    f"Placeholder idx {placeholder_idx} is not a "
                    "PICTURE placeholder."
                )

        file_content = await download_file(file_id=file_id, token=token)

        try:
            if placeholder_idx is None:
                _ops.add_centered_picture(
                    presentation, slide, BytesIO(file_content),
                )
            else:
                placeholder.insert_picture(BytesIO(file_content))
        except OSError as error:
            raise ValueError(
                f"File '{file_id}' is not a supported image."
            ) from error

    def add_chart(
        slide_index: int,
        kind: str,
        categories: list[str],
        series: dict[str, list[float]],
        title: str | None = None,
        placeholder_idx: int | None = None,
    ) -> None:

        slide = _slide(slide_index)

        placeholder = (
            _placeholder(slide, placeholder_idx)
            if placeholder_idx is not None else None
        )

        _ops.add_chart(
            presentation, slide, kind, categories, series, title, placeholder,
        )

    def add_comment(slide_index: int, text: str) -> None:
        _ops.add_comment(_slide(slide_index), text)

    def list_slides() -> list[dict]:
        return [
            info.model_dump()
            for info in _ops.list_slide_infos(presentation)
        ]

    def move_slide(from_index: int, to_index: int) -> None:
        _ops.move_slide(presentation, from_index, to_index)

    def remove_slides(indices: list[int]) -> None:
        _ops.remove_slides(presentation, indices)

    return {
        "set_master": set_master,
        "add_slide": add_slide,
        "fill": fill,
        "set_notes": set_notes,
        "add_image": add_image,
        "add_chart": add_chart,
        "add_comment": add_comment,
        "list_slides": list_slides,
        "move_slide": move_slide,
        "remove_slides": remove_slides,
    }
