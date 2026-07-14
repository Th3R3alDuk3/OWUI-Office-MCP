import re
from io import BytesIO
from pathlib import Path

from fastmcp.utilities.logging import get_logger
from pptx import Presentation
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml import parse_xml
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.oxml.presentation import CT_SlideId, CT_SlideIdList
from pptx.presentation import Presentation as PresentationType
from pptx.shapes.base import BaseShape
from pptx.slide import Slide, SlideMaster

from models.pptx import LayoutInfo, MasterInfo, PlaceholderInfo, SlideInfo

logger = get_logger(__name__)


_RID_ATTR = (
    "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
)

_BULLET_MARKER = re.compile(
    r"^(?P<indent>[ \t]*)(?:[-‐‑‒–—―−*+•∙‣⁃◦·▪▫■□●○◆◇❖]|\d+[.)])[ \t]+",
    re.MULTILINE,
)

_MAX_BULLET_LEVEL = 8

# Centered pictures and charts fill at most this fraction of the slide.
_PICTURE_SLIDE_FRACTION = 0.8

_CHART_TYPES = {
    "bar": XL_CHART_TYPE.COLUMN_CLUSTERED,
    "line": XL_CHART_TYPE.LINE,
    "pie": XL_CHART_TYPE.PIE,
}


def list_template_names(
    templates_dir: Path,
) -> list[str]:

    if not templates_dir.is_dir():
        raise RuntimeError(f"Templates directory not found: {templates_dir}")

    template_names: list[str] = []

    for file_path in sorted(templates_dir.glob("*.pptx")):
        try:
            _ = Presentation(file_path)
        except Exception as error:
            logger.warning(f"Skipping template {file_path.name}: {error}")
        else:
            template_names.append(file_path.name)

    return template_names


def list_masters(
    presentation: PresentationType,
) -> dict[str, SlideMaster]:

    masters: dict[str, SlideMaster] = {}

    for index, master in enumerate(presentation.slide_masters, start=1):

        # PowerPoint-native files leave masters unnamed and show the theme
        # name instead; names may still repeat across masters.
        name = master.name
    
        if not name:
            theme_part = master.part.part_related_by(RT.THEME)
            name = parse_xml(theme_part.blob).get("name")
    
        name = name or f"Master {index}"
    
        if name in masters:
            name = f"{name} ({index})"

        masters[name] = master

    return masters


def list_master_infos(
    presentation: PresentationType,
) -> dict[str, MasterInfo]:

    master_infos: dict[str, MasterInfo] = {}

    for master_name, master in list_masters(presentation).items():

        layouts: dict[str, LayoutInfo] = {}

        for layout in master.slide_layouts:

            placeholders: list[PlaceholderInfo] = []

            for placeholder in layout.placeholders:

                placeholders.append(
                    PlaceholderInfo(
                        idx=placeholder.placeholder_format.idx,
                        name=placeholder.name,
                        type=str(placeholder.placeholder_format.type),
                    )
                )

            layouts[layout.name] = LayoutInfo(placeholders=placeholders)

        master_infos[master_name] = MasterInfo(layouts=layouts)

    return master_infos


def count_slides(
    presentation: PresentationType,
) -> int:
    return len(presentation.slides)


def list_slide_infos(
    presentation: PresentationType,
) -> list[SlideInfo]:

    slide_infos: list[SlideInfo] = []

    for slide in presentation.slides:

        texts: list[str] = []

        for shape in slide.shapes:
            if shape.has_text_frame and shape.text:
                texts.append(shape.text)

        slide_infos.append(
            SlideInfo(
                layout=slide.slide_layout.name,
                text="\n".join(texts),
            )
        )

    return slide_infos


def _marker_to_indent(match: re.Match[str]) -> str:
    # Placeholders render their own bullets, so drop the marker but keep the
    # nesting as tabs (one tab per level; two spaces count as one tab).
    indent = match.group("indent")
    level = indent.count("\t") + len(indent.replace("\t", "")) // 2
    return "\t" * level


def fill_placeholder(
    placeholder: BaseShape,
    text: str,
) -> None:

    if not placeholder.has_text_frame:
        return

    text_frame = placeholder.text_frame
    text_frame.clear()

    for i, line in enumerate(
        _BULLET_MARKER.sub(_marker_to_indent, text).splitlines()
    ):

        stripped = line.lstrip("\t")

        paragraph = (
            text_frame.paragraphs[0] if i == 0 else text_frame.add_paragraph()
        )
        paragraph.text = stripped
        paragraph.level = min(len(line) - len(stripped), _MAX_BULLET_LEVEL)


def add_centered_picture(
    presentation: PresentationType,
    slide: Slide,
    image: BytesIO,
) -> None:

    picture = slide.shapes.add_picture(image, 0, 0)

    max_width = int(presentation.slide_width * _PICTURE_SLIDE_FRACTION)
    max_height = int(presentation.slide_height * _PICTURE_SLIDE_FRACTION)

    scale = min(max_width / picture.width, max_height / picture.height, 1.0)

    picture.width = int(picture.width * scale)
    picture.height = int(picture.height * scale)
    picture.left = (presentation.slide_width - picture.width) // 2
    picture.top = (presentation.slide_height - picture.height) // 2


def add_chart(
    presentation: PresentationType,
    slide: Slide,
    kind: str,
    categories: list[str],
    series: dict[str, list[float]],
    title: str | None,
    placeholder: BaseShape | None,
) -> None:

    if kind not in _CHART_TYPES:
        raise ValueError(
            f"Chart kind '{kind}' not supported. "
            f"Use one of: {', '.join(_CHART_TYPES)}."
        )

    if not categories or not series:
        raise ValueError("categories and series must not be empty.")

    for name, values in series.items():
        if len(values) != len(categories):
            raise ValueError(
                f"Series '{name}' has {len(values)} values, "
                f"but there are {len(categories)} categories."
            )

    if kind == "pie" and len(series) > 1:
        raise ValueError("A pie chart takes exactly one series.")

    chart_data = CategoryChartData()
    chart_data.categories = categories

    for name, values in series.items():
        chart_data.add_series(name, values)

    if placeholder is not None:
        left, top = placeholder.left, placeholder.top
        width, height = placeholder.width, placeholder.height
        # Drop the now-covered empty placeholder box.
        placeholder._element.getparent().remove(placeholder._element)
    else:
        width = int(presentation.slide_width * _PICTURE_SLIDE_FRACTION)
        height = int(presentation.slide_height * _PICTURE_SLIDE_FRACTION)
        left = (presentation.slide_width - width) // 2
        top = (presentation.slide_height - height) // 2

    graphic_frame = slide.shapes.add_chart(
        _CHART_TYPES[kind], left, top, width, height, chart_data,
    )

    if title is not None:
        chart = graphic_frame.chart
        chart.has_title = True
        chart.chart_title.text_frame.text = title


def add_comment(
    slide: Slide,
    text: str,
) -> None:

    frame = slide.notes_slide.notes_text_frame

    if frame.text:
        frame.add_paragraph().text = text
    else:
        frame.text = text


def _detach_slide(
    presentation: PresentationType,
    slide_id_lst: CT_SlideIdList,
    slide_id: CT_SlideId,
) -> None:

    # A slide is attached twice: package relationship + sldId entry.
    if rid := slide_id.get(_RID_ATTR):
        presentation.part.drop_rel(rid)

    slide_id_lst.remove(slide_id)


def clear_presentation(
    presentation: PresentationType,
) -> None:

    slide_id_lst = presentation.slides._sldIdLst

    for slide_id in list(slide_id_lst):
        _detach_slide(presentation, slide_id_lst, slide_id)


def remove_slides(
    presentation: PresentationType,
    indices: list[int],
) -> None:

    slide_id_lst = presentation.slides._sldIdLst
    slide_ids = list(slide_id_lst)

    targets: list[CT_SlideId] = []

    for index in sorted(set(indices)):
        try:
            targets.append(slide_ids[index])
        except IndexError:
            raise ValueError(
                f"Slide index {index} out of range."
            ) from None

    for slide_id in targets:
        _detach_slide(presentation, slide_id_lst, slide_id)


def move_slide(
    presentation: PresentationType,
    from_index: int,
    to_index: int,
) -> None:

    slide_id_lst = presentation.slides._sldIdLst
    slide_ids = list(slide_id_lst)

    try:
        slide_id = slide_ids[from_index]
    except IndexError:
        raise ValueError(
            f"Slide index {from_index} out of range."
        ) from None

    if not -len(slide_ids) <= to_index < len(slide_ids):
        raise ValueError(f"Target index {to_index} out of range.")

    slide_id_lst.remove(slide_id)
    slide_id_lst.insert(to_index % len(slide_ids), slide_id)
