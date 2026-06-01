from pathlib import Path

from fastmcp.utilities.logging import get_logger
from pptx import Presentation
from pptx.opc.constants import RELATIONSHIP_TYPE
from pptx.oxml import parse_xml
from pptx.oxml.presentation import CT_SlideId, CT_SlideIdList
from pptx.presentation import Presentation as PresentationType

from models.pptx import LayoutInfo, PlaceholderInfo, SlideInfo


logger = get_logger(__name__)


_RID_ATTR = (
    "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
)


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


def list_master_names(
    presentation: PresentationType,
) -> dict[int, str]:

    master_names: dict[int, str] = {}

    for i, master in enumerate(presentation.slide_masters):

        master_name = master.name

        if not master_name:
            theme_part = master.part.part_related_by(RELATIONSHIP_TYPE.THEME)
            theme_xml = parse_xml(theme_part.blob)
            master_name = theme_xml.get("name")

        if master_name:
            master_names[i] = master_name

    return master_names


def list_layout_infos(
    presentation: PresentationType,
    master_index: int,
) -> dict[str, LayoutInfo]:

    master = presentation.slide_masters[master_index]

    layout_infos: dict[str, LayoutInfo] = {}

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

        layout_infos[layout.name] = LayoutInfo(placeholders=placeholders)

    return layout_infos


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


def _detach_slide(
    presentation: PresentationType,
    slide_id_lst: CT_SlideIdList,
    slide_id: CT_SlideId,
) -> None:

    if rid := slide_id.get(_RID_ATTR):
        presentation.part.drop_rel(rid)

    slide_id_lst.remove(slide_id)


def drop_slides(
    presentation: PresentationType,
    indices: list[int],
) -> None:

    slide_id_lst = presentation.slides._sldIdLst
    slide_ids = list(slide_id_lst)

    targets = []

    for index in sorted(set(indices)):
        try:
            targets.append(slide_ids[index])
        except IndexError:
            raise ValueError(
                f"Slide index {index} out of range."
            ) from None

    for slide_id in targets:
        _detach_slide(presentation, slide_id_lst, slide_id)


def drop_all_slides(
    presentation: PresentationType,
) -> None:

    slide_id_lst = presentation.slides._sldIdLst

    for slide_id in list(slide_id_lst):
        _detach_slide(presentation, slide_id_lst, slide_id)


def move_slide(
    presentation: PresentationType,
    from_index: int,
    to_index: int,
) -> None:

    slide_id_lst = presentation.slides._sldIdLst
    slide_ids = list(slide_id_lst)

    try:
        sld_id = slide_ids[from_index]
    except IndexError:
        raise ValueError(
            f"Slide index {from_index} out of range."
        ) from None

    if not -len(slide_ids) <= to_index < len(slide_ids):
        raise ValueError(f"Target index {to_index} out of range.")

    slide_id_lst.remove(sld_id)
    slide_id_lst.insert(to_index % len(slide_ids), sld_id)
