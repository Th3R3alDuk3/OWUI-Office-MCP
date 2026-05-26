import logging
from pathlib import Path

from pptx import Presentation
from pptx.presentation import Presentation as PresentationType

from models.template import LayoutInfo, PlaceholderInfo, TemplateInfo

logger = logging.getLogger(__name__)


def analyze_templates(templates_dir: Path) -> dict[str, TemplateInfo]:
    if not templates_dir.is_dir():
        raise RuntimeError(f"Templates directory not found: {templates_dir}")

    templates: dict[str, TemplateInfo] = {}
    for path in sorted(templates_dir.glob("*.pptx")):
        try:
            prs = Presentation(str(path))
        except Exception as exc:
            logger.warning("Skipping template %s: %s", path.name, exc)
            continue

        layouts = {
            layout.name: LayoutInfo(
                placeholders=[
                    PlaceholderInfo(
                        idx=ph.placeholder_format.idx,
                        name=ph.name,
                        type=str(ph.placeholder_format.type),
                    )
                    for ph in layout.placeholders
                ],
            )
            for layout in prs.slide_layouts
        }

        templates[path.stem] = TemplateInfo(
            path=str(path.resolve()),
            slide_count=len(prs.slides),
            layouts=layouts,
        )

    return templates


_RID_ATTR = (
    "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
)


def drop_slide(pptx: PresentationType, index: int) -> None:
    sld_id_lst = pptx.slides._sldIdLst
    sld_ids = list(sld_id_lst)
    sld_id = sld_ids[index]
    rid = sld_id.get(_RID_ATTR)
    if rid:
        pptx.part.drop_rel(rid)
    sld_id_lst.remove(sld_id)


def drop_all_slides(pptx: PresentationType) -> None:
    while len(pptx.slides._sldIdLst) > 0:
        drop_slide(pptx, 0)
