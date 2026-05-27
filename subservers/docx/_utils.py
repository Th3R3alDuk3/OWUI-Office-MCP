from fastmcp.utilities.logging import get_logger
from pathlib import Path

from docx import Document
from docx.document import Document as DocumentType
from docx.enum.style import WD_STYLE_TYPE

from models.docx import StyleInfo


logger = get_logger(__name__)


_W_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
_SECT_PR = f"{_W_NS}sectPr"

_RELEVANT_STYLE_TYPES = {
    WD_STYLE_TYPE.PARAGRAPH,
    WD_STYLE_TYPE.TABLE,
}


def list_template_names(
    templates_dir: Path,
) -> list[str]:

    if not templates_dir.is_dir():
        raise RuntimeError(f"Templates directory not found: {templates_dir}")

    template_names: list[str] = []

    for file_path in sorted(templates_dir.glob("*.docx")):
        try:
            _ = Document(file_path)
        except Exception as error:
            logger.warning(f"Skipping template {file_path.name}: {error}")
        else:
            template_names.append(file_path.name)

    return template_names


def list_style_infos(
    templates_dir: Path,
    template_name: str,
) -> dict[str, StyleInfo]:

    document = Document(templates_dir.joinpath(template_name))

    style_infos: dict[str, StyleInfo] = {}

    for style in document.styles:

        if style.type not in _RELEVANT_STYLE_TYPES:
            continue

        style_infos[style.name] = StyleInfo(
            type=str(style.type),
            builtin=bool(style.builtin),
        )

    return style_infos


def _content_blocks(document: DocumentType) -> list:
    return [
        child for child in document.element.body
        if child.tag != _SECT_PR
    ]


def count_blocks(document: DocumentType) -> int:
    return len(_content_blocks(document))


def drop_block(document: DocumentType, index: int) -> None:

    blocks = _content_blocks(document)

    try:
        block = blocks[index]
    except IndexError:
        raise ValueError(
            f"Block index {index} out of range."
        )

    document.element.body.remove(block)


def drop_all_blocks(document: DocumentType) -> None:
    for block in _content_blocks(document):
        document.element.body.remove(block)
