from pathlib import Path

from docx import Document
from docx.document import Document as DocumentType
from docx.enum.style import WD_STYLE_TYPE
from docx.table import Table
from docx.text.paragraph import Paragraph
from fastmcp.utilities.logging import get_logger

from models.docx import BlockInfo, StyleInfo

logger = get_logger(__name__)


_W_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


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
    document: DocumentType,
) -> dict[str, StyleInfo]:

    style_infos: dict[str, StyleInfo] = {}

    for style in document.styles:

        if style.type not in {WD_STYLE_TYPE.PARAGRAPH, WD_STYLE_TYPE.TABLE}:
            continue

        style_infos[style.name] = StyleInfo(
            type=str(style.type),
            builtin=bool(style.builtin),
        )

    return style_infos


def _block_elements(
    document: DocumentType,
) -> list:

    children: list = []

    for child in document.element.body:
        if child.tag != f"{_W_NS}sectPr":
            children.append(child)

    return children


def count_blocks(
    document: DocumentType,
) -> int:
    return len(_block_elements(document))


def content_blocks(
    document: DocumentType,
) -> list:

    body = document.element.body

    blocks: list = []

    for child in _block_elements(document):

        if child.tag == f"{_W_NS}p":
            blocks.append(Paragraph(child, body))
        elif child.tag == f"{_W_NS}tbl":
            blocks.append(Table(child, body))
        else:
            blocks.append(child)

    return blocks


def list_block_infos(
    document: DocumentType,
) -> list[BlockInfo]:

    block_infos: list[BlockInfo] = []

    for block in content_blocks(document):

        if isinstance(block, Paragraph):
            block_infos.append(BlockInfo(type="paragraph", text=block.text))
        elif isinstance(block, Table):
            block_infos.append(BlockInfo(
                type="table",
                text=f"{len(block.rows)}x{len(block.columns)} table",
            ))
        else:
            block_infos.append(BlockInfo(type=block.tag, text=""))

    return block_infos


def drop_blocks(
    document: DocumentType,
    indices: list[int],
) -> None:

    blocks = _block_elements(document)
    body = document.element.body

    targets: list = []

    for index in sorted(set(indices)):
        try:
            targets.append(blocks[index])
        except IndexError:
            raise ValueError(
                f"Block index {index} out of range."
            ) from None

    for block in targets:
        body.remove(block)


def drop_all_blocks(
    document: DocumentType,
) -> None:
    for block in _block_elements(document):
        document.element.body.remove(block)


def move_block(
    document: DocumentType,
    from_index: int,
    to_index: int,
) -> None:

    body = document.element.body
    blocks = _block_elements(document)

    try:
        block = blocks[from_index]
    except IndexError:
        raise ValueError(
            f"Block index {from_index} out of range."
        ) from None

    if not -len(blocks) <= to_index < len(blocks):
        raise ValueError(f"Target index {to_index} out of range.")

    body.remove(block)
    body.insert(to_index % len(blocks), block)
