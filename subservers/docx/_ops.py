from io import BytesIO
from pathlib import Path

from docx import Document
from docx.document import Document as DocumentType
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import Cm, Emu
from docx.table import Table
from docx.text.paragraph import Paragraph
from fastmcp.utilities.logging import get_logger
from matplotlib.figure import Figure

from models.docx import BlockInfo, StyleInfo

logger = get_logger(__name__)


_W_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

_COMMENT_AUTHOR = "Assistant"
_COMMENT_INITIALS = "AI"

_CHART_KINDS = ("bar", "line", "pie")

_FIGURE_SIZE = (8.0, 4.5)
_FIGURE_DPI = 150


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


def _block_proxies(
    document: DocumentType,
) -> list:

    # The _Body proxy (not the raw oxml body) so block proxies can reach
    # `.part`, which e.g. the style setter needs.
    body = document._body

    blocks: list = []

    for child in _block_elements(document):

        if child.tag == f"{_W_NS}p":
            blocks.append(Paragraph(child, body))
        elif child.tag == f"{_W_NS}tbl":
            blocks.append(Table(child, body))
        else:
            blocks.append(child)

    return blocks


def count_blocks(
    document: DocumentType,
) -> int:
    return len(_block_elements(document))


def list_block_infos(
    document: DocumentType,
) -> list[BlockInfo]:

    block_infos: list[BlockInfo] = []

    for block in _block_proxies(document):

        if isinstance(block, Paragraph):
            text = block.text

            if not text and block._p.xpath(".//w:drawing"):
                text = "[image]"

            # Its sectPr carries the page setup of all content above it.
            if block._p.xpath("./w:pPr/w:sectPr"):
                text = f"[section break] {text}".rstrip()

            block_infos.append(BlockInfo(type="paragraph", text=text))
        elif isinstance(block, Table):
            block_infos.append(BlockInfo(
                type="table",
                text=f"{len(block.rows)}x{len(block.columns)} table",
            ))
        else:
            block_infos.append(BlockInfo(type=block.tag, text=""))

    return block_infos


def get_paragraph(
    document: DocumentType,
    block_index: int,
) -> Paragraph:

    blocks = _block_proxies(document)

    try:
        block = blocks[block_index]
    except IndexError:
        raise ValueError(
            f"Block index {block_index} out of range."
        ) from None

    if not isinstance(block, Paragraph):
        raise ValueError(f"Block {block_index} is not a paragraph.")

    return block


def insert_picture(
    document: DocumentType,
    image: BytesIO,
    width_cm: float | None,
) -> None:

    if width_cm is not None:
        document.add_picture(image, width=Cm(width_cm))
        return

    picture = document.add_picture(image)
    section = document.sections[-1]

    if section.page_width is None:
        return

    usable_width = Emu(
        section.page_width
        - (section.left_margin or 0)
        - (section.right_margin or 0)
    )

    if picture.width > usable_width:
        picture.height = Emu(
            round(picture.height * usable_width / picture.width)
        )
        picture.width = usable_width


def add_chart(
    document: DocumentType,
    kind: str,
    categories: list[str],
    series: dict[str, list[float]],
    title: str | None,
    width_cm: float | None,
) -> None:

    if kind not in _CHART_KINDS:
        raise ValueError(
            f"Chart kind '{kind}' not supported. "
            f"Use one of: {', '.join(_CHART_KINDS)}."
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

    figure = Figure(figsize=_FIGURE_SIZE, dpi=_FIGURE_DPI, layout="constrained")
    axes = figure.add_subplot()

    if kind == "bar":
        # Center each category's bar group (total width 0.8) on its tick.
        bar_width = 0.8 / len(series)

        for offset, (name, values) in enumerate(series.items()):
            positions = [
                i - 0.4 + bar_width * (offset + 0.5)
                for i in range(len(categories))
            ]
            axes.bar(positions, values, width=bar_width, label=name)

        axes.set_xticks(range(len(categories)), categories)

    elif kind == "line":
        for name, values in series.items():
            axes.plot(categories, values, marker="o", label=name)

    else:
        (values,) = series.values()
        axes.pie(values, labels=categories, autopct="%1.1f%%")

    if kind != "pie":
        axes.grid(axis="y", alpha=0.3)
        axes.legend()

    if title is not None:
        axes.set_title(title)

    buffer = BytesIO()
    figure.savefig(buffer, format="png")
    buffer.seek(0)

    insert_picture(document, buffer, width_cm)


def add_comment(
    document: DocumentType,
    block_index: int,
    text: str,
) -> None:

    paragraph = get_paragraph(document, block_index)

    # Text inside hyperlinks lives outside `paragraph.runs`.
    runs = paragraph.runs or [
        run
        for hyperlink in paragraph.hyperlinks
        for run in hyperlink.runs
    ]

    if not runs:
        raise ValueError(
            f"Block {block_index} is empty; a comment needs text to anchor on."
        )

    document.add_comment(
        runs=runs,
        text=text,
        author=_COMMENT_AUTHOR,
        initials=_COMMENT_INITIALS,
    )


def clear_document(
    document: DocumentType,
) -> None:
    for block in _block_elements(document):
        document.element.body.remove(block)


def remove_blocks(
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
