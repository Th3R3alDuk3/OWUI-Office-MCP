from asyncio import to_thread
from pathlib import Path
from posixpath import basename, dirname, join, normpath
from urllib.parse import unquote
from xml.etree import ElementTree
from xml.etree.ElementTree import Element
from zipfile import ZipFile

from fastmcp.exceptions import ToolError

from models.office import Layout, Master, PptxInventory
from services.officecli import run_batch

FORMAT = "pptx"
MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
INVENTORY = PptxInventory


START_HINT = """
Bound to the master in `inventory.masters`: use its layouts and slots only.
Build in steps: a few slides per batch, then `view outline` or
`preview_project`, then continue. Engine rules: a slide takes `layout` = the
layout's `index` (or name); fill a slot with `add` type `placeholder`, `phType`
and, when the slot names one, `idx` (`slots` are `phType:idx`, use them
exactly as listed); change an existing placeholder with `set` on its path
from `query placeholder`. `\n`
in `text` starts a new paragraph. Charts take `chartType`, `categories`
"Q1,Q2" and `data` "Revenue:1,2;Costs:3,4"; tables `data` "H1,H2;r1c1,r1c2";
images come from attached files as `image` (placeholder) or `src` (picture)
"file:<file_id>"; free elements are placed by `x`, `y`, `width`, `height`
within `inventory.slide_size`. Paths are 1-based, `index` in add/move is
0-based. Stay within the design: change colors, fonts and other appearance
only when the user asks. `get_reference` documents elements and props.
""".strip()

# Slots the master fills itself.
_FIXED_SLOTS = {"date", "dt", "footer", "ftr", "header", "hdr", "slidenum", "sldnum"}
_NAMESPACE = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}
_RELATIONSHIP_ID = f"{{{_NAMESPACE['r']}}}id"


async def prepare(
    file: Path,
) -> None:

    (slides,) = await run_batch(file, [
        {"command": "query", "selector": "slide"},
    ])

    if paths := [slide["path"] for slide in slides["output"]["results"]]:
        await run_batch(file, [
            {"command": "remove", "path": path} for path in reversed(paths)
        ])


def _relationships(
    archive: ZipFile,
    part: str,
) -> dict[str, tuple[str, str]]:

    path = join(dirname(part), "_rels", f"{basename(part)}.rels")
    return {
        link.attrib["Id"]: (
            link.attrib["Type"].rsplit("/", 1)[-1],
            normpath(join(dirname(part), unquote(link.attrib["Target"]))).lstrip("/"),
        )
        for link in ElementTree.fromstring(archive.read(path))
        if link.get("TargetMode") != "External"
    }


def _presentation(
    archive: ZipFile,
) -> str:

    return next(
        target for kind, target in _relationships(archive, "").values()
        if kind == "officeDocument"
    )


def _children(
    archive: ZipFile,
    part: str,
    kind: str,
    id_list: str,
) -> tuple[Element, list[str]]:

    root = ElementTree.fromstring(archive.read(part))
    links = _relationships(archive, part)
    declared = [
        links[node.attrib[_RELATIONSHIP_ID]][1]
        for node in root.findall(id_list, _NAMESPACE)
    ]
    linked = {target for link_kind, target in links.values() if link_kind == kind}

    # Match OfficeCLI's write order: declared IDs, then sorted orphan parts.
    return root, list(dict.fromkeys([*declared, *sorted(linked - set(declared))]))


def _name(
    element: Element,
) -> str:

    common = element.find("p:cSld", _NAMESPACE)
    return common.get("name", "") if common is not None else ""


def _design(
    file: Path,
) -> tuple[list[Master], dict[str, int], set[str]]:

    masters: list[Master] = []
    layout_ids: dict[str, int] = {}
    layout_names: set[str] = set()

    with ZipFile(file) as archive:

        _, master_parts = _children(
            archive, _presentation(archive), "slideMaster",
            "p:sldMasterIdLst/p:sldMasterId",
        )

        for index, master_part in enumerate(master_parts, start=1):

            master, layout_parts = _children(
                archive, master_part, "slideLayout", "p:sldLayoutIdLst/p:sldLayoutId",
            )
            layouts: list[Layout] = []

            for part in layout_parts:
                layout_ids[part] = len(layout_ids) + 1
                layout = ElementTree.fromstring(archive.read(part))
                name = _name(layout)
                layout_names.update((name, layout.get("matchingName", "")))
                slots = []
                for slot in layout.findall(".//p:ph", _NAMESPACE):
                    slot_type = slot.get("type", "obj")
                    if slot_type.lower() not in _FIXED_SLOTS:
                        slots.append(
                            f"{slot_type}:{slot.attrib['idx']}"
                            if "idx" in slot.attrib else slot_type
                        )
                layouts.append(Layout(
                    index=layout_ids[part], name=name, slots=slots,
                ))

            masters.append(Master(
                index=index, name=_name(master) or f"Master {index}", layouts=layouts,
            ))

    return masters, layout_ids, layout_names


async def inventory(
    file: Path,
) -> PptxInventory:

    (root,) = await run_batch(file, [
        {"command": "get", "path": "/", "depth": 0},
    ])
    properties = root["output"]["results"][0]["format"]

    return PptxInventory(
        slide_size=f"{properties['slideWidth']} x {properties['slideHeight']}",
        masters=(await to_thread(_design, file))[0],
    )


def bind(
    inventory: PptxInventory,
    master: int | None,
) -> PptxInventory:

    names = ", ".join(
        f"{candidate.index} '{candidate.name}'" for candidate in inventory.masters
    )

    if master is None and len(inventory.masters) > 1:
        raise ToolError(
            f"The design has several masters: {names}. Pass `master` with the "
            "one the user named; if they named none, ask the user."
        )

    masters = [
        candidate for candidate in inventory.masters
        if master in (None, candidate.index)
    ]

    if not masters:
        raise ToolError(f"Master {master} not found. Use one of: {names}.")

    return inventory.model_copy(update={"masters": masters})


def disambiguate_layouts(
    file: Path,
    batch: list[dict],
) -> None:

    _, _, names = _design(file)

    # The engine tries names before integers; whitespace makes a number a number.
    for command in batch:
        props = command.get("props", {})
        hint = str(props.get("layout", ""))
        if hint.isdigit():
            while hint in names:
                hint = f" {hint}"
            props["layout"] = hint


def check(
    file: Path,
    inventory: PptxInventory,
) -> None:

    _, layout_ids, _ = _design(file)
    allowed = {
        layout.index for master in inventory.masters for layout in master.layouts
    }

    with ZipFile(file) as archive:

        presentation = _presentation(archive)
        root = ElementTree.fromstring(archive.read(presentation))
        links = _relationships(archive, presentation)

        for index, slide in enumerate(
            root.findall("p:sldIdLst/p:sldId", _NAMESPACE), start=1,
        ):
            part = links[slide.attrib[_RELATIONSHIP_ID]][1]
            layout = next((
                target for kind, target in _relationships(archive, part).values()
                if kind == "slideLayout"
            ), None)
            if layout is None or layout_ids.get(layout) not in allowed:
                raise ToolError(
                    f"Slide {index} is outside the selected master. Use one of "
                    "its layouts; no document changes were published."
                )

            # Compared lower-cased like the engine does; shown as authored.
            slots = {
                (slot.get("type", "obj").lower(), int(slot.get("idx", "0"))):
                f"{slot.get('type', 'obj')}:{slot.get('idx')}" if "idx" in slot.attrib
                else slot.get("type", "obj")
                for slot in ElementTree.fromstring(archive.read(layout)).findall(
                    ".//p:ph", _NAMESPACE,
                )
            }
            for slot in ElementTree.fromstring(archive.read(part)).findall(
                ".//p:ph", _NAMESPACE,
            ):
                slot_type = slot.get("type", "obj").lower()
                slot_index = int(slot.get("idx", "0"))
                if (
                    slot_type not in _FIXED_SLOTS
                    and (slot_type, slot_index) not in slots
                ):
                    raise ToolError(
                        f"Slide {index}: placeholder {slot_type}:{slot_index} "
                        "does not belong to its layout, whose slots are "
                        + ", ".join(
                            label for (kind, _), label in sorted(slots.items())
                            if kind not in _FIXED_SLOTS
                        )
                        + ". Use `phType` and `idx` from `slots`; no document "
                        "changes were published."
                    )
