import re
from pathlib import Path
from xml.etree import ElementTree

from fastmcp.exceptions import ToolError

from models.office import Layout, PptxInventory, SlideMaster
from services.officecli import run_batch

FORMAT = "pptx"
MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
INVENTORY = PptxInventory


START_HINT = (
    "Bound to the slide master in `inventory.slide_masters`: use its layouts "
    "and slots only. Check the result with `view outline` or "
    "`preview_project`, text overflow with `view issues`. Add a slide with "
    "`add` type `slide`, `parent` `/` and the prop `layout` = the layout's "
    "`index`; the props `title` and `text` fill its title and first body "
    "slot. Other placeholders, speaker notes and comments take the slide's "
    "path, e.g. `/slide[1]`, as `parent`: fill a slot with `add` type "
    "`placeholder` and the props `phType` and `idx` from `slots`, notes with "
    "type `notes` and the prop `text`, comments with type `comment` and the "
    "props `text` and `author`. Change an existing placeholder with `set` on "
    "its path from `query placeholder`. A line break in `text` starts a new "
    "paragraph. Charts take `chartType` (`column` is vertical, `bar` "
    "horizontal), `categories` \"Q1,Q2\" and `data` \"Revenue:1,2;Costs:3,4\"; "
    "tables `data` \"H1,H2;r1c1,r1c2\"; images come from attached files as "
    "`image` (placeholder) or `src` (picture) \"file:<file_id>\"; free elements "
    "need `x`, `y`, `width` and `height` with units, e.g. `100pt`, within "
    "`inventory.slide_size` or at a slot's position from `slots`: OfficeCLI's "
    "defaults ignore the slide size. Paths are 1-based, `index` in add/move "
    "is 0-based. Change colors and fonts only when the user asks or text "
    "would be unreadable, then with theme colors such as `light1` or "
    "`accent1`. `get_reference` documents elements and props."
)

# Slots the master fills itself.
_FIXED_SLOTS = {"date", "footer", "header", "slidenum"}
_POSITION = ("x", "y", "width", "height")
# Appended to every edit batch, so the check needs no run of its own.
DECK = {"command": "get", "path": "/", "depth": 2}

_PRESENTATION = "/ppt/presentation.xml"
_PRESENTATION_RELS = "/ppt/_rels/presentation.xml.rels"
_NAMESPACES = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}
_RELATIONSHIP_ID = f"{{{_NAMESPACES['r']}}}id"
_MASTER_IDS = "p:sldMasterIdLst/p:sldMasterId"

_SLIDE_PATH = re.compile(r"/slide\[\d+\]", re.IGNORECASE)


async def prepare(
    file: Path,
) -> None:

    slides, presentation = await run_batch(file, [
        {"command": "query", "selector": "slide"},
        {"command": "raw", "part": _PRESENTATION},
    ])
    commands = [
        # Highest first, so the remaining paths stay valid.
        {"command": "remove", "path": slide["path"]}
        for slide in reversed(slides["output"]["results"])
    ]

    # Sections would keep pointing at the removed slides.
    if "sectionLst" in presentation["output"]:
        commands.append({
            "command": "raw-set", "part": _PRESENTATION, "action": "remove",
            "xpath": "//p:extLst/p:ext[*[local-name()='sectionLst']]",
        })

    if commands:
        await run_batch(file, commands)


def _slots(
    shapes: list[dict],
) -> dict[str, dict[str, str]]:

    slots: dict[str, dict[str, str]] = {}

    for shape in shapes:
        properties = shape["format"]
        kind = properties.get("phType", "")
        if kind and kind not in _FIXED_SLOTS:
            slots[f"{kind}:{properties.get('phIndex', 0)}"] = {
                key: properties[key] for key in _POSITION if key in properties
            }

    return slots


async def inventory(
    file: Path,
) -> PptxInventory:

    root, masters, presentation, relationships = await run_batch(file, [
        {"command": "get", "path": "/", "depth": 0},
        {"command": "query", "selector": "slidemaster"},
        {"command": "raw", "part": _PRESENTATION},
        {"command": "raw", "part": _PRESENTATION_RELS},
    ])
    # OfficeCLI numbers `/slidemaster[M]` in relationship order, but counts
    # `layout` through the masters in PowerPoint's order (sldMasterIdLst).
    masters_by_id = dict(zip(
        (
            link.attrib["Id"]
            for link in ElementTree.fromstring(relationships["output"])
            if link.attrib["Type"].endswith("/slideMaster")
        ),
        masters["output"]["results"],
        strict=True,
    ))
    master_ids = [
        node.attrib[_RELATIONSHIP_ID]
        for node in ElementTree.fromstring(presentation["output"])
        .iterfind(_MASTER_IDS, _NAMESPACES)
    ]
    listed: list[SlideMaster] = []
    layout_index = 0

    for master_index, master_id in enumerate(master_ids, start=1):

        master = masters_by_id[master_id]
        # Within a master, paths and `layout` count alike.
        details = await run_batch(file, [
            {
                "command": "get", "depth": 2,
                "path": f"{master['path']}/slidelayout[{position}]",
            }
            for position in range(1, master["format"]["layoutCount"] + 1)
        ])
        layouts: list[Layout] = []

        for detail in details:
            layout_index += 1
            layout = detail["output"]["results"][0]
            layouts.append(Layout(
                index=layout_index,
                name=layout["format"].get("name", ""),
                slots=_slots(layout["children"]),
            ))

        listed.append(SlideMaster(
            index=master_index,
            # PowerPoint shows a master under its theme's name.
            name=master["format"].get("theme") or f"Master {master_index}",
            layouts=layouts,
        ))

    properties = root["output"]["results"][0]["format"]

    return PptxInventory(
        slide_size=f"{properties['slideWidth']} x {properties['slideHeight']}",
        slide_masters=listed,
    )


def bind(
    inventory: PptxInventory,
    slide_master: int | None,
) -> PptxInventory:

    names = ", ".join(
        f"{candidate.index} '{candidate.name}'"
        for candidate in inventory.slide_masters
    )

    if slide_master is None and len(inventory.slide_masters) > 1:
        raise ToolError(
            f"The design has several slide masters: {names}. Pass "
            "`slide_master` with the one the user named; if they named none, "
            "ask the user."
        )

    masters = [
        candidate for candidate in inventory.slide_masters
        if slide_master in (None, candidate.index)
    ]

    if not masters:
        raise ToolError(
            f"Slide master {slide_master} not found. Use one of: {names}."
        )

    return inventory.model_copy(update={"slide_masters": masters})


def resolve_layouts(
    inventory: PptxInventory,
    batch: list[dict],
) -> None:

    layouts = [
        layout for master in inventory.slide_masters for layout in master.layouts
    ]
    by_index = {str(layout.index): layout.index for layout in layouts}
    # Names work too, case-insensitively as in OfficeCLI.
    by_name = {layout.name.lower(): layout.index for layout in layouts}

    for command in batch:

        props = command.get("props", {})
        # Diagrams have a `layout` prop of their own.
        for_slide = (
            str(command.get("type", "")).lower() == "slide"
            or _SLIDE_PATH.fullmatch(str(command.get("path", ""))) is not None
        )

        if "layout" not in props or not for_slide:
            continue

        hint = str(props["layout"]).strip()
        index = by_index.get(hint) or by_name.get(hint.lower())

        if index is None:
            raise ToolError(
                f"Layout '{hint}' is not in the bound slide master. Use one of: "
                + ", ".join(f"{layout.index} '{layout.name}'" for layout in layouts)
                + "; nothing was changed."
            )

        # OfficeCLI tries names before numbers; the space makes it a number.
        props["layout"] = f" {index}"


def check(
    deck: dict,
    inventory: PptxInventory,
) -> None:

    slots: dict[str, set[str]] = {}

    for master in inventory.slide_masters:
        for layout in master.layouts:
            slots.setdefault(layout.name, set()).update(layout.slots)

    slides = deck["output"]["results"][0]["children"]

    for number, slide in enumerate(slides, start=1):

        layout = slide["format"].get("layout", "")

        if layout not in slots:
            raise ToolError(
                f"Slide {number} uses layout '{layout}', which is outside the "
                "bound slide master; no document changes were published."
            )

        for slot in _slots(slide["children"]):
            if slot not in slots[layout]:
                raise ToolError(
                    f"Slide {number}: placeholder {slot} does not belong to layout "
                    f"'{layout}', whose slots are "
                    f"{', '.join(sorted(slots[layout]))}. Use `phType` and `idx` "
                    "from `slots`; no document changes were published."
                )
