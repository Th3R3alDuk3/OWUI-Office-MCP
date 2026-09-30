from pathlib import Path

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
layout's `index` (or name); fill a slot with `add` type `placeholder` and its
`phType` and `idx` from `slots`; change an existing placeholder with `set` on
its path from `query placeholder`. `\n` in `text` starts a new paragraph.
Charts take `chartType`, `categories` "Q1,Q2" and `data` "Revenue:1,2;Costs:3,4";
tables `data` "H1,H2;r1c1,r1c2"; images come from attached files as `image`
(placeholder) or `src` (picture) "file:<file_id>"; free elements are placed
by `x`, `y`, `width`, `height` within `inventory.slide_size`. Paths are
1-based, `index` in add/move is 0-based. Stay within the design: change
colors, fonts and other appearance only when the user asks. `get_reference`
documents elements and props.
""".strip()

# Slots the master fills itself.
_FIXED_SLOTS = {"date", "footer", "header", "slidenum"}


async def prepare(
    file: Path,
) -> None:

    slides, presentation = await run_batch(file, [
        {"command": "query", "selector": "slide"},
        {"command": "raw", "part": "/presentation"},
    ])
    commands = [
        # Highest first, so the remaining paths stay valid.
        {"command": "remove", "path": slide["path"]}
        for slide in reversed(slides["output"]["results"])
    ]

    # Sections would keep pointing at the removed slides.
    if "sectionLst" in presentation["output"]:
        commands.append({
            "command": "raw-set", "part": "/presentation", "action": "remove",
            "xpath": "//p:extLst/p:ext[*[local-name()='sectionLst']]",
        })

    if commands:
        await run_batch(file, commands)


def _slots(
    shapes: list[dict],
) -> list[str]:

    slots: list[str] = []

    for shape in shapes:
        kind = shape["format"].get("phType", "")
        if kind and kind not in _FIXED_SLOTS:
            slots.append(f"{kind}:{shape['format'].get('phIndex', 0)}")

    return slots


async def inventory(
    file: Path,
) -> PptxInventory:

    presentation, masters, layouts = await run_batch(file, [
        {"command": "get", "path": "/", "depth": 0},
        {"command": "query", "selector": "slidemaster"},
        {"command": "query", "selector": "slidelayout"},
    ])
    details = await run_batch(file, [
        {"command": "get", "path": layout["path"], "depth": 2}
        for layout in layouts["output"]["results"]
    ])

    listed: list[Master] = []
    # The engine numbers layouts through all masters, in master order.
    index = 0

    for number, master in enumerate(masters["output"]["results"], start=1):

        count = master["format"]["layoutCount"]
        owned: list[Layout] = []

        for detail in details[index:index + count]:
            index += 1
            layout = detail["output"]["results"][0]
            owned.append(Layout(
                index=index,
                name=layout["format"].get("name", ""),
                slots=_slots(layout["children"]),
            ))

        listed.append(Master(
            index=number,
            # PowerPoint shows a master under its theme's name.
            name=master["format"].get("theme") or f"Master {number}",
            layouts=owned,
        ))

    properties = presentation["output"]["results"][0]["format"]

    return PptxInventory(
        slide_size=f"{properties['slideWidth']} x {properties['slideHeight']}",
        masters=listed,
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


def resolve_layouts(
    inventory: PptxInventory,
    batch: list[dict],
) -> None:

    layouts = [layout for master in inventory.masters for layout in master.layouts]
    indexes = {layout.index for layout in layouts}
    by_name = {layout.name: layout.index for layout in layouts}

    for command in batch:

        props = command.get("props", {})

        if "layout" not in props:
            continue

        hint = str(props["layout"]).strip()
        index = int(hint) if hint.isdigit() else by_name.get(hint)

        if index not in indexes:
            raise ToolError(
                f"Layout '{hint}' is not in the selected master. Use one of: "
                + ", ".join(f"{layout.index} '{layout.name}'" for layout in layouts)
                + "; nothing was changed."
            )

        # The engine tries names before integers; the space makes it a number.
        props["layout"] = f" {index}"


async def check(
    file: Path,
    inventory: PptxInventory,
) -> None:

    (deck,) = await run_batch(file, [
        {"command": "get", "path": "/", "depth": 2},
    ])
    slots: dict[str, set[str]] = {}

    for master in inventory.masters:
        for layout in master.layouts:
            slots.setdefault(layout.name, set()).update(layout.slots)

    for number, slide in enumerate(deck["output"]["results"][0]["children"], start=1):

        layout = slide["format"].get("layout", "")

        if layout not in slots:
            raise ToolError(
                f"Slide {number} uses layout '{layout}', which is outside the "
                "selected master; no document changes were published."
            )

        for slot in _slots(slide["children"]):
            if slot not in slots[layout]:
                raise ToolError(
                    f"Slide {number}: placeholder {slot} does not belong to layout "
                    f"'{layout}', whose slots are {', '.join(sorted(slots[layout]))}. "
                    "Use `phType` and `idx` from `slots`; no document changes were "
                    "published."
                )
