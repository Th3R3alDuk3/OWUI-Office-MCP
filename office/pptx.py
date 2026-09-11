from asyncio import to_thread
from pathlib import Path
from posixpath import basename, dirname, join, normpath
from urllib.parse import unquote
from xml.etree import ElementTree
from zipfile import ZipFile

from fastmcp.exceptions import ToolError

from models.inventory import Layout, Master, PptxInventory
from office import _officecli

FORMAT = "pptx"
MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
CONTENT_TYPE = (
    "application/vnd.openxmlformats-officedocument.presentationml."
    "presentation.main+xml"
)
INVENTORY = PptxInventory

ELEMENT_TYPES = {
    "slide", "placeholder", "paragraph", "notes", "comment", "chart", "picture",
}
CONTENT_PROPS = {
    "layout", "hidden", "phtype", "idx", "text", "level", "image", "author",
    "initials", "charttype", "categories", "data", "title", "legend",
    "datalabels", "src", "x", "y", "width", "height",
}
APPEARANCE_PROPS = {
    "bold", "italic", "underline", "color", "size", "font", "fill", "align",
    "background", "colors",
}
PROTECTED_PATHS = ("/slidemaster", "/slidelayout", "/theme")

START_HINT = """
The project is bound to the master in `inventory.masters`; only its layouts
and slots are accepted. Build with `run_commands`, e.g.:
- slide: {"command": "add", "parent": "/", "type": "slide",
  "props": {"layout": "3"}} — `layout` is the layout's `index`.
- fill a slot of its layout: {"command": "add", "parent": "/slide[1]",
  "type": "placeholder", "props": {"phType": "body", "idx": "1",
  "text": "Point\\nSub-point"}} — `slots` name `phType:idx`; always pass the
  `idx` when a slot has one. `\\n` starts a new paragraph. Nest one:
  {"command": "set", "path": "/slide[1]/placeholder[@idx=1]/paragraph[2]",
  "props": {"level": "1"}}.
- notes: type `notes`, props `text`; review comment: type `comment`, props
  `text` and `author`.
- chart: type `chart`, props `chartType` (bar, line, pie, ...),
  `categories` "Q1,Q2", `data` "Revenue:1,2;Costs:3,4".
- image: into a slot, its placeholder with props `image` "file:<file_id>";
  elsewhere type `picture`, props `src` "file:<file_id>". Place charts and
  images by `x`, `y`, `width`, `height` within `inventory.slide_size`; a
  placeholder's geometry can be read with `get` after adding it. A picture
  with only `width` keeps its aspect ratio; `width`
  and `height` together stretch it.
Colors, fonts and other appearance props need `design_mode="custom"` and an
explicit user request. Paths are 1-based, `index` in add/move is 0-based.
Sample slides of a template are not part of its masters.
""".strip()

# Slots the master fills itself.
_FIXED_SLOTS = {"date", "dt", "footer", "ftr", "header", "hdr", "slidenum", "sldnum"}
_NAMESPACE = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}


async def prepare(
    file: Path,
) -> None:

    (slides,) = await _officecli.run(file, [
        {"command": "query", "selector": "slide"},
    ])

    if paths := [slide["path"] for slide in slides["output"]["results"]]:
        await _officecli.run(file, [
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


def _design(
    file: Path,
) -> tuple[list[Master], dict[str, int], set[str]]:

    masters: list[Master] = []
    layout_ids: dict[str, int] = {}
    layout_names: set[str] = set()
    relationship_id = f"{{{_NAMESPACE['r']}}}id"

    with ZipFile(file) as archive:
        presentation = next(
            target for kind, target in _relationships(archive, "").values()
            if kind == "officeDocument"
        )
        root = ElementTree.fromstring(archive.read(presentation))
        links = _relationships(archive, presentation)
        available = {
            target for kind, target in links.values() if kind == "slideMaster"
        }
        declared = [
            links[node.attrib[relationship_id]][1]
            for node in root.findall("p:sldMasterIdLst/p:sldMasterId", _NAMESPACE)
        ]

        # Match OfficeCLI's write order: declared IDs, then sorted orphan parts.
        for index, master in enumerate(dict.fromkeys([
            *declared, *sorted(available - set(declared)),
        ]), start=1):
            root = ElementTree.fromstring(archive.read(master))
            links = _relationships(archive, master)
            available = {
                target for kind, target in links.values() if kind == "slideLayout"
            }
            declared = [
                links[node.attrib[relationship_id]][1]
                for node in root.findall("p:sldLayoutIdLst/p:sldLayoutId", _NAMESPACE)
            ]
            layouts: list[Layout] = []

            for part in dict.fromkeys([
                *declared, *sorted(available - set(declared)),
            ]):
                layout_ids[part] = len(layout_ids) + 1
                layout = ElementTree.fromstring(archive.read(part))
                common = layout.find("p:cSld", _NAMESPACE)
                name = common.get("name", "") if common is not None else ""
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

            common = root.find("p:cSld", _NAMESPACE)
            name = common.get("name", "") if common is not None else ""
            masters.append(Master(
                index=index, name=name or f"Master {index}", layouts=layouts,
            ))

    return masters, layout_ids, layout_names


async def inventory(
    file: Path,
) -> PptxInventory:

    (root,) = await _officecli.run(file, [
        {"command": "get", "path": "/", "depth": 0},
    ])
    masters, _, _ = await to_thread(_design, file)

    properties = root["output"]["results"][0]["format"]

    return PptxInventory(
        theme=_officecli.theme(root),
        slide_size=f"{properties['slideWidth']} x {properties['slideHeight']}",
        masters=masters,
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

    # Layouts and slots of other masters are rejected from now on.
    return inventory.model_copy(update={"masters": masters})


def adapt(
    command: dict,
    inventory: PptxInventory,
) -> dict:

    props = command.get("props", {})
    layouts = [layout for master in inventory.masters for layout in master.layouts]

    if (
        command["command"] == "add" and command.get("type") == "slide"
        and "layout" not in props
    ):
        raise ToolError("A new slide needs a `layout` index from the inventory.")

    if "layout" in props and str(props["layout"]) not in {
        str(layout.index) for layout in layouts
    }:
        raise ToolError(
            f"Layout '{props['layout']}' not found. Use a layout `index` "
            "from the inventory."
        )

    if command["command"] == "add" and command.get("type") == "placeholder":

        slot = str(props.get("phtype", ""))

        if "idx" in props:
            slot += f":{props['idx']}"

        # The engine would silently add an unbound placeholder.
        if slot.lower() not in {
            known.lower() for layout in layouts for known in layout.slots
        }:
            raise ToolError(
                f"Slot '{slot}' is not in the template's layouts. Use one of "
                "the layout's `slots`, with its `idx`."
            )

    return command


def check(
    file: Path,
    inventory: PptxInventory,
    batch: list[dict] | None = None,
) -> None:

    _, layout_ids, names = _design(file)

    if batch is not None:
        # The engine tries names before integers; whitespace disambiguates them.
        for command in batch:
            props = command.get("props", {})
            if "layout" in props:
                hint = str(props["layout"])
                while hint in names:
                    hint = f" {hint}"
                props["layout"] = hint
        return

    allowed = {
        layout.index for master in inventory.masters for layout in master.layouts
    }
    with ZipFile(file) as archive:
        presentation = next(
            target for kind, target in _relationships(archive, "").values()
            if kind == "officeDocument"
        )
        root = ElementTree.fromstring(archive.read(presentation))
        links = _relationships(archive, presentation)

        for index, slide in enumerate(
            root.findall("p:sldIdLst/p:sldId", _NAMESPACE), start=1,
        ):
            part = links[slide.attrib[f"{{{_NAMESPACE['r']}}}id"]][1]
            layout = next((
                target for kind, target in _relationships(archive, part).values()
                if kind == "slideLayout"
            ), None)
            if layout is None or layout_ids.get(layout) not in allowed:
                raise ToolError(
                    f"Slide {index} is outside the selected master. Use one of "
                    "its layouts; no document changes were published."
                )

            slots = {
                (slot.get("type", "obj").lower(), int(slot.get("idx", "0")))
                for slot in ElementTree.fromstring(archive.read(layout)).findall(
                    ".//p:ph", _NAMESPACE,
                )
            }
            for slot in ElementTree.fromstring(archive.read(part)).findall(
                ".//p:ph", _NAMESPACE,
            ):
                slot_type = slot.get("type", "obj").lower()
                slot_index = int(slot.get("idx", "0"))
                if slot_type not in _FIXED_SLOTS and (slot_type, slot_index) not in slots:
                    raise ToolError(
                        f"Slide {index}: placeholder {slot_type}:{slot_index} "
                        "does not belong to its layout. Use that layout's "
                        "slots; no document changes were published."
                    )
