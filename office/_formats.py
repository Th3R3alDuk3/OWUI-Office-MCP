from pathlib import Path
from types import ModuleType
from xml.etree import ElementTree
from zipfile import BadZipFile, ZipFile

from fastmcp.exceptions import ToolError

from office import docx, pptx, xlsx

FORMATS: dict[str, ModuleType] = {
    module.FORMAT: module for module in (pptx, docx, xlsx)
}

_MAX_MANIFEST_BYTES = 2**20
_MANIFEST = "[Content_Types].xml"


def detect_format(
    file: Path,
) -> ModuleType:

    # OfficeCLI rejects bombs and broken packages; macros only show here.
    try:
        
        with ZipFile(file) as archive:

            if archive.getinfo(_MANIFEST).file_size > _MAX_MANIFEST_BYTES:
                raise ToolError("The file is too large when unpacked.")

            names = set(archive.namelist())
            # Expat blocks entity expansion bombs; external entities never load.
            content_types = {
                entry.get("ContentType")
                for entry in ElementTree.fromstring(archive.read(_MANIFEST))
                if entry.get("PartName", "").lstrip("/") in names
            }

    except (BadZipFile, KeyError, ElementTree.ParseError):
        raise ToolError("The file is not a valid Office document.") from None

    supported = [
        module for module in FORMATS.values()
        if module.CONTENT_TYPE in content_types
    ]

    if len(supported) != 1:
        raise ToolError(
            "Expected one `.pptx`, `.docx` or `.xlsx` document "
            "(no macros or template formats such as `.potx`)."
        )

    return supported[0]
