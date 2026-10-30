from pyvips import Error, ForeignKeep, Image, block_untrusted_set

# Uploads are untrusted: fuzzed loaders only.
block_untrusted_set(True)

_SHRINK_FROM_BYTES = 1 * 2**20
_MAX_EDGE_PIXELS = 2560
_MIN_SAVING = 0.25
_JPEG = b"\xff\xd8\xff"
_PNG = b"\x89PNG\r\n\x1a\n"
# Camera, location and author details.
_METADATA_FIELDS = {"exif-data", "xmp-data", "iptc-data"}


def shrink_image(
    data: bytes,
) -> tuple[bytes, str] | None:

    jpeg = data.startswith(_JPEG)
    small = len(data) <= _SHRINK_FROM_BYTES
    # JPEG and PNG barely grow when re-encoded: metadata goes at any size.
    strippable = jpeg or data.startswith(_PNG)

    if small and not strippable:
        return None

    try:

        source = Image.new_from_buffer(data, "")
        has_metadata = strippable and not _METADATA_FIELDS.isdisjoint(
            source.get_fields(),
        )

        # An animation would keep only its first frame.
        if (small and not has_metadata) or source.get_n_pages() > 1:
            return None

        image = Image.thumbnail_buffer(
            data, _MAX_EDGE_PIXELS, size="down", output_profile="srgb",
        )

        # Photos stay JPEG; the rest may need transparency or sharp edges.
        if jpeg:
            output = image.jpegsave_buffer(
                Q=85, optimize_coding=True, keep=ForeignKeep.NONE,
            )
        else:
            output = image.pngsave_buffer(compression=9, keep=ForeignKeep.NONE)

    except Error:
        return None

    # Without metadata to drop, only a clear saving is worth a re-encode.
    if not has_metadata and len(output) > len(data) * (1 - _MIN_SAVING):
        return None

    return output, ".jpg" if jpeg else ".png"
