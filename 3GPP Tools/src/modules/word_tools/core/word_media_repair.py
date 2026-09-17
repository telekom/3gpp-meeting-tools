# --- File: src/modules/word_tools/core/word_media_repair.py ---
"""
Temporary DOCX media repair helpers for Word PDF export.

The original document is never modified.  On Windows, EMF/WMF images are
rasterized with the native GDI+ API and their DOCX relationships are redirected
to PNG copies.  This is intentionally dependency-free: PyQt/Pillow/ImageMagick
are not required.
"""
import ctypes
import shutil
import tempfile
import zipfile
from ctypes import wintypes
from pathlib import Path
from typing import Callable, List, Optional, Sequence, Tuple
from xml.etree import ElementTree as ET

CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

ET.register_namespace("", CONTENT_TYPES_NS)
ET.register_namespace("", REL_NS)


class GdiplusStartupInput(ctypes.Structure):
    _fields_ = [
        ("GdiplusVersion", ctypes.c_uint32),
        ("DebugEventCallback", ctypes.c_void_p),
        ("SuppressBackgroundThread", wintypes.BOOL),
        ("SuppressExternalCodecs", wintypes.BOOL),
    ]


class ImageCodecInfo(ctypes.Structure):
    _fields_ = [
        ("Clsid", ctypes.c_byte * 16),
        ("FormatID", ctypes.c_byte * 16),
        ("CodecName", wintypes.LPWSTR),
        ("DllName", wintypes.LPWSTR),
        ("FormatDescription", wintypes.LPWSTR),
        ("FilenameExtension", wintypes.LPWSTR),
        ("MimeType", wintypes.LPWSTR),
        ("Flags", ctypes.c_uint32),
        ("Version", ctypes.c_uint32),
        ("SigCount", ctypes.c_uint32),
        ("SigSize", ctypes.c_uint32),
        ("SigPattern", ctypes.c_void_p),
        ("SigMask", ctypes.c_void_p),
    ]


def _status_ok(status: int, operation: str) -> None:
    if status != 0:
        raise RuntimeError(f"GDI+ {operation} failed with status {status}")


def _png_encoder_clsid(gdiplus) -> bytes:
    count = ctypes.c_uint()
    size = ctypes.c_uint()
    _status_ok(
        gdiplus.GdipGetImageEncodersSize(ctypes.byref(count), ctypes.byref(size)),
        "GdipGetImageEncodersSize",
    )
    if not count.value or not size.value:
        raise RuntimeError("GDI+ reported no image encoders.")

    buf = ctypes.create_string_buffer(size.value)
    _status_ok(
        gdiplus.GdipGetImageEncoders(count.value, size.value, buf),
        "GdipGetImageEncoders",
    )
    codecs = ctypes.cast(buf, ctypes.POINTER(ImageCodecInfo))
    for i in range(count.value):
        mime = codecs[i].MimeType or ""
        if mime.lower() == "image/png":
            return bytes(codecs[i].Clsid)
    raise RuntimeError("GDI+ PNG encoder was not found.")


def rasterize_metafile_to_png(source: Path, target: Path) -> None:
    """Rasterize EMF/WMF to PNG using Windows GDI+."""
    if not hasattr(ctypes, "windll"):
        raise RuntimeError("EMF/WMF repair is available only on Windows.")

    gdiplus = ctypes.windll.gdiplus
    token = ctypes.c_void_p()
    startup = GdiplusStartupInput(1, None, False, False)
    _status_ok(
        gdiplus.GdiplusStartup(ctypes.byref(token), ctypes.byref(startup), None),
        "startup",
    )

    image = ctypes.c_void_p()
    bitmap = ctypes.c_void_p()
    graphics = ctypes.c_void_p()

    try:
        _status_ok(
            gdiplus.GdipLoadImageFromFile(str(source), ctypes.byref(image)),
            "load metafile",
        )

        width = ctypes.c_uint()
        height = ctypes.c_uint()
        _status_ok(gdiplus.GdipGetImageWidth(image, ctypes.byref(width)), "get width")
        _status_ok(gdiplus.GdipGetImageHeight(image, ctypes.byref(height)), "get height")

        if width.value == 0 or height.value == 0:
            raise RuntimeError(f"Metafile has invalid dimensions: {width.value}x{height.value}")

        # Render at native GDI+ dimensions.  Word preserves the display size via
        # DrawingML extents, so changing the media pixels does not change layout.
        _status_ok(
            gdiplus.GdipCreateBitmapFromScan0(
                width.value, height.value, 0, 0x26200A, None, ctypes.byref(bitmap)
            ),
            "create bitmap",
        )
        _status_ok(gdiplus.GdipGetImageGraphicsContext(bitmap, ctypes.byref(graphics)), "graphics")
        _status_ok(
            gdiplus.GdipDrawImageRectI(graphics, image, 0, 0, width.value, height.value),
            "draw metafile",
        )

        clsid_bytes = _png_encoder_clsid(gdiplus)
        clsid = (ctypes.c_byte * 16).from_buffer_copy(clsid_bytes)
        target.parent.mkdir(parents=True, exist_ok=True)
        _status_ok(
            gdiplus.GdipSaveImageToFile(bitmap, str(target), ctypes.byref(clsid), None),
            "save PNG",
        )
    finally:
        if graphics:
            gdiplus.GdipDeleteGraphics(graphics)
        if bitmap:
            gdiplus.GdipDisposeImage(bitmap)
        if image:
            gdiplus.GdipDisposeImage(image)
        gdiplus.GdiplusShutdown(token)


def list_legacy_metafiles(docx_path: Path) -> List[str]:
    """Return DOCX package media paths ending in .emf/.wmf."""
    with zipfile.ZipFile(docx_path, "r") as zf:
        return sorted(
            name for name in zf.namelist()
            if name.lower().startswith("word/media/")
            and name.lower().endswith((".emf", ".wmf"))
        )


def _rewrite_relationships(xml_bytes: bytes, old_name: str, new_name: str) -> bytes:
    root = ET.fromstring(xml_bytes)
    changed = False
    old_tail = f"media/{old_name}"
    new_tail = f"media/{new_name}"

    for rel in root:
        target = rel.attrib.get("Target", "")
        normalized = target.replace("\\", "/")
        if normalized.endswith(old_tail) or normalized == old_tail:
            prefix = target[:-len(old_tail)]
            rel.set("Target", prefix + new_tail)
            changed = True

    if not changed:
        return xml_bytes
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _ensure_png_content_type(xml_bytes: bytes) -> bytes:
    root = ET.fromstring(xml_bytes)
    for child in root:
        if child.tag.endswith("Default") and child.attrib.get("Extension", "").lower() == "png":
            return xml_bytes

    ET.SubElement(
        root,
        f"{{{CONTENT_TYPES_NS}}}Default",
        {"Extension": "png", "ContentType": "image/png"},
    )
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def create_docx_with_rasterized_media(
    source_docx: Path,
    output_docx: Path,
    media_paths: Sequence[str],
) -> List[Tuple[str, str]]:
    """
    Create a temporary DOCX where selected EMF/WMF media are PNG.

    Returns [(old_package_path, new_package_path), ...].
    """
    selected = set(media_paths)
    if not selected:
        shutil.copy2(source_docx, output_docx)
        return []

    work = Path(tempfile.mkdtemp(prefix="3gpp_emf_render_"))
    replacements = {}

    try:
        with zipfile.ZipFile(source_docx, "r") as zin:
            names = set(zin.namelist())
            for package_path in selected:
                if package_path not in names:
                    raise RuntimeError(f"DOCX media item not found: {package_path}")

                old_name = Path(package_path).name
                new_name = f"{Path(old_name).stem}_3gpp_raster.png"
                new_package_path = f"word/media/{new_name}"

                source_media = work / old_name
                source_media.write_bytes(zin.read(package_path))
                target_png = work / new_name
                rasterize_metafile_to_png(source_media, target_png)

                if not target_png.exists() or target_png.stat().st_size == 0:
                    raise RuntimeError(f"Rasterization produced no output for {package_path}")

                replacements[package_path] = (new_package_path, target_png)

            output_docx.parent.mkdir(parents=True, exist_ok=True)
            with zipfile.ZipFile(output_docx, "w", zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    if item.filename in selected:
                        continue

                    data = zin.read(item.filename)

                    if item.filename == "[Content_Types].xml":
                        data = _ensure_png_content_type(data)
                    elif item.filename.lower().endswith(".rels"):
                        for old_package_path, (new_package_path, _) in replacements.items():
                            data = _rewrite_relationships(
                                data,
                                Path(old_package_path).name,
                                Path(new_package_path).name,
                            )

                    zout.writestr(item, data)

                for _, (new_package_path, png_path) in replacements.items():
                    zout.write(png_path, new_package_path)

        return [(old, new) for old, (new, _) in replacements.items()]
    finally:
        shutil.rmtree(work, ignore_errors=True)


def create_pdf_safe_docx(
    source_docx: Path,
    output_docx: Path,
    log: Optional[Callable[[str], None]] = None,
) -> List[Tuple[str, str]]:
    """
    Create one temporary PDF-safe DOCX by rasterizing all EMF/WMF media.

    This is intentionally optimized for the normal conversion path: after Word
    has already failed once, rebuilding/exporting the full specification once is
    far cheaper than testing every legacy graphic with a complete Word export.
    The original DOCX is never modified.
    """
    media = list_legacy_metafiles(source_docx)
    if not media:
        return []

    if log:
        log(
            f"⏳ Creating temporary PDF-safe copy: rasterizing "
            f"{len(media)} EMF/WMF graphic(s)..."
        )

    replacements = create_docx_with_rasterized_media(
        source_docx=source_docx,
        output_docx=output_docx,
        media_paths=media,
    )

    if log:
        log(
            f"✅ Temporary PDF-safe copy created with "
            f"{len(replacements)} rasterized legacy graphic(s)."
        )

    return replacements
