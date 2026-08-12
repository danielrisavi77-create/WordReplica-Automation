from pathlib import PurePosixPath

from word_replica.domain.enums import CapabilityClass
from word_replica.domain.model import DrawingRef
from word_replica.domain.reconstruction import CapabilityDecision


SUPPORTED_IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tif", ".tiff", ".emf", ".wmf"}
_CONTENT_TYPE_EXTENSION = {
    "image/png": ".png", "image/jpeg": ".jpg", "image/gif": ".gif",
    "image/bmp": ".bmp", "image/tiff": ".tiff",
    "image/x-emf": ".emf", "image/emf": ".emf",
    "image/x-wmf": ".wmf", "image/wmf": ".wmf",
}


def detect_image_extension(source_path: str, asset=None) -> str:
    extension = PurePosixPath(source_path).suffix.lower()
    if extension in SUPPORTED_IMAGE_EXTENSIONS:
        return extension
    if asset is not None:
        content_type = str(getattr(asset, "content_type", "") or "").lower()
        if content_type in _CONTENT_TYPE_EXTENSION:
            return _CONTENT_TYPE_EXTENSION[content_type]
        data = bytes(getattr(asset, "bytes_data", b"") or b"")
        if data.startswith(b"\x89PNG\r\n\x1a\n"):
            return ".png"
        if data.startswith(b"\xff\xd8\xff"):
            return ".jpg"
        if data.startswith((b"GIF87a", b"GIF89a")):
            return ".gif"
        if data.startswith(b"BM"):
            return ".bmp"
        if data.startswith((b"II*\x00", b"MM\x00*")):
            return ".tiff"
    return extension


def classify_drawing(drawing: DrawingRef, *, asset=None) -> CapabilityDecision:
    representation = drawing.representation.lower()
    if representation == "preserved":
        return CapabilityDecision(CapabilityClass.PRESERVED, "Object can only be preserved as its original OOXML payload", drawing.element_id)
    if representation not in {"inline", "floating"}:
        return CapabilityDecision(CapabilityClass.UNSUPPORTED, f"Unsupported drawing representation: {drawing.representation}", drawing.element_id)
    extension = detect_image_extension(drawing.source_path, asset)
    if extension not in SUPPORTED_IMAGE_EXTENSIONS:
        return CapabilityDecision(CapabilityClass.UNSUPPORTED, f"Unsupported image format: {extension or 'unknown'}", drawing.element_id)
    return CapabilityDecision(CapabilityClass.RECONSTRUCTED, "Supported Word image drawing", drawing.element_id)
