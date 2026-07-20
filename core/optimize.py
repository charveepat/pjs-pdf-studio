"""Compress and watermark. Compression re-encodes embedded images to JPEG at a
target quality after PyMuPDF's own structural cleanup (dedup objects, deflate
streams) — the same two levers Ghostscript's presets rely on."""
import math
from pathlib import Path

import fitz

QUALITY_BY_LEVEL = {"low": 85, "recommended": 60, "extreme": 35}


def compress_pdf(path: str, level: str, save_path: str) -> dict:
    quality = QUALITY_BY_LEVEL.get(level, 60)
    before = Path(path).stat().st_size

    doc = fitz.open(path)
    for xref in range(1, doc.xref_length()):
        try:
            if doc.xref_get_key(xref, "Subtype")[1] != "/Image":
                continue
            pix = fitz.Pixmap(doc, xref)
            if pix.colorspace is None:
                continue
            if pix.n - pix.alpha >= 4:  # CMYK -> RGB, JPEG can't hold CMYK cleanly here
                pix = fitz.Pixmap(fitz.csRGB, pix)
            if pix.alpha:  # JPEG has no alpha channel
                pix = fitz.Pixmap(pix, 0)
            jpg_bytes = pix.tobytes("jpeg", jpg_quality=quality)
            doc.update_stream(xref, jpg_bytes)
            doc.xref_set_key(xref, "Filter", "/DCTDecode")
            doc.xref_set_key(xref, "ColorSpace", "/DeviceRGB" if pix.n >= 3 else "/DeviceGray")
            doc.xref_set_key(xref, "BitsPerComponent", "8")
            doc.xref_set_key(xref, "Width", str(pix.width))
            doc.xref_set_key(xref, "Height", str(pix.height))
            doc.xref_set_key(xref, "DecodeParms", "null")
        except Exception:
            continue  # a handful of odd/indexed images shouldn't sink the whole file

    doc.save(save_path, garbage=4, deflate=True, clean=True)
    doc.close()

    after = Path(save_path).stat().st_size
    return {"before_bytes": before, "after_bytes": after}


def watermark_pdf(path: str, text: str, save_path: str, opacity: float = 0.25) -> None:
    doc = fitz.open(path)
    theta = math.radians(45)
    rot = fitz.Matrix(math.cos(theta), math.sin(theta), -math.sin(theta), math.cos(theta), 0, 0)
    for page in doc:
        center = (page.rect.width / 2, page.rect.height / 2)
        page.insert_text(
            center,
            text,
            fontsize=48,
            fontname="helv",
            color=(0.5, 0.5, 0.5),
            fill_opacity=opacity,
            morph=(fitz.Point(center), rot),
        )
    doc.save(save_path)
    doc.close()
