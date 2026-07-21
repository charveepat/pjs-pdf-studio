"""Compress and watermark.

Compression uses the same two levers real PDF compressors (Ghostscript's
/screen, /ebook presets; iLovePDF, etc.) rely on: downsampling embedded
images to a target resolution for how large they're actually drawn on the
page, and re-encoding at a target JPEG quality. Quality alone barely moves
the needle on a typical scanned/photo-heavy PDF — a 300 DPI page scan is
still a 300 DPI page scan at quality 35 unless the pixel count itself comes
down too. Images are replaced via Page.replace_image(), which handles the
PDF object bookkeeping correctly; hand-editing the xref's stream/filter keys
directly (an earlier version of this function did that) produces PDFs whose
JPEG streams silently corrupt on save."""
import io
import math
from pathlib import Path

import fitz
from PIL import Image

# (jpeg quality, target DPI for however large the image is actually drawn on
# the page). Only downsamples — an image already below the target DPI is
# left alone, never upscaled.
LEVELS = {
    "low": {"quality": 82, "target_dpi": 200},
    "recommended": {"quality": 65, "target_dpi": 150},
    "extreme": {"quality": 40, "target_dpi": 96},
}


def compress_pdf(path: str, level: str, save_path: str) -> dict:
    cfg = LEVELS.get(level, LEVELS["recommended"])
    before = Path(path).stat().st_size

    doc = fitz.open(path)

    # Map each image xref to one on-page rect (in points) so we can tell how
    # many pixels it actually needs. The same image can be reused across
    # pages (e.g. a letterhead) — process each xref exactly once, since
    # re-encoding an already-recompressed JPEG again would lose quality for
    # nothing.
    xref_pages_rects = {}
    for page in doc:
        for img_info in page.get_images(full=True):
            xref = img_info[0]
            rects = page.get_image_rects(xref)
            if rects and xref not in xref_pages_rects:
                xref_pages_rects[xref] = (page, rects[0])

    for xref, (page, rect) in xref_pages_rects.items():
        try:
            pix = fitz.Pixmap(doc, xref)
            if pix.colorspace is None:  # e.g. a stencil/mask image, leave alone
                continue
            if pix.n - pix.alpha >= 4:  # CMYK -> RGB; JPEG can't hold CMYK reliably
                pix = fitz.Pixmap(fitz.csRGB, pix)
            if pix.alpha:  # JPEG has no alpha channel
                pix = fitz.Pixmap(pix, 0)

            mode = "L" if pix.n == 1 else "RGB"
            pil_img = Image.frombytes(mode, (pix.width, pix.height), pix.samples)

            draw_width_in = rect.width / 72.0
            if draw_width_in > 0:
                effective_dpi = pix.width / draw_width_in
                if effective_dpi > cfg["target_dpi"]:
                    scale = cfg["target_dpi"] / effective_dpi
                    new_size = (max(1, round(pix.width * scale)), max(1, round(pix.height * scale)))
                    pil_img = pil_img.resize(new_size, Image.LANCZOS)

            buf = io.BytesIO()
            pil_img.save(buf, format="JPEG", quality=cfg["quality"], optimize=True)
            page.replace_image(xref, stream=buf.getvalue())
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
