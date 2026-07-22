"""Compress and watermark.

Compression uses three levers, applied in this order:
  1. Font subsetting (doc.subset_fonts()) — embedded font programs are
     trimmed down to only the glyphs actually used. Run in an isolated
     subprocess (see _subset_fonts_isolated below) because PyMuPDF's
     subset_fonts() can segfault outright on PDFs whose fonts are already
     subsetted — which is the normal case for most Word/PowerPoint/
     Illustrator-generated PDFs, since those tools subset on export too.
     A segfault is a native crash, not a Python exception, so no amount of
     try/except in this file could ever catch it; only running it in a
     throwaway process can contain it.
  2. Downsampling embedded images to a target resolution for how large
     they're actually drawn on the page — a 300 DPI scan is still a
     300 DPI scan at low JPEG quality unless the pixel count itself comes
     down too.
  3. Re-encoding at a target JPEG quality, but only keeping the result if
     it's actually smaller — an image that's already an efficiently-encoded
     JPEG below the target resolution can come out *larger* after a naive
     re-encode, which is what made "Low" inflate files instead of shrinking
     them.
Images are replaced via Page.replace_image(), which handles the PDF object
bookkeeping correctly; hand-editing the xref's stream/filter keys directly
(an earlier version of this function did that) produces PDFs whose JPEG
streams silently corrupt on save.

Finally, the whole output is compared to the original size — a PDF that's
already efficiently packed can legitimately have nothing left to save once
every image is skipped, and re-saving it can add a little structural
overhead. Rather than hand back something bigger than what came in, we
fall back to the original bytes and report 0% saved."""
import io
import math
import multiprocessing
import shutil
import tempfile
from pathlib import Path

import fitz
from PIL import Image

# (jpeg quality, target DPI for however large the image is actually drawn on
# the page). Only downsamples — an image already below the target DPI is
# left alone, never upscaled.
LEVELS = {
    "low": {"quality": 75, "target_dpi": 160},
    "recommended": {"quality": 48, "target_dpi": 115},
    "extreme": {"quality": 25, "target_dpi": 78},
}


def _subset_fonts_worker(path: str, out_path: str, result_queue):
    try:
        doc = fitz.open(path)
        doc.subset_fonts()
        doc.save(out_path)
        doc.close()
        result_queue.put(True)
    except Exception:
        result_queue.put(False)


def _subset_fonts_isolated(path: str) -> str | None:
    """Returns a path to a font-subsetted copy, or None if subsetting failed
    or crashed — callers should fall back to the original file either way."""
    tmp = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False)
    tmp.close()
    queue = multiprocessing.Queue()
    proc = multiprocessing.Process(target=_subset_fonts_worker, args=(path, tmp.name, queue))
    proc.start()
    proc.join(timeout=25)

    ok = False
    if proc.is_alive():
        proc.terminate()
        proc.join()
    elif proc.exitcode == 0:
        try:
            ok = queue.get_nowait()
        except Exception:
            ok = False

    if ok:
        return tmp.name
    Path(tmp.name).unlink(missing_ok=True)
    return None


def compress_pdf(path: str, level: str, save_path: str) -> dict:
    cfg = LEVELS.get(level, LEVELS["recommended"])
    before = Path(path).stat().st_size

    subsetted_path = _subset_fonts_isolated(path)
    working_path = subsetted_path or path
    doc = fitz.open(working_path)

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
            original_size = len(doc.extract_image(xref)["image"])

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
            new_bytes = buf.getvalue()

            # Re-encoding an already-efficient JPEG at a conservative quality
            # (typical of the "low" tier) can come out bigger than it started
            # — only replace the image if this pass actually helped.
            if len(new_bytes) < original_size:
                page.replace_image(xref, stream=new_bytes)
        except Exception:
            continue  # a handful of odd/indexed images shouldn't sink the whole file

    doc.save(
        save_path,
        garbage=4,
        deflate=True,
        deflate_images=True,
        deflate_fonts=True,
        clean=True,
        use_objstms=True,
        compression_effort=100,
    )
    doc.close()
    if subsetted_path:
        Path(subsetted_path).unlink(missing_ok=True)

    after = Path(save_path).stat().st_size
    if after >= before:
        # Nothing here actually helped this particular file — never hand
        # back something bigger than what came in.
        shutil.copyfile(path, save_path)
        after = before

    return {"before_bytes": before, "after_bytes": after}


def watermark_pdf(
    path: str, text: str, save_path: str, opacity: float = 0.25, font_size: float = 48
) -> None:
    """Applied to every page — text is centered and stamped diagonally."""
    doc = fitz.open(path)
    theta = math.radians(45)
    rot = fitz.Matrix(math.cos(theta), math.sin(theta), -math.sin(theta), math.cos(theta), 0, 0)
    text_width = fitz.get_text_length(text, fontname="helv", fontsize=font_size)
    for page in doc:
        center = (page.rect.width / 2, page.rect.height / 2)
        # insert_text's point is the baseline start, not the text's visual
        # center, so start half the text's width to the left of center —
        # otherwise the stamp reads as noticeably off-center on the page.
        origin = (center[0] - text_width / 2, center[1])
        page.insert_text(
            origin,
            text,
            fontsize=font_size,
            fontname="helv",
            color=(0.5, 0.5, 0.5),
            fill_opacity=opacity,
            morph=(fitz.Point(center), rot),
        )
    doc.save(save_path)
    doc.close()
