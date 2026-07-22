"""Compress and watermark.

Not every PDF is compressible the same way, so compress_pdf picks one of two
fundamentally different strategies per file:

  - RECOMPRESS: shrink the embedded images in place (downsample + re-encode
    as JPEG) and subset fonts. Works well, and preserves the document
    exactly as-is otherwise, whenever there's real image data to shrink or
    real selectable text worth keeping crisp.

  - RASTERIZE: flatten every page into a single JPEG image. This is the only
    thing that helps when a PDF's size comes from neither real images nor
    real text — some report/statement generators draw every character as
    vector path outlines instead of text or a picture, which produces
    enormous, barely-compressible content streams with zero selectable
    text. Verified on a real 21MB bank statement built exactly that way:
    RECOMPRESS got 11%; RASTERIZE got 89-96%, and since the document had
    no selectable text to begin with, rasterizing costs nothing that wasn't
    already lost. RASTERIZE is also used at the Extreme tier for ordinary
    text/vector PDFs (Word/PPT/Excel exports) once a user has explicitly
    asked for maximum compression, since that's the only lever left after
    font subsetting and there's nothing else to try.

_choose_strategy() decides which applies to a given file and tier:
  - image_ratio > 50% of file size            -> RECOMPRESS (a scan/photo
    PDF; the current per-image approach already hits 90-96%+ on these)
  - it has real extractable text              -> RECOMPRESS, except at
    Extreme, which RASTERIZEs anyway for the size win once quality is an
    accepted trade-off
  - otherwise (no real text, images are minor) -> RASTERIZE at every tier,
    since there's no text quality to protect

Whichever strategy runs, the output is compared to the original size before
returning — never hand back something bigger than what came in."""
import io
import math
import multiprocessing
import shutil
import tempfile
from pathlib import Path

import fitz
from PIL import Image

# RECOMPRESS: (jpeg quality, target DPI for however large the image is
# actually drawn on the page). Only downsamples — an image already below
# the target DPI is left alone, never upscaled.
LEVELS = {
    "low": {"quality": 75, "target_dpi": 160},
    "recommended": {"quality": 35, "target_dpi": 95},
    "extreme": {"quality": 20, "target_dpi": 65},
}

# RASTERIZE: (jpeg quality, page render DPI). Used for PDFs with no real
# image content and no real text (see module docstring). The "extreme"
# entry here also doubles as the Extreme-tier fallback for ordinary text/
# vector PDFs, tuned gentler than the other two rows since that case is
# giving up real selectable text for the size win, not just flattening an
# already-unselectable document.
RASTER_LEVELS = {
    "low": {"quality": 70, "dpi": 150},
    "recommended": {"quality": 42, "dpi": 92},
    "extreme": {"quality": 25, "dpi": 60},
}
RASTER_EXTREME_FOR_TEXT_PDF = {"quality": 20, "dpi": 45}


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


def _image_ratio(doc: fitz.Document, file_size: int) -> float:
    """Fraction of the file's size that's real embedded image data."""
    total = 0
    seen = set()
    for page in doc:
        for img_info in page.get_images(full=True):
            xref = img_info[0]
            if xref in seen:
                continue
            seen.add(xref)
            try:
                total += len(doc.extract_image(xref)["image"])
            except Exception:
                continue
    return total / file_size if file_size else 0.0


def _has_real_text(doc: fitz.Document, sample_pages: int = 5) -> bool:
    """A very low bar — just "is there any selectable text at all". Some
    statement/report generators draw every character as vector path
    outlines instead of text, which extracts as nothing; rasterizing such a
    file costs no selectability, since it never had any."""
    n = min(sample_pages, doc.page_count)
    total_chars = sum(len(doc[i].get_text().strip()) for i in range(n))
    return total_chars > 20 * n


def _recompress_images(doc: fitz.Document, cfg: dict) -> None:
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


def _rasterize(doc: fitz.Document, cfg: dict) -> fitz.Document:
    """Returns a new document with every page flattened to one JPEG image."""
    out = fitz.open()
    mat = fitz.Matrix(cfg["dpi"] / 72, cfg["dpi"] / 72)
    for page in doc:
        pix = page.get_pixmap(matrix=mat)
        pil_img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
        buf = io.BytesIO()
        pil_img.save(buf, format="JPEG", quality=cfg["quality"], optimize=True)
        newpage = out.new_page(width=page.rect.width, height=page.rect.height)
        newpage.insert_image(newpage.rect, stream=buf.getvalue())
    return out


def compress_pdf(path: str, level: str, save_path: str) -> dict:
    before = Path(path).stat().st_size

    probe = fitz.open(path)
    image_ratio = _image_ratio(probe, before)
    has_text = _has_real_text(probe)
    probe.close()

    if image_ratio > 0.5:
        strategy = "recompress"
    elif not has_text:
        strategy = "rasterize"
    elif level == "extreme":
        strategy = "rasterize"
    else:
        strategy = "recompress"

    if strategy == "rasterize":
        cfg = RASTER_LEVELS[level] if not has_text else RASTER_EXTREME_FOR_TEXT_PDF
        doc = fitz.open(path)
        out = _rasterize(doc, cfg)
        doc.close()
        out.save(save_path, garbage=4, deflate=True, deflate_images=True, clean=True, use_objstms=True, compression_effort=100)
        out.close()
    else:
        cfg = LEVELS.get(level, LEVELS["recommended"])
        subsetted_path = _subset_fonts_isolated(path)
        working_path = subsetted_path or path
        doc = fitz.open(working_path)
        _recompress_images(doc, cfg)
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
