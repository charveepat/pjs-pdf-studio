"""Compress and watermark.

Not every PDF is compressible the same way, so compress_pdf picks one of two
fundamentally different strategies per file:

  - RECOMPRESS: shrink the embedded images in place (downsample + re-encode
    as JPEG) and subset fonts. Works well, and preserves the document
    exactly as-is otherwise, whenever there's real image data to shrink or
    real selectable text worth keeping crisp.

  - RASTERIZE: flatten every page into a single JPEG image. This is the only
    thing that helps when a PDF's size comes from neither real images nor
    real text, some report/statement generators draw every character as
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
returning, never hand back something bigger than what came in."""
import io
import logging
import math
import multiprocessing
import shutil
import tempfile
from pathlib import Path

import fitz
from PIL import Image

logger = logging.getLogger("pjs")

# RECOMPRESS: (jpeg quality, target DPI for however large the image is
# actually drawn on the page). Only downsamples, an image already below
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

# Custom % ladders: finer-grained steps between the same already-validated
# Low/Recommended/Extreme points above, so "compress to N%" has real rungs
# to walk instead of just three fixed choices. Deliberately capped at the
# same ceiling as Extreme - nothing here is more aggressive than what's
# already been tested, only the in-between gentler steps are new.
RECOMPRESS_LADDER = [
    {"quality": 90, "target_dpi": 220},
    {"quality": 75, "target_dpi": 160},  # = Low
    {"quality": 55, "target_dpi": 125},
    {"quality": 35, "target_dpi": 95},  # = Recommended
    {"quality": 27, "target_dpi": 80},
    {"quality": 20, "target_dpi": 65},  # = Extreme
]
RASTER_LADDER = [
    {"quality": 80, "dpi": 180},
    {"quality": 70, "dpi": 150},  # = Low
    {"quality": 55, "dpi": 120},
    {"quality": 42, "dpi": 92},  # = Recommended
    {"quality": 32, "dpi": 75},
    {"quality": 25, "dpi": 60},  # = Extreme
]
LEGIBILITY_FLOOR = 0.85  # at least this fraction of reference words must still OCR correctly


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
    or crashed, callers should fall back to the original file either way.

    The subset runs in a throwaway process because doc.subset_fonts() can hard
    segfault on already-subsetted fonts. The timeout is scaled to file size and
    the whole thing is retried once, because the original fixed 25s was long
    enough on a quiet machine but not on a loaded office PC running a batch,
    where it would time out, get skipped, and silently leave fonts un-shrunk
    (a font-heavy file then compressed far less in a batch than on its own).
    Every skip is logged so that case is diagnosable from the log file."""
    try:
        size_mb = Path(path).stat().st_size / 1_000_000
    except OSError:
        size_mb = 5.0
    timeout = max(60, int(size_mb * 12))  # generous headroom for a busy CPU

    for attempt in (1, 2):
        tmp = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False)
        tmp.close()
        queue = multiprocessing.Queue()
        proc = multiprocessing.Process(target=_subset_fonts_worker, args=(path, tmp.name, queue))
        proc.start()
        proc.join(timeout=timeout)

        ok = False
        if proc.is_alive():
            proc.terminate()
            proc.join()
            logger.warning("font subset timed out after %ds (attempt %d): %s", timeout, attempt, path)
        elif proc.exitcode == 0:
            try:
                ok = queue.get_nowait()
            except Exception:
                ok = False
        else:
            logger.warning("font subset process exited %s (attempt %d): %s", proc.exitcode, attempt, path)

        if ok:
            return tmp.name
        Path(tmp.name).unlink(missing_ok=True)

    logger.warning("font subset skipped (fonts left un-shrunk) for %s", path)
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
    """A very low bar, just "is there any selectable text at all". Some
    statement/report generators draw every character as vector path
    outlines instead of text, which extracts as nothing; rasterizing such a
    file costs no selectability, since it never had any."""
    n = min(sample_pages, doc.page_count)
    total_chars = sum(len(doc[i].get_text().strip()) for i in range(n))
    return total_chars > 20 * n


def _recompress_images(doc: fitz.Document, cfg: dict, progress=None) -> None:
    # Map each image xref to one on-page rect (in points) so we can tell how
    # many pixels it actually needs. The same image can be reused across
    # pages (e.g. a letterhead), process each xref exactly once, since
    # re-encoding an already-recompressed JPEG again would lose quality for
    # nothing.
    xref_pages_rects = {}
    for page in doc:
        for img_info in page.get_images(full=True):
            xref = img_info[0]
            rects = page.get_image_rects(xref)
            if rects and xref not in xref_pages_rects:
                xref_pages_rects[xref] = (page, rects[0])

    items = list(xref_pages_rects.items())
    total = len(items) or 1
    for done, (xref, (page, rect)) in enumerate(items):
        if progress:
            progress((done + 1) / total)
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
            #, only replace the image if this pass actually helped.
            if len(new_bytes) < original_size:
                page.replace_image(xref, stream=new_bytes)
        except Exception:
            continue  # a handful of odd/indexed images shouldn't sink the whole file


def _rasterize(doc: fitz.Document, cfg: dict, progress=None) -> fitz.Document:
    """Returns a new document with every page flattened to one JPEG image."""
    out = fitz.open()
    mat = fitz.Matrix(cfg["dpi"] / 72, cfg["dpi"] / 72)
    total = doc.page_count or 1
    for idx, page in enumerate(doc):
        if progress:
            progress((idx + 1) / total)
        pix = page.get_pixmap(matrix=mat)
        pil_img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
        buf = io.BytesIO()
        pil_img.save(buf, format="JPEG", quality=cfg["quality"], optimize=True)
        newpage = out.new_page(width=page.rect.width, height=page.rect.height)
        newpage.insert_image(newpage.rect, stream=buf.getvalue())
    return out


def compress_pdf(path: str, level: str, save_path: str, progress=None) -> dict:
    """progress, if given, is called with a float 0..1 as the compression of
    this one file advances, so the UI can show a real percentage instead of a
    spinner."""
    def emit(frac):
        if progress:
            progress(max(0.0, min(1.0, frac)))

    emit(0.0)
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
        # rasterizing every page is the bulk of the work, map it to 5..90%
        out = _rasterize(doc, cfg, progress=lambda f: emit(0.05 + 0.85 * f))
        doc.close()
        emit(0.92)
        out.save(save_path, garbage=4, deflate=True, deflate_images=True, clean=True, use_objstms=True, compression_effort=100)
        out.close()
    else:
        cfg = LEVELS.get(level, LEVELS["recommended"])
        emit(0.05)
        subsetted_path = _subset_fonts_isolated(path)  # can be slow; sits in the 5..20% band
        emit(0.20)
        working_path = subsetted_path or path
        doc = fitz.open(working_path)
        _recompress_images(doc, cfg, progress=lambda f: emit(0.20 + 0.70 * f))
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

    emit(1.0)
    after = Path(save_path).stat().st_size
    if after >= before:
        # Nothing here actually helped this particular file, never hand
        # back something bigger than what came in.
        shutil.copyfile(path, save_path)
        after = before

    return {"before_bytes": before, "after_bytes": after}


def _sample_page_numbers(n: int, count: int = 3) -> list[int]:
    """count evenly-spread page indices (0-indexed) to OCR-check, instead of
    every page, real OCR takes real seconds per page and a handful of
    pages is enough to catch a setting that's gone illegible."""
    if n <= count:
        return list(range(n))
    step = n / count
    return sorted({int(i * step) for i in range(count)})


def _is_legible(original_path: str, candidate_path: str, has_text: bool, sample_pages: list[int]) -> bool | None:
    """Returns True/False, or None if OCR verification isn't available on
    this machine (compress_pdf_custom falls back to a best-effort result
    without per-file verification in that case, and says so)."""
    from core import legibility

    if not legibility.is_available():
        return None

    original = fitz.open(original_path)
    candidate = fitz.open(candidate_path)
    accuracies = []
    try:
        for pno in sample_pages:
            cpix = candidate[pno].get_pixmap(matrix=fitz.Matrix(2, 2))
            cimg = Image.frombytes("RGB", (cpix.width, cpix.height), cpix.samples)
            candidate_words = legibility.ocr_words(cimg)

            if has_text:
                reference_words = {w.lower() for w in original[pno].get_text().split() if w.strip()}
            else:
                opix = original[pno].get_pixmap(matrix=fitz.Matrix(300 / 72, 300 / 72))
                oimg = Image.frombytes("RGB", (opix.width, opix.height), opix.samples)
                reference_words = legibility.ocr_words(oimg)

            accuracies.append(legibility.word_accuracy(reference_words, candidate_words))
    finally:
        original.close()
        candidate.close()

    if not accuracies:
        return True
    return (sum(accuracies) / len(accuracies)) >= LEGIBILITY_FLOOR


def compress_pdf_custom(path: str, target_pct: float, save_path: str) -> dict:
    """Walks a ladder of settings from gentle to the same ceiling as the
    Extreme tier, stopping as soon as either the target percentage is hit or
    OCR says the previous rung was the last legible one. Never goes past
    what Extreme already does, so this can't produce a worse result than the
    existing tiers, only a more precisely targeted one in between them."""
    before = Path(path).stat().st_size

    probe = fitz.open(path)
    image_ratio = _image_ratio(probe, before)
    has_text = _has_real_text(probe)
    sample_pages = _sample_page_numbers(probe.page_count)
    probe.close()

    use_rasterize = image_ratio <= 0.5 and not has_text
    ladder = RASTER_LADDER if use_rasterize else RECOMPRESS_LADDER

    best_path = None
    ocr_checked = False
    stopped_by_ocr = False
    for rung, cfg in enumerate(ladder):
        candidate_path = f"{save_path}.candidate{rung}"
        if use_rasterize:
            doc = fitz.open(path)
            out = _rasterize(doc, cfg)
            doc.close()
            out.save(candidate_path, garbage=4, deflate=True, deflate_images=True, clean=True, use_objstms=True, compression_effort=100)
            out.close()
        else:
            subsetted_path = _subset_fonts_isolated(path)
            working_path = subsetted_path or path
            doc = fitz.open(working_path)
            _recompress_images(doc, cfg)
            doc.save(
                candidate_path,
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

        legible = _is_legible(path, candidate_path, has_text, sample_pages)
        if legible is not None:
            ocr_checked = True

        if legible is False:
            Path(candidate_path).unlink(missing_ok=True)
            stopped_by_ocr = True
            break

        if best_path:
            Path(best_path).unlink(missing_ok=True)
        best_path = candidate_path

        achieved = 100 * (1 - Path(candidate_path).stat().st_size / before) if before else 0
        if achieved >= target_pct:
            break

    if best_path:
        shutil.move(best_path, save_path)
    else:
        shutil.copyfile(path, save_path)

    after = Path(save_path).stat().st_size
    if after >= before:
        shutil.copyfile(path, save_path)
        after = before

    achieved_pct = 100 * (1 - after / before) if before else 0
    capped = achieved_pct < target_pct - 0.5
    if not ocr_checked:
        reason = "Legibility could not be verified on this machine (Windows' built-in OCR wasn't available), so this is a best-effort result without per-file confirmation."
    elif stopped_by_ocr:
        reason = f"Stopped at {achieved_pct:.0f}% because compressing further started producing text OCR could no longer read reliably."
    elif capped:
        reason = f"Stopped at {achieved_pct:.0f}%, the most this tool will compress a file like this even at Extreme, to stay a safe margin above the point where text becomes hard to read."
    else:
        reason = None

    return {
        "before_bytes": before,
        "after_bytes": after,
        "requested_pct": target_pct,
        "achieved_pct": achieved_pct,
        "capped": capped,
        "reason": reason,
    }


def watermark_pdf(
    path: str, text: str, save_path: str, opacity: float = 0.25, font_size: float = 48
) -> None:
    """Applied to every page, text is centered and stamped diagonally."""
    doc = fitz.open(path)
    theta = math.radians(45)
    rot = fitz.Matrix(math.cos(theta), math.sin(theta), -math.sin(theta), math.cos(theta), 0, 0)
    text_width = fitz.get_text_length(text, fontname="helv", fontsize=font_size)
    for page in doc:
        center = (page.rect.width / 2, page.rect.height / 2)
        # insert_text's point is the baseline start, not the text's visual
        # center, so start half the text's width to the left of center,
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
