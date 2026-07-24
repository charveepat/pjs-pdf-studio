"""Tesseract OCR for scanned (image-only) PDFs.

Tesseract is bundled into the packaged exe as a self-contained program plus its
English language data (see .github/workflows/build-windows.yml), and called via
pytesseract. This is deliberately a bundled binary we invoke, not a Python import
that depends on Windows Runtime projections (the reason the old winocr path
silently failed inside the frozen exe): an external .exe survives PyInstaller
packaging far more reliably.

Running from source (dev), it falls back to whatever tesseract is on PATH.
"""
import os
import sys
from pathlib import Path


def _bundled_base() -> Path:
    # PyInstaller extracts bundled data under sys._MEIPASS at runtime; from
    # source, the repo root (this file's parent's parent) is the base.
    return Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent.parent))


def _configure():
    """Point pytesseract at the Tesseract bundled in the exe when present.
    Bundled layout: <base>/tesseract/tesseract(.exe) and <base>/tesseract/tessdata."""
    import pytesseract

    tess_dir = _bundled_base() / "tesseract"
    for exe_name in ("tesseract.exe", "tesseract"):
        exe = tess_dir / exe_name
        if exe.exists():
            pytesseract.pytesseract.tesseract_cmd = str(exe)
            tessdata = tess_dir / "tessdata"
            if tessdata.exists():
                os.environ["TESSDATA_PREFIX"] = str(tessdata)
            break
    return pytesseract


def available() -> bool:
    """True if a usable Tesseract (bundled or on PATH) is reachable."""
    try:
        _configure().get_tesseract_version()
        return True
    except Exception:
        return False


def ocr_words(pil_image, scale: float):
    """OCR one page image. Returns (words, line_texts):
    - words: [{x0,x1,top,bottom,text}, ...] in PDF points (scale = 72/dpi),
      the same shape the text-extraction path uses so the shared grid builder
      can reconstruct columns.
    - line_texts: the plain text of each detected line, a guaranteed fallback
      so a scanned page never comes out empty even if the grid is thin."""
    pt = _configure()
    from pytesseract import Output

    data = pt.image_to_data(pil_image, output_type=Output.DICT)
    words = []
    lines: dict = {}
    for i in range(len(data["text"])):
        text = (data["text"][i] or "").strip()
        if not text:
            continue
        left, top = data["left"][i], data["top"][i]
        width, height = data["width"][i], data["height"][i]
        words.append(
            {
                "text": text,
                "x0": left * scale,
                "x1": (left + width) * scale,
                "top": top * scale,
                "bottom": (top + height) * scale,
            }
        )
        key = (data["block_num"][i], data["par_num"][i], data["line_num"][i])
        lines.setdefault(key, []).append(text)
    line_texts = [" ".join(parts) for parts in lines.values()]
    return words, line_texts
