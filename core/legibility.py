"""Legibility verification for Custom % compression, via Windows' own
built-in OCR engine (winocr wraps the same Windows.Media.Ocr API behind the
Snipping Tool's "Extract text"). No external OCR binary to bundle - it's
already part of every Windows 10/11 install, which is why this only works
on Windows and degrades gracefully (compress_pdf_custom keeps working
without per-file verification) everywhere else."""
import sys


def is_available() -> bool:
    if sys.platform != "win32":
        return False
    try:
        import winocr  # noqa: F401

        return True
    except Exception:
        return False


def ocr_words(pil_img) -> set[str]:
    import winocr

    result = winocr.recognize_pil_sync(pil_img)
    return {w.lower() for w in result.text.split() if w.strip()}


def word_accuracy(reference_words: set[str], candidate_words: set[str]) -> float:
    """Fraction of reference words that also appear in the OCR'd candidate.
    Order-independent and case-insensitive on purpose - what matters here is
    whether the words are still readable, not whether OCR reconstructs them
    in the original reading order."""
    if not reference_words:
        return 1.0
    return len(reference_words & candidate_words) / len(reference_words)
