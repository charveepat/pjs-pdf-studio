"""Merge, split, rotate, and remove pages — pure PyMuPDF, no external dependency."""
from pathlib import Path

import fitz


def merge_pdfs(paths: list[str], save_path: str) -> None:
    out = fitz.open()
    for p in paths:
        with fitz.open(p) as src:
            out.insert_pdf(src)
    out.save(save_path)
    out.close()


def split_pdf(path: str, save_dir: str, ranges: list[tuple[int, int]] | None = None) -> list[str]:
    """ranges is a list of 1-indexed (start, end) page pairs, inclusive.
    Defaults to one output file per page."""
    doc = fitz.open(path)
    base = Path(path).stem
    if ranges is None:
        ranges = [(i, i) for i in range(1, len(doc) + 1)]
    outputs = []
    for start, end in ranges:
        part = fitz.open()
        part.insert_pdf(doc, from_page=start - 1, to_page=end - 1)
        label = f"p{start}" if start == end else f"p{start}-{end}"
        out_path = str(Path(save_dir) / f"{base}-{label}.pdf")
        part.save(out_path)
        part.close()
        outputs.append(out_path)
    doc.close()
    return outputs


def rotate_pdf(path: str, degrees: int, save_path: str, pages: list[int] | None = None) -> None:
    """degrees must be a multiple of 90. pages is 1-indexed; None means every page."""
    doc = fitz.open(path)
    target = pages if pages else range(1, len(doc) + 1)
    for pno in target:
        page = doc[pno - 1]
        page.set_rotation((page.rotation + degrees) % 360)
    doc.save(save_path)
    doc.close()


def remove_pages(path: str, pages_to_remove: list[int], save_path: str) -> None:
    """pages_to_remove is 1-indexed."""
    doc = fitz.open(path)
    doc.delete_pages([p - 1 for p in pages_to_remove])
    doc.save(save_path)
    doc.close()
