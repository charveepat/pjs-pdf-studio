"""Merge, split, rotate, and remove pages, pure PyMuPDF, no external dependency."""
from pathlib import Path

import fitz


def merge_pdfs(paths: list[str], save_path: str) -> None:
    out = fitz.open()
    for p in paths:
        with fitz.open(p) as src:
            out.insert_pdf(src)
    out.save(save_path)
    out.close()


def split_pdf(
    path: str,
    save_dir: str,
    ranges: list[tuple[int, int]] | None = None,
    merge: bool = False,
) -> list[str]:
    """ranges is a list of 1-indexed (start, end) page pairs, inclusive, a
    single page is just (n, n). Defaults to one output file per page.

    merge=False (default): each range becomes its own output file.
    merge=True: every range is combined, in the given order, into one file,
    this is how "pick individual pages visually" becomes a single export:
    the caller passes each picked page as its own (n, n) range with
    merge=True."""
    doc = fitz.open(path)
    base = Path(path).stem
    if ranges is None:
        ranges = [(i, i) for i in range(1, len(doc) + 1)]

    if merge:
        out = fitz.open()
        for start, end in ranges:
            out.insert_pdf(doc, from_page=start - 1, to_page=end - 1)
        out_path = str(Path(save_dir) / f"{base}-selected.pdf")
        out.save(out_path)
        out.close()
        doc.close()
        return [out_path]

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


def rotate_pdf(path: str, save_path: str, rotations: dict[int, int]) -> None:
    """rotations maps 1-indexed page number -> degrees to add (multiple of
    90, can be negative). Pages not present are left untouched. For a
    whole-document rotation, the caller just passes every page number with
    the same degree value."""
    doc = fitz.open(path)
    for pno, degrees in rotations.items():
        page = doc[int(pno) - 1]
        page.set_rotation((page.rotation + degrees) % 360)
    doc.save(save_path)
    doc.close()


def remove_pages(path: str, pages_to_remove: list[int], save_path: str) -> None:
    """pages_to_remove is 1-indexed."""
    doc = fitz.open(path)
    doc.delete_pages([p - 1 for p in pages_to_remove])
    doc.save(save_path)
    doc.close()
