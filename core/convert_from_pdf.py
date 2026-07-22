"""PDF -> Word / Excel / PowerPoint / Images. All pure Python, no Office needed."""
import tempfile
from pathlib import Path

import fitz
from openpyxl import Workbook
from pdf2docx import Converter
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR
from pptx.util import Emu, Pt

EMU_PER_POINT = 12700  # exact: 914400 EMU/inch / 72 points/inch
PDF_BOLD_FLAG = 1 << 4
PDF_ITALIC_FLAG = 1 << 1
PUA_START, PUA_END = 0xE000, 0xF8FF  # Unicode Private Use Area


def _sanitize_text(text: str) -> str:
    """Bullet-list markers in PowerPoint-exported PDFs are usually a
    character from a symbol font (Wingdings, Webdings, etc.) rendered via
    that font's own private glyph mapping — e.g. Wingdings code point
    U+F0D8 is a right-pointing arrow. Extracted as plain text and dropped
    into a normal font, those code points have no defined glyph at all and
    render as broken boxes throughout the deck (this was the "huge issue").
    Since we don't carry the original font family over (see _add_text_block),
    there's no font install that would make the real glyph show up either —
    swapping in a plain bullet character preserves what the mark actually
    meant instead of showing a hole in the page."""
    return "".join("•" if PUA_START <= ord(c) <= PUA_END else c for c in text)


def pdf_to_word(path: str, save_path: str) -> None:
    cv = Converter(path)
    try:
        cv.convert(save_path)
    finally:
        cv.close()


def pdf_to_excel(path: str, save_path: str) -> dict:
    import pdfplumber

    wb = Workbook()
    wb.remove(wb.active)
    table_count = 0
    with pdfplumber.open(path) as pdf:
        for i, page in enumerate(pdf.pages):
            for t_idx, table in enumerate(page.extract_tables()):
                table_count += 1
                name = f"Page{i + 1}" if t_idx == 0 else f"Page{i + 1}_{t_idx + 1}"
                ws = wb.create_sheet(title=name[:31])
                for row in table:
                    ws.append(["" if c is None else c for c in row])
    if not wb.sheetnames:
        wb.create_sheet(title="Sheet1")["A1"] = "No tables were detected in this PDF."
    wb.save(save_path)
    return {"tables_found": table_count}


def _visual_lines(block):
    """PyMuPDF's own "lines" grouping is driven by the PDF's text-showing
    operators, not by what's visually on the same row. PDFs that position
    every word with its own explicit placement (PowerPoint's PDF export does
    this routinely) come out of get_text("dict") as one "line" per *word* —
    same y0/y1, but each its own entry. Re-cluster spans by vertical bbox
    overlap instead, so a sentence PyMuPDF fragmented into eight one-word
    lines becomes the single flowing line it actually is on the page."""
    raw = [ln for ln in block.get("lines", []) if ln.get("spans")]
    clusters = []
    for ln in raw:
        y0, y1 = ln["bbox"][1], ln["bbox"][3]
        target = None
        for c in clusters:
            overlap = min(y1, c["y1"]) - max(y0, c["y0"])
            if overlap > 0.5 * min(y1 - y0, c["y1"] - c["y0"]):
                target = c
                break
        if target is None:
            clusters.append({"y0": y0, "y1": y1, "spans": list(ln["spans"])})
        else:
            target["spans"].extend(ln["spans"])
            target["y0"] = min(target["y0"], y0)
            target["y1"] = max(target["y1"], y1)
    for c in clusters:
        c["spans"].sort(key=lambda s: s["bbox"][0])  # left-to-right reading order
    clusters.sort(key=lambda c: c["y0"])
    return clusters


def _add_text_block(slide, block, page_height_pt):
    lines = _visual_lines(block)
    if not lines:
        return
    x0, y0, x1, y1 = block["bbox"]
    # PDF text bboxes hug the glyphs tightly; give the textbox a little
    # breathing room so descenders/ascenders don't get clipped in PowerPoint.
    pad_pt = 2
    box = slide.shapes.add_textbox(
        Emu(int((x0 - pad_pt) * EMU_PER_POINT)),
        Emu(int((y0 - pad_pt) * EMU_PER_POINT)),
        Emu(int((x1 - x0 + 2 * pad_pt) * EMU_PER_POINT)),
        Emu(int((y1 - y0 + 2 * pad_pt) * EMU_PER_POINT)),
    )
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.TOP
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0

    for i, line in enumerate(lines):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        spans = line["spans"]
        for j, span in enumerate(spans):
            text = _sanitize_text(span.get("text", ""))
            if not text:
                continue
            # Independently-placed words carry no space of their own — infer
            # one from the real gap between this word and the previous.
            if j > 0:
                gap = span["bbox"][0] - spans[j - 1]["bbox"][2]
                if gap > 1.0:
                    text = " " + text
            run = para.add_run()
            run.text = text
            size = span.get("size", 12)
            run.font.size = Pt(max(1, round(size)))
            flags = span.get("flags", 0)
            run.font.bold = bool(flags & PDF_BOLD_FLAG)
            run.font.italic = bool(flags & PDF_ITALIC_FLAG)
            color_int = span.get("color", 0)
            r = (color_int >> 16) & 255
            g = (color_int >> 8) & 255
            b = color_int & 255
            run.font.color.rgb = RGBColor(r, g, b)
            # Font *family* isn't preserved — PDF embeds subset fonts under
            # internal names (e.g. "ABCDEE+Calibri-Bold") that rarely match
            # an installed PowerPoint font, so guessing one in would more
            # often replace a correct-looking font with a wrong one.


def _add_page_images(slide, page, doc):
    for xref, *_ in page.get_images(full=True):
        rects = page.get_image_rects(xref)
        if not rects:
            continue
        rect = rects[0]
        try:
            info = doc.extract_image(xref)
        except Exception:
            continue
        ext = info.get("ext", "png")
        tmp = tempfile.NamedTemporaryFile(suffix=f".{ext}", delete=False)
        try:
            tmp.write(info["image"])
            tmp.close()
            slide.shapes.add_picture(
                tmp.name,
                Emu(int(rect.x0 * EMU_PER_POINT)),
                Emu(int(rect.y0 * EMU_PER_POINT)),
                Emu(int(rect.width * EMU_PER_POINT)),
                Emu(int(rect.height * EMU_PER_POINT)),
            )
        finally:
            Path(tmp.name).unlink(missing_ok=True)


def pdf_to_ppt(path: str, save_path: str) -> None:
    """Rebuilds each page as real, editable PowerPoint content: text blocks
    become editable text boxes (position, size, bold/italic, and color
    preserved; font family is not — see _add_text_block), and embedded
    images become separate picture shapes rather than one flattened
    page-sized raster. Works best on simple, single-column layouts; dense
    multi-column pages may need manual tidying after conversion, the same
    trade-off as PDF to Word."""
    doc = fitz.open(path)
    if doc.page_count == 0:
        doc.close()
        raise ValueError("This PDF has no pages.")

    first = doc[0]
    prs = Presentation()
    prs.slide_width = Emu(int(first.rect.width * EMU_PER_POINT))
    prs.slide_height = Emu(int(first.rect.height * EMU_PER_POINT))
    blank_layout = prs.slide_layouts[6]

    for page in doc:
        slide = prs.slides.add_slide(blank_layout)
        _add_page_images(slide, page, doc)
        text_dict = page.get_text("dict")
        for block in text_dict.get("blocks", []):
            if block.get("type") == 0:  # 0 = text, 1 = image (images handled separately above)
                _add_text_block(slide, block, first.rect.height)

    doc.close()
    prs.save(save_path)


def pdf_to_images(path: str, save_dir: str, fmt: str = "png", dpi: int = 150) -> list[str]:
    doc = fitz.open(path)
    base = Path(path).stem
    mat = fitz.Matrix(dpi / 72, dpi / 72)
    outputs = []
    for page in doc:
        pix = page.get_pixmap(matrix=mat)
        out_path = str(Path(save_dir) / f"{base}-page{page.number + 1}.{fmt}")
        pix.save(out_path)
        outputs.append(out_path)
    doc.close()
    return outputs
