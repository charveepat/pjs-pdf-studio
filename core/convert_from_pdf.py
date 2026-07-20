"""PDF -> Word / Excel / PowerPoint / Images. All pure Python, no Office needed."""
import shutil
import tempfile
from pathlib import Path

import fitz
from openpyxl import Workbook
from pdf2docx import Converter
from pptx import Presentation
from pptx.util import Inches


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


def pdf_to_ppt(path: str, save_path: str, dpi: int = 150) -> None:
    """Renders each page to an image and places it full-bleed on a slide.
    Reliable for any PDF; slides are not text-editable, only image-editable."""
    doc = fitz.open(path)
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]
    mat = fitz.Matrix(dpi / 72, dpi / 72)
    tmp_dir = tempfile.mkdtemp(prefix="pjs_ppt_")
    try:
        for page in doc:
            pix = page.get_pixmap(matrix=mat)
            img_path = str(Path(tmp_dir) / f"p{page.number}.png")
            pix.save(img_path)
            slide = prs.slides.add_slide(blank_layout)
            slide.shapes.add_picture(img_path, 0, 0, width=prs.slide_width, height=prs.slide_height)
        prs.save(save_path)
    finally:
        doc.close()
        shutil.rmtree(tmp_dir, ignore_errors=True)


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
