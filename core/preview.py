"""Page thumbnails for the manual-redaction canvas in the UI."""
import base64

import fitz


def page_count(path: str) -> int:
    doc = fitz.open(path)
    n = doc.page_count
    doc.close()
    return n


def render_page(path: str, page_num: int, max_width: int = 520) -> dict:
    """page_num is 1-indexed. Returns a PNG plus the scale factor (device
    pixels per PDF point) so the UI can map a drawn canvas rectangle back
    to real PDF coordinates for redaction."""
    doc = fitz.open(path)
    page = doc[page_num - 1]
    zoom = max_width / page.rect.width
    pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
    png_bytes = pix.tobytes("png")
    doc.close()
    return {
        "image_b64": base64.b64encode(png_bytes).decode("ascii"),
        "width": pix.width,
        "height": pix.height,
        "scale": zoom,
    }
