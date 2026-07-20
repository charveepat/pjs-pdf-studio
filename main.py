"""PJS Pdf Studio entry point.

Everything below runs locally: file dialogs are native OS dialogs, all PDF/
Office processing happens with the libraries in core/, and nothing here ever
opens a network socket. Built for Piyush J. Shah & Co., Chartered Accountant.
"""
import sys
from pathlib import Path

import webview

sys.path.insert(0, str(Path(__file__).parent))

from core import convert_from_pdf, convert_to_pdf, optimize, organize, paths, preview, security

FILE_TYPE_LABELS = {
    ".pdf": "PDF Files (*.pdf)",
    ".docx": "Word Documents (*.docx;*.doc)",
    ".doc": "Word Documents (*.docx;*.doc)",
    ".xlsx": "Excel Workbooks (*.xlsx;*.xls)",
    ".xls": "Excel Workbooks (*.xlsx;*.xls)",
    ".pptx": "PowerPoint Files (*.pptx;*.ppt)",
    ".ppt": "PowerPoint Files (*.pptx;*.ppt)",
    ".jpg": "Images (*.jpg;*.jpeg;*.png)",
    ".jpeg": "Images (*.jpg;*.jpeg;*.png)",
    ".png": "Images (*.jpg;*.jpeg;*.png)",
}


def _file_types(accept_csv: str):
    first_ext = accept_csv.split(",")[0].strip()
    label = FILE_TYPE_LABELS.get(first_ext, "All files (*.*)")
    return (label, "All files (*.*)")


def _file_info(path: str) -> dict:
    p = Path(path)
    return {"path": str(p), "name": p.name, "size": p.stat().st_size}


class Api:
    def __init__(self):
        self.window = None  # attached after window creation, needed for dialogs

    # ---------- file / save dialogs (native OS dialogs, no fake modal) ----------
    def pick_open_file(self, accept: str = ""):
        result = self.window.create_file_dialog(webview.OPEN_DIALOG, file_types=_file_types(accept))
        return _file_info(result[0]) if result else None

    def pick_open_files(self, accept: str = ""):
        result = self.window.create_file_dialog(
            webview.OPEN_DIALOG, allow_multiple=True, file_types=_file_types(accept)
        )
        return [_file_info(p) for p in result] if result else []

    def pick_save_path(self, suggested_name: str):
        result = self.window.create_file_dialog(
            webview.SAVE_DIALOG,
            directory=str(paths.default_output_dir()),
            save_filename=suggested_name,
        )
        if not result:
            return None
        return result if isinstance(result, str) else result[0]

    def pick_save_dir(self):
        result = self.window.create_file_dialog(
            webview.FOLDER_DIALOG, directory=str(paths.default_output_dir())
        )
        if not result:
            return None
        return result[0] if isinstance(result, (list, tuple)) else result

    # ---------- organize ----------
    def merge(self, file_paths, save_path):
        organize.merge_pdfs(file_paths, save_path)
        return {"ok": True}

    def split(self, file_path, save_dir, ranges=None):
        outputs = organize.split_pdf(file_path, save_dir, ranges)
        return {"ok": True, "outputs": outputs}

    def rotate(self, file_path, degrees, save_path):
        organize.rotate_pdf(file_path, degrees, save_path)
        return {"ok": True}

    def remove_pages(self, file_path, page_numbers, save_path):
        organize.remove_pages(file_path, page_numbers, save_path)
        return {"ok": True}

    def page_count(self, file_path):
        return preview.page_count(file_path)

    # ---------- optimize ----------
    def compress(self, file_path, level, save_path):
        return optimize.compress_pdf(file_path, level, save_path)

    def watermark(self, file_path, text, save_path):
        optimize.watermark_pdf(file_path, text, save_path)
        return {"ok": True}

    # ---------- security ----------
    def scan_sensitive(self, file_path, pattern_keys):
        return security.scan_sensitive(file_path, pattern_keys)

    def render_page_preview(self, file_path, page_num, max_width=520):
        return preview.render_page(file_path, page_num, max_width)

    def redact(self, file_path, boxes, save_path):
        security.redact_pdf(file_path, boxes, save_path)
        return {"ok": True}

    def protect(self, file_path, password, save_path):
        security.password_protect(file_path, password, save_path)
        return {"ok": True}

    # ---------- convert to PDF ----------
    def word_to_pdf(self, file_path, save_path):
        convert_to_pdf.word_to_pdf(file_path, save_path)
        return {"ok": True}

    def excel_to_pdf(self, file_path, save_path):
        convert_to_pdf.excel_to_pdf(file_path, save_path)
        return {"ok": True}

    def ppt_to_pdf(self, file_path, save_path):
        convert_to_pdf.ppt_to_pdf(file_path, save_path)
        return {"ok": True}

    def images_to_pdf(self, file_paths, save_path):
        convert_to_pdf.images_to_pdf(file_paths, save_path)
        return {"ok": True}

    # ---------- convert from PDF ----------
    def pdf_to_word(self, file_path, save_path):
        convert_from_pdf.pdf_to_word(file_path, save_path)
        return {"ok": True}

    def pdf_to_excel(self, file_path, save_path):
        return convert_from_pdf.pdf_to_excel(file_path, save_path)

    def pdf_to_ppt(self, file_path, save_path):
        convert_from_pdf.pdf_to_ppt(file_path, save_path)
        return {"ok": True}

    def pdf_to_images(self, file_path, save_dir):
        outputs = convert_from_pdf.pdf_to_images(file_path, save_dir)
        return {"ok": True, "outputs": outputs}


def resource_path(relative: str) -> Path:
    """Files bundled via PyInstaller's --add-data extract to sys._MEIPASS at
    runtime; when running from source, fall back to this file's directory."""
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).parent))
    return base / relative


def main():
    api = Api()
    window = webview.create_window(
        "PJS Pdf Studio",
        str(resource_path("ui/index.html")),
        js_api=api,
        width=1180,
        height=800,
        min_size=(940, 660),
    )
    api.window = window
    webview.start()


if __name__ == "__main__":
    main()
