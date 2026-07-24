"""PJS Pdf Studio entry point.

Everything below runs locally: file dialogs are native OS dialogs, all PDF/
Office processing happens with the libraries in core/, and nothing here ever
opens a network socket. Built for Piyush J. Shah & Co., Chartered Accountant.
"""
import base64
import functools
import logging
import multiprocessing
import sys
import tempfile
import traceback
import uuid
from pathlib import Path

import webview

sys.path.insert(0, str(Path(__file__).parent))

from core import convert_from_pdf, convert_to_pdf, legibility, optimize, organize, paths, preview, security

# The packaged app runs with --windowed (no console), so without a log file
# a failure like "nothing happens, no error shown" leaves zero trace to
# debug from. Every Api call's exceptions get written here.
LOG_DIR = paths.default_output_dir().parent / "PJS Logs"
LOG_DIR.mkdir(parents=True, exist_ok=True)
LOG_FILE = LOG_DIR / "pjs-studio.log"
logging.basicConfig(
    filename=str(LOG_FILE),
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
)
logger = logging.getLogger("pjs")


def _log_uncaught(exc_type, exc_value, exc_tb):
    logger.error("Uncaught exception:\n%s", "".join(traceback.format_exception(exc_type, exc_value, exc_tb)))


sys.excepthook = _log_uncaught


def _log_errors(fn):
    """Wraps an Api method so any exception is logged to disk and re-raised
    as a plain RuntimeError, which pywebview reliably marshals back to the
    UI's error box as a readable string."""

    @functools.wraps(fn)
    def wrapper(self, *args, **kwargs):
        try:
            return fn(self, *args, **kwargs)
        except Exception as e:
            logger.exception("Api.%s failed (args=%r)", fn.__name__, args)
            raise RuntimeError(str(e) or f"{type(e).__name__} in {fn.__name__}") from None

    return wrapper


def _log_all_methods(cls):
    for name, value in list(vars(cls).items()):
        if callable(value) and not name.startswith("_"):
            setattr(cls, name, _log_errors(value))
    return cls


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


@_log_all_methods
class Api:
    def __init__(self):
        self.window = None  # attached after window creation, needed for dialogs
        # Live progress for long operations (compression). The UI polls
        # get_progress() on a timer while a job runs; pywebview dispatches each
        # call on its own thread, so the worker can update this dict while the
        # poller reads it. A plain dict read/write under the GIL is enough here.
        self._progress = {"active": False, "pct": 0, "label": ""}

    def _set_progress(self, pct, label):
        self._progress = {"active": True, "pct": max(0, min(100, int(pct))), "label": label}

    def _clear_progress(self):
        self._progress = {"active": False, "pct": 100, "label": ""}

    def get_progress(self):
        return dict(self._progress)

    # ---------- drag-and-drop ----------
    def receive_dropped_file(self, name: str, data_b64: str):
        """The browser side can't reliably read a real filesystem path off a
        dropped file in every pywebview/OS combination, so drag-and-drop
        ships the file's actual bytes over the same bridge every other call
        uses instead of guessing at a path. Written to a per-drop temp file
        so the rest of the app can treat it exactly like a dialog-picked file."""
        drop_dir = Path(tempfile.gettempdir()) / "pjs-pdf-studio-drops"
        drop_dir.mkdir(parents=True, exist_ok=True)
        safe_name = Path(name).name  # strip any path component the browser sent
        dest = drop_dir / f"{uuid.uuid4().hex}_{safe_name}"
        dest.write_bytes(base64.b64decode(data_b64))
        info = _file_info(str(dest))
        info["name"] = safe_name  # show the real filename, not the uuid-prefixed temp one
        return info

    # ---------- password-protected input files ----------
    def is_encrypted(self, file_path):
        return security.is_encrypted(file_path)

    def unlock(self, file_path, password):
        """Decrypt a locked PDF into a temp file and return it as a normal
        file-info dict, so every downstream tool works on it without any
        password handling of its own."""
        decrypted = security.decrypt_to_temp(file_path, password)
        info = _file_info(decrypted)
        info["name"] = _file_info(file_path)["name"]  # keep the user's real filename
        return info

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
        ext = Path(suggested_name).suffix  # e.g. ".xlsx"
        label = FILE_TYPE_LABELS.get(ext.lower(), "All files (*.*)")
        result = self.window.create_file_dialog(
            webview.SAVE_DIALOG,
            directory=str(paths.default_output_dir()),
            save_filename=suggested_name,
            file_types=(label, "All files (*.*)"),
        )
        if not result:
            return None
        path = result if isinstance(result, str) else result[0]
        # The native save dialog can hand back a path with no extension (its
        # "Save as type" was All Files), which then saves e.g. an .xlsx with no
        # extension and won't open in Excel. Force the intended extension on.
        if ext and not path.lower().endswith(ext.lower()):
            path += ext
        return path

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

    def split(self, file_path, save_dir, ranges=None, merge=False):
        outputs = organize.split_pdf(file_path, save_dir, ranges, merge)
        return {"ok": True, "outputs": outputs}

    def rotate(self, file_path, save_path, rotations):
        organize.rotate_pdf(file_path, save_path, rotations)
        return {"ok": True}

    def remove_pages(self, file_path, page_numbers, save_path):
        organize.remove_pages(file_path, page_numbers, save_path)
        return {"ok": True}

    def page_count(self, file_path):
        return preview.page_count(file_path)

    # ---------- optimize ----------
    def compress(self, file_path, level, save_path):
        self._set_progress(0, "Compressing " + Path(file_path).name)
        try:
            return optimize.compress_pdf(
                file_path, level, save_path,
                progress=lambda f: self._set_progress(f * 100, "Compressing " + Path(file_path).name),
            )
        finally:
            self._clear_progress()

    def compress_custom(self, file_path, target_pct, save_path):
        # The custom ladder runs several compression passes with OCR checks, so
        # a true per-page percent isn't meaningful; show a moving indeterminate
        # state instead of a stuck 0.
        self._set_progress(5, "Compressing " + Path(file_path).name + " (custom target)")
        try:
            return optimize.compress_pdf_custom(file_path, target_pct, save_path)
        finally:
            self._clear_progress()

    def compress_batch(self, file_paths, level, save_dir, target_pct=None, prefix=""):
        """Compress several PDFs in one run, each written into save_dir as
        <prefix>_<name>_compressed.pdf (prefix optional, capped at 4 chars).
        One file failing (e.g. a corrupt PDF) doesn't abort the rest: its error
        is captured and the batch continues, so a long run of statements never
        loses the files that did compress."""
        prefix = (prefix or "").strip()[:4]
        results = []
        n = len(file_paths) or 1
        try:
            for i, fp in enumerate(file_paths):
                name = Path(fp).name
                base = (f"{prefix}_" if prefix else "") + Path(fp).stem + "_compressed"
                out = Path(save_dir) / f"{base}.pdf"
                dup = 2
                while out.exists():
                    out = Path(save_dir) / f"{base}-{dup}.pdf"
                    dup += 1

                def on_file_progress(frac, i=i, name=name):
                    self._set_progress((i + frac) / n * 100, f"File {i + 1} of {n}: {name}")

                try:
                    if level == "custom":
                        self._set_progress((i + 0.05) / n * 100, f"File {i + 1} of {n}: {name}")
                        res = optimize.compress_pdf_custom(fp, target_pct, str(out))
                    else:
                        res = optimize.compress_pdf(fp, level, str(out), progress=on_file_progress)
                    results.append(
                        {
                            "name": name,
                            "ok": True,
                            "output": str(out),
                            "before_bytes": res["before_bytes"],
                            "after_bytes": res["after_bytes"],
                            "achieved_pct": res.get("achieved_pct"),
                            "reason": res.get("reason"),
                        }
                    )
                except Exception as e:
                    logger.exception("compress_batch item failed: %s", fp)
                    results.append({"name": name, "ok": False, "error": str(e) or type(e).__name__})
        finally:
            self._clear_progress()
        return {"results": results, "save_dir": save_dir}

    def ocr_available(self):
        return legibility.is_available()

    def watermark(self, file_path, text, save_path, opacity=0.25, font_size=48):
        optimize.watermark_pdf(file_path, text, save_path, opacity, font_size)
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
    # Required for multiprocessing (used to isolate a crash-prone PyMuPDF
    # call in compress_pdf) to work correctly in a frozen PyInstaller exe on
    # Windows, without this, spawning a child process can re-launch the
    # whole app instead of running the worker function.
    multiprocessing.freeze_support()
    main()
