"""PJS Pdf Studio entry point.

Everything below runs locally: file dialogs are native OS dialogs, all PDF/
Office processing happens with the libraries in core/, and nothing here ever
opens a network socket. Built for Piyush J. Shah & Co., Chartered Accountant.

Startup performance: only webview, stdlib, and core.paths are imported at
module load time. Every heavy library (PyMuPDF, pdf2docx, pdfplumber,
python-pptx, Pillow, pytesseract, etc.) is imported lazily on the first API
call that needs it. This lets the window appear in roughly 1-2 seconds instead
of waiting for all libraries to initialise before the UI is shown.
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

from core import paths   # lightweight, no heavy deps

# Core modules are imported lazily via _mod() at runtime for fast startup.
# The explicit imports below are NEVER executed (guarded by False) but are
# required so PyInstaller's static analyser sees and bundles every module.
# The build also passes --collect-submodules core as a belt-and-suspenders.
if False:  # pragma: no cover
    from core import (  # noqa: F401
        convert_from_pdf, convert_to_pdf, legibility,
        optimize, organize, ocr, preview, security,
    )

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


# Lazy module cache: each core module is imported once on first use and
# cached here. Access via _mod("optimize") etc.
_modules: dict = {}

def _mod(name: str):
    if name not in _modules:
        import importlib
        _modules[name] = importlib.import_module(f"core.{name}")
    return _modules[name]


def _log_errors(fn):
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
        self.window = None
        self._progress = {"active": False, "pct": 0, "label": ""}

    def _set_progress(self, pct, label):
        self._progress = {"active": True, "pct": max(0, min(100, int(pct))), "label": label}

    def _clear_progress(self):
        self._progress = {"active": False, "pct": 100, "label": ""}

    def get_progress(self):
        return dict(self._progress)

    # ---------- drag-and-drop ----------
    def receive_dropped_file(self, name: str, data_b64: str):
        drop_dir = Path(tempfile.gettempdir()) / "pjs-pdf-studio-drops"
        drop_dir.mkdir(parents=True, exist_ok=True)
        safe_name = Path(name).name
        dest = drop_dir / f"{uuid.uuid4().hex}_{safe_name}"
        dest.write_bytes(base64.b64decode(data_b64))
        info = _file_info(str(dest))
        info["name"] = safe_name
        return info

    # ---------- password-protected input files ----------
    def is_encrypted(self, file_path):
        return _mod("security").is_encrypted(file_path)

    def unlock(self, file_path, password):
        decrypted = _mod("security").decrypt_to_temp(file_path, password)
        info = _file_info(decrypted)
        info["name"] = _file_info(file_path)["name"]
        return info

    # ---------- file / save dialogs ----------
    def pick_open_file(self, accept: str = ""):
        result = self.window.create_file_dialog(webview.OPEN_DIALOG, file_types=_file_types(accept))
        return _file_info(result[0]) if result else None

    def pick_open_files(self, accept: str = ""):
        result = self.window.create_file_dialog(
            webview.OPEN_DIALOG, allow_multiple=True, file_types=_file_types(accept)
        )
        return [_file_info(p) for p in result] if result else []

    def pick_save_path(self, suggested_name: str):
        ext = Path(suggested_name).suffix
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

    def pick_input_folder(self):
        result = self.window.create_file_dialog(webview.FOLDER_DIALOG)
        if not result:
            return None
        return result[0] if isinstance(result, (list, tuple)) else result

    def scan_folder_for_pdfs(self, folder_path: str):
        """Return info dicts for every PDF in folder_path and all subfolders.
        Case-insensitive on extension so .PDF / .Pdf / .pdf are all caught.
        Each result includes a 'rel' key with the path relative to folder_path
        so the UI can show subfolder context and the batch can mirror the
        folder structure in the output directory."""
        folder = Path(folder_path)
        pdfs = []
        for p in folder.rglob("*"):
            if p.is_file() and p.suffix.lower() == ".pdf":
                info = _file_info(str(p))
                info["rel"] = str(p.relative_to(folder))
                pdfs.append(info)
        pdfs.sort(key=lambda x: x["rel"].lower())
        return pdfs

    # ---------- organize ----------
    def merge(self, file_paths, save_path):
        _mod("organize").merge_pdfs(file_paths, save_path)
        return {"ok": True}

    def merge_pages(self, items, save_path):
        _mod("organize").merge_pages(items, save_path)
        return {"ok": True}

    def split(self, file_path, save_dir, ranges=None, merge=False):
        outputs = _mod("organize").split_pdf(file_path, save_dir, ranges, merge)
        return {"ok": True, "outputs": outputs}

    def rotate(self, file_path, save_path, rotations):
        _mod("organize").rotate_pdf(file_path, save_path, rotations)
        return {"ok": True}

    def remove_pages(self, file_path, page_numbers, save_path):
        _mod("organize").remove_pages(file_path, page_numbers, save_path)
        return {"ok": True}

    def page_count(self, file_path):
        return _mod("preview").page_count(file_path)

    # ---------- optimize ----------
    def compress(self, file_path, level, save_path):
        self._set_progress(0, "Compressing " + Path(file_path).name)
        try:
            return _mod("optimize").compress_pdf(
                file_path, level, save_path,
                progress=lambda f: self._set_progress(f * 100, "Compressing " + Path(file_path).name),
            )
        finally:
            self._clear_progress()

    def compress_custom(self, file_path, target_pct, save_path):
        self._set_progress(5, "Compressing " + Path(file_path).name + " (custom target)")
        try:
            return _mod("optimize").compress_pdf_custom(file_path, target_pct, save_path)
        finally:
            self._clear_progress()

    def compress_batch(self, file_paths, level, save_dir, target_pct=None, prefix="", rel_paths=None):
        """Compress a list of PDFs into save_dir.
        If rel_paths is provided (list of paths relative to the source folder),
        the subfolder structure is mirrored inside save_dir so files from
        different subfolders never collide."""
        optimize = _mod("optimize")
        prefix = (prefix or "").strip()[:4]
        results = []
        n = len(file_paths) or 1
        try:
            for i, fp in enumerate(file_paths):
                name = Path(fp).name
                rel = Path(rel_paths[i]) if rel_paths and i < len(rel_paths) else Path(name)
                base = (f"{prefix}_" if prefix else "") + rel.stem + "_compressed"
                out_dir = Path(save_dir) / rel.parent
                out_dir.mkdir(parents=True, exist_ok=True)
                out = out_dir / f"{base}.pdf"
                dup = 2
                while out.exists():
                    out = out_dir / f"{base}-{dup}.pdf"
                    dup += 1

                def on_file_progress(frac, i=i, name=name):
                    self._set_progress((i + frac) / n * 100, f"File {i + 1} of {n}: {name}")

                try:
                    if level == "custom":
                        self._set_progress((i + 0.05) / n * 100, f"File {i + 1} of {n}: {name}")
                        res = optimize.compress_pdf_custom(fp, target_pct, str(out))
                    else:
                        res = optimize.compress_pdf(fp, level, str(out), progress=on_file_progress)
                    results.append({
                        "name": name, "ok": True, "output": str(out),
                        "before_bytes": res["before_bytes"], "after_bytes": res["after_bytes"],
                        "achieved_pct": res.get("achieved_pct"), "reason": res.get("reason"),
                    })
                except Exception as e:
                    logger.exception("compress_batch item failed: %s", fp)
                    results.append({"name": name, "ok": False, "error": str(e) or type(e).__name__})
        finally:
            self._clear_progress()
        return {"results": results, "save_dir": save_dir}

    def ocr_available(self):
        return _mod("legibility").is_available()

    def watermark(self, file_path, text, save_path, opacity=0.25, font_size=48):
        _mod("optimize").watermark_pdf(file_path, text, save_path, opacity, font_size)
        return {"ok": True}

    # ---------- security ----------
    def scan_sensitive(self, file_path, pattern_keys):
        return _mod("security").scan_sensitive(file_path, pattern_keys)

    def render_page_preview(self, file_path, page_num, max_width=520):
        return _mod("preview").render_page(file_path, page_num, max_width)

    def image_thumbnail(self, file_path, max_width=240):
        import io
        from PIL import Image
        img = Image.open(file_path)
        if img.mode not in ("RGB", "L"):
            img = img.convert("RGB")
        w, h = img.size
        if w > max_width:
            img = img.resize((max_width, max(1, round(h * max_width / w))), Image.LANCZOS)
        buf = io.BytesIO()
        img.save(buf, format="JPEG", quality=80)
        return {
            "image_b64": base64.b64encode(buf.getvalue()).decode("ascii"),
            "width": img.width,
            "height": img.height,
        }

    def redact(self, file_path, boxes, save_path):
        _mod("security").redact_pdf(file_path, boxes, save_path)
        return {"ok": True}

    def protect(self, file_path, password, save_path):
        _mod("security").password_protect(file_path, password, save_path)
        return {"ok": True}

    # ---------- convert to PDF ----------
    def word_to_pdf(self, file_path, save_path):
        _mod("convert_to_pdf").word_to_pdf(file_path, save_path)
        return {"ok": True}

    def excel_to_pdf(self, file_path, save_path):
        _mod("convert_to_pdf").excel_to_pdf(file_path, save_path)
        return {"ok": True}

    def ppt_to_pdf(self, file_path, save_path):
        _mod("convert_to_pdf").ppt_to_pdf(file_path, save_path)
        return {"ok": True}

    def images_to_pdf(self, file_paths, save_path):
        _mod("convert_to_pdf").images_to_pdf(file_paths, save_path)
        return {"ok": True}

    # ---------- convert from PDF ----------
    def pdf_to_word(self, file_path, save_path):
        _mod("convert_from_pdf").pdf_to_word(file_path, save_path)
        return {"ok": True}

    def pdf_to_excel(self, file_path, save_path):
        return _mod("convert_from_pdf").pdf_to_excel(file_path, save_path)

    def pdf_to_ppt(self, file_path, save_path):
        _mod("convert_from_pdf").pdf_to_ppt(file_path, save_path)
        return {"ok": True}

    def pdf_to_images(self, file_path, save_dir):
        outputs = _mod("convert_from_pdf").pdf_to_images(file_path, save_dir)
        return {"ok": True, "outputs": outputs}

    def batch_convert(self, file_paths: list, kind: str, save_dir: str):
        cfp = _mod("convert_from_pdf")
        ctp = _mod("convert_to_pdf")
        results = []
        n = len(file_paths) or 1
        ext_map = {
            "word_to_pdf": ".pdf", "excel_to_pdf": ".pdf", "ppt_to_pdf": ".pdf",
            "pdf_to_word": ".docx", "pdf_to_ppt": ".pptx", "pdf_to_excel": ".xlsx",
            "pdf_to_images": None,
        }
        out_ext = ext_map.get(kind)
        try:
            for i, fp in enumerate(file_paths):
                name = Path(fp).name
                self._set_progress((i / n) * 100, f"File {i+1} of {n}: {name}")
                try:
                    if out_ext:
                        out = Path(save_dir) / (Path(fp).stem + out_ext)
                        dup = 2
                        while out.exists():
                            out = Path(save_dir) / f"{Path(fp).stem}-{dup}{out_ext}"
                            dup += 1
                        if kind == "word_to_pdf":   ctp.word_to_pdf(fp, str(out))
                        elif kind == "excel_to_pdf": ctp.excel_to_pdf(fp, str(out))
                        elif kind == "ppt_to_pdf":   ctp.ppt_to_pdf(fp, str(out))
                        elif kind == "pdf_to_word":  cfp.pdf_to_word(fp, str(out))
                        elif kind == "pdf_to_ppt":   cfp.pdf_to_ppt(fp, str(out))
                        elif kind == "pdf_to_excel": cfp.pdf_to_excel(fp, str(out))
                        results.append({"name": name, "ok": True, "output": str(out)})
                    else:
                        sub = Path(save_dir) / Path(fp).stem
                        sub.mkdir(parents=True, exist_ok=True)
                        outs = cfp.pdf_to_images(fp, str(sub))
                        results.append({"name": name, "ok": True, "output": str(sub), "count": len(outs)})
                except Exception as e:
                    logger.exception("batch_convert item failed: %s", fp)
                    results.append({"name": name, "ok": False, "error": str(e) or type(e).__name__})
        finally:
            self._clear_progress()
        return {"results": results, "save_dir": save_dir}

    def batch_watermark(self, file_paths: list, text: str, save_dir: str, opacity: float = 0.25, font_size: int = 48):
        optimize = _mod("optimize")
        results = []
        n = len(file_paths) or 1
        try:
            for i, fp in enumerate(file_paths):
                name = Path(fp).name
                self._set_progress((i / n) * 100, f"File {i+1} of {n}: {name}")
                try:
                    out = Path(save_dir) / (Path(fp).stem + "-watermarked.pdf")
                    dup = 2
                    while out.exists():
                        out = Path(save_dir) / f"{Path(fp).stem}-watermarked-{dup}.pdf"
                        dup += 1
                    optimize.watermark_pdf(fp, text, str(out), opacity, font_size)
                    results.append({"name": name, "ok": True, "output": str(out)})
                except Exception as e:
                    logger.exception("batch_watermark item failed: %s", fp)
                    results.append({"name": name, "ok": False, "error": str(e) or type(e).__name__})
        finally:
            self._clear_progress()
        return {"results": results, "save_dir": save_dir}

    def batch_protect(self, file_paths: list, password: str, save_dir: str):
        security = _mod("security")
        results = []
        n = len(file_paths) or 1
        try:
            for i, fp in enumerate(file_paths):
                name = Path(fp).name
                self._set_progress((i / n) * 100, f"File {i+1} of {n}: {name}")
                try:
                    out = Path(save_dir) / (Path(fp).stem + "-protected.pdf")
                    dup = 2
                    while out.exists():
                        out = Path(save_dir) / f"{Path(fp).stem}-protected-{dup}.pdf"
                        dup += 1
                    security.password_protect(fp, password, str(out))
                    results.append({"name": name, "ok": True, "output": str(out)})
                except Exception as e:
                    logger.exception("batch_protect item failed: %s", fp)
                    results.append({"name": name, "ok": False, "error": str(e) or type(e).__name__})
        finally:
            self._clear_progress()
        return {"results": results, "save_dir": save_dir}


def resource_path(relative: str) -> Path:
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
    # Force Edge WebView2 backend (pre-installed on all Win10 1803+ / Win11).
    # The WinForms backend requires pythonnet/.NET and fails on some machines.
    webview.start(gui="edgechromium")


if __name__ == "__main__":
    multiprocessing.freeze_support()
    main()
