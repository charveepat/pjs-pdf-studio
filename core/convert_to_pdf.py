"""Images -> PDF (pure Python, any OS) and Word/Excel/PowerPoint -> PDF
(Windows + Microsoft Office only, via COM automation against whatever
version of Office is already installed on the machine).

Two easy-to-hit COM traps, both handled below:
  1. pywebview calls each Api method on its own worker thread, and COM must
     be explicitly initialized on every thread that touches it, otherwise
     Office throws a generic "Exception occurred" with no useful detail.
  2. win32com.client.gencache.EnsureDispatch() writes generated wrapper code
     to a cache directory that often isn't writable (or doesn't persist)
     inside a frozen PyInstaller exe. Dispatch() (late binding) sidesteps
     that cache entirely, at the cost of not exposing named constants, so
     every Office enum value below is passed as a raw integer instead.
"""
import os
import sys

from PIL import Image


def images_to_pdf(paths: list[str], save_path: str) -> None:
    images = [Image.open(p).convert("RGB") for p in paths]
    first, rest = images[0], images[1:]
    first.save(save_path, save_all=True, append_images=rest)


def _require_windows_office():
    if sys.platform != "win32":
        raise RuntimeError(
            "Converting Office files to PDF uses Microsoft Office and only works on Windows."
        )


class _ComSession:
    """Initializes COM on the current thread for the life of one Office
    automation call, and guarantees CoUninitialize runs even on error."""

    def __enter__(self):
        import pythoncom

        pythoncom.CoInitialize()
        return self

    def __exit__(self, *exc_info):
        import pythoncom

        pythoncom.CoUninitialize()


def word_to_pdf(path: str, save_path: str) -> None:
    _require_windows_office()
    import win32com.client as win32

    with _ComSession():
        word = win32.Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0  # wdAlertsNone
        try:
            doc = word.Documents.Open(os.path.abspath(path), ReadOnly=True)
            doc.SaveAs(os.path.abspath(save_path), FileFormat=17)  # wdFormatPDF
            doc.Close(False)
        finally:
            word.Quit()


def excel_to_pdf(path: str, save_path: str) -> None:
    _require_windows_office()
    import win32com.client as win32

    with _ComSession():
        excel = win32.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        try:
            excel.AskToUpdateLinks = False
        except Exception:
            pass  # not exposed on every Excel version; safe to skip
        try:
            wb = excel.Workbooks.Open(os.path.abspath(path), ReadOnly=True)
            wb.ExportAsFixedFormat(0, os.path.abspath(save_path))  # 0 = xlTypePDF
            wb.Close(False)
        finally:
            excel.Quit()


def ppt_to_pdf(path: str, save_path: str) -> None:
    _require_windows_office()
    import win32com.client as win32

    with _ComSession():
        powerpoint = win32.Dispatch("PowerPoint.Application")
        powerpoint.DisplayAlerts = 2  # ppAlertsNone
        try:
            pres = powerpoint.Presentations.Open(
                os.path.abspath(path), WithWindow=False, ReadOnly=True
            )
            pres.SaveAs(os.path.abspath(save_path), 32)  # 32 = ppSaveAsPDF
            pres.Close()
        finally:
            powerpoint.Quit()
