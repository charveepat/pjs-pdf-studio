"""Images -> PDF (pure Python, any OS) and Word/Excel/PowerPoint -> PDF
(Windows + Microsoft Office only, via COM automation against whatever
version of Office is already installed on the machine)."""
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


def word_to_pdf(path: str, save_path: str) -> None:
    _require_windows_office()
    import win32com.client as win32  # noqa: local import, Windows-only dependency

    word = win32.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    try:
        doc = word.Documents.Open(os.path.abspath(path))
        doc.SaveAs(os.path.abspath(save_path), FileFormat=17)  # wdFormatPDF
        doc.Close(False)
    finally:
        word.Quit()


def excel_to_pdf(path: str, save_path: str) -> None:
    _require_windows_office()
    import win32com.client as win32

    excel = win32.gencache.EnsureDispatch("Excel.Application")
    excel.Visible = False
    try:
        wb = excel.Workbooks.Open(os.path.abspath(path))
        wb.ExportAsFixedFormat(0, os.path.abspath(save_path))  # 0 = xlTypePDF
        wb.Close(False)
    finally:
        excel.Quit()


def ppt_to_pdf(path: str, save_path: str) -> None:
    _require_windows_office()
    import win32com.client as win32

    powerpoint = win32.gencache.EnsureDispatch("PowerPoint.Application")
    try:
        pres = powerpoint.Presentations.Open(os.path.abspath(path), WithWindow=False)
        pres.SaveAs(os.path.abspath(save_path), 32)  # 32 = ppSaveAsPDF
        pres.Close()
    finally:
        powerpoint.Quit()
