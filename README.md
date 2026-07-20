# PJS Pdf Studio

A local, offline PDF/Office toolkit for Piyush J. Shah & Co. Every operation
(merge, split, compress, redact, password-protect, Office↔PDF conversion)
runs on-device with PyMuPDF, pdf2docx, pdfplumber and Microsoft Office COM
automation — nothing is ever uploaded anywhere.

## Project layout

- `main.py` — pywebview entry point; exposes a Python `Api` class to the UI
- `ui/index.html` — the entire UI (self-contained HTML/CSS/JS)
- `core/` — the real engine: `organize.py`, `optimize.py`, `security.py`,
  `convert_to_pdf.py`, `convert_from_pdf.py`, `preview.py`
- `.github/workflows/build-windows.yml` — builds `PJS Pdf Studio.exe` on a
  free GitHub-hosted Windows runner every time `main` is pushed

## Running from source (any OS except Office↔PDF conversion, which is Windows-only)

```
python3 -m venv .venv
source .venv/bin/activate      # Windows: .venv\Scripts\activate
pip install -r requirements.txt
python main.py
```

## Getting the .exe

Push to `main` (or run the workflow manually from the Actions tab) and
download `PJS-Pdf-Studio-windows` from the finished run's Artifacts section —
that's `PJS Pdf Studio.exe`, ready to copy to any office computer.

## Known limitation

GitHub's Windows runners don't have Microsoft Office installed, so Word/Excel/
PowerPoint → PDF (which drives the real Office apps via COM automation) can't
be exercised in CI — only that the exe itself builds and launches. Those three
tools need a real run on an office PC with Office installed before you trust
them for client work.
