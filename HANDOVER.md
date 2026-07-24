# Handover: PJS Pdf Studio

Local, offline PDF/Office toolkit for Piyush J. Shah & Co., Chartered Accountant. Built with Python + pywebview, packaged as a single Windows `.exe` via GitHub Actions. Repo: `/Users/charveepatel/Claude/pjs-pdf-studio`, GitHub: `charveepat/pjs-pdf-studio` (private).

**Start a new chat with this file open and say what you want changed.** Everything below is already built, tested, and pushed, current as of commit `cd11652`.

## How this project actually works (read this first)

- **Stack**: `main.py` (pywebview entry point, exposes a Python `Api` class) + `core/*.py` (the real engine: PyMuPDF/fitz, pdf2docx, pdfplumber, python-docx/openpyxl/python-pptx, Pillow, pywin32) + `ui/index.html` (the entire UI, one self-contained HTML/CSS/JS file, talks to Python via `window.pywebview.api.*`).
- **Build**: `.github/workflows/build-windows.yml` builds on every push to `main` on a free GitHub-hosted Windows runner, produces `PJS Pdf Studio.exe` as a downloadable Actions artifact. Takes about 2-3 minutes.
- **The push workflow the user prefers**: I commit locally and give her the exact `git push` command; she runs it herself in Terminal (her stated preference, don't push for her without being told to). After she pushes, I check `gh run list` / `gh run view` / `gh api .../artifacts` to confirm the build succeeded, then give her the download path: open the run on GitHub, scroll to Artifacts, download the zip, unzip, run the exe.
- **I test everything locally on Mac first**, using a venv at `.venv` (already set up, has all dependencies installed except the Windows-only ones like pywin32/winocr). I validate PDF output independently with poppler (`pdfinfo`, `pdftoppm`), not just by trusting PyMuPDF's own read-back. I visually render and inspect compression output, not just trust the size numbers, especially anything legibility-sensitive.
- **She's testing on a real Windows PC separately.** Anything Windows-only (Office COM automation, the OCR legibility check) genuinely needs her machine to verify; I can't run those from this Mac.

## Current state: fully shipped and working

All 16 tools work: Merge (with file reordering, see below), Split (range + visual page picker + merge-to-one-file toggle), Rotate (per-page), Remove Pages (visual picker), Word/Excel/PowerPoint to PDF (COM automation), Images to PDF, PDF to Word (pdf2docx), PDF to Excel (adaptive per-page extractor, see below), PDF to PowerPoint (real editable text boxes, not rasterized pages), PDF to Images, Compress (single or batch, see below), Watermark (adjustable size/opacity, live preview), Redact (auto-detect + manual draw), Password Protect.

### Merge: file reordering

The Merge file list (and Images to PDF, which is also order-sensitive) supports reordering before combining: drag a row, or use the per-row Up/Down arrows. Files combine top to bottom. `merge_pdfs` concatenates in list order, so controlling list order is the whole fix (`renderFileList` in `ui/index.html`).

**Planned (Phase 2, not yet built):** a two-stage Merge preview. Stage 1: each added file shown as a card with a first-page thumbnail. Stage 2: click in to a page-level board showing every page of every file as draggable thumbnails, rearrange pages across files, then merge in that exact page order. Needs a new page-level merge (current `merge_pdfs` only concatenates whole files) plus a drag grid; the thumbnail plumbing (`preview.render_page`, `renderPageThumbs`) already exists. Deferred to a second push at the user's request (weekly limit).

### Compress: batch (multiple PDFs at once)

Compress accepts one file or several. One file keeps the original Save-As-with-name flow. Two or more use `Api.compress_batch` (`main.py`): pick an output folder, each file is written as `<prefix>_<name>_compressed.pdf` (prefix is an optional field capped at 4 chars, shown only when more than one file is queued), and one bad file doesn't abort the rest (its error is captured per-file, batch continues). The per-file compression logic is unchanged (same already-validated `compress_pdf` / `compress_pdf_custom` paths).

### Password-protected input files (all tools)

Every file added to any tool passes through `unlockIfNeeded` in `ui/index.html` (folded into `wireSingleFileDrop` / `wireMultiFileDrop`, so it covers all 16 tools). It calls `Api.is_encrypted`; if locked, a styled password modal (`askPassword`) collects the password and `Api.unlock` decrypts to a temp copy (`security.decrypt_to_temp`, `PDF_ENCRYPT_NONE`) that every downstream tool then treats as a normal file. Wrong password re-prompts; cancel drops the file. Only `.pdf` files are checked.

### Compression progress bar

`runWithProgress` in `ui/index.html` shows a real percentage during compression by polling `Api.get_progress` every 300ms. The backend reports progress via a callback threaded through `compress_pdf` into `_rasterize` / `_recompress_images` (per page / per image); `compress_batch` maps each file into an overall `File X of N` percentage. Custom % is coarse (its ladder runs many passes, so a true per-page number isn't meaningful). Requested so the office can watch it work.

### Batch-vs-single compression (was: "batch compresses less")

Reported that a font-heavy file compressed far less inside a batch (4.5MB to 3.1MB) than alone (4.5MB to 1.8MB), both at Recommended. Could NOT reproduce on Mac: single and batch are byte-identical even on font-heavy files, because the batch is literally the same `compress_pdf` in a loop. The only Windows-only mechanism is the font-subsetting helper process (`_subset_fonts_isolated`) timing out or failing to spawn under batch load, which leaves fonts un-shrunk (the ~1.3MB gap). Hardened it: timeout now scales with file size (was a flat 25s), retries once, and logs every skip to `pjs-studio.log` so the cause is visible on her machine. If she still sees a gap at a matched level after this, the log will say why.

### Save-dialog extension enforcement

`Api.pick_save_path` now forces the intended extension onto the returned path (the native Windows Save dialog was handing back paths with no extension, so PDF-to-Excel saved a file Excel wouldn't open). Also sets the "Save as type" filter. Fixes Excel, Word, and PowerPoint saves.

### PDF to Excel: adaptive per-page extractor

The old code called pdfplumber's default `extract_tables()`, which fails badly on bank statements: on a ruled statement it returns one tiny "table" per bordered row (fragmented garbage). `pdf_to_excel` (`core/convert_from_pdf.py`) now picks a strategy per page:
- **Ruled tables** (`_ruled_table`): cluster every vertical + horizontal edge into explicit column/row lines and cut on exactly those. This is the common case for Indian bank statements and groups a transaction whose text wraps onto 3 physical lines back into one row. Validated on the ICICI statement (see test files): 210/210 transactions, all 9 columns correct, zero loss.
- **Borderless text tables** (`_grid_from_words`): infer columns from persistent vertical whitespace gaps, cluster words into rows, fold wrapped continuation lines back into the row above. Shared engine, also used by the OCR path.
- **Scanned image-only pages** (`_ocr_page`): some statements have NO text layer at all (e.g. the Ameris statement is 30 pages of scanned images, 0 characters). The only way to get data out is OCR. Reuses the already-bundled `winocr` (Windows built-in OCR, offline, no extra binary, same engine as Custom % legibility), renders the page at 300 DPI, OCRs it, and runs the word boxes through `_grid_from_words`. **Never returns an empty page:** `_ocr_page` also returns each OCR line's plain text, and if the positional grid comes back empty or too thin, the raw text is written out instead (split into rough columns on runs of 2+ spaces via `_split_line_cols`), so a scanned file always produces usable content even if columns need light manual cleanup. The earlier version produced an empty file on the Ameris scan (this was the reported bug); a simulated-winocr test now yields 30 populated sheets with reconstructed columns. **Still Windows-only and unverified on real winocr output** (winocr can't import on Mac; on Mac these pages get a "needs the Windows OCR engine" note). `_winrt_iter` defensively handles winocr's runtime vectors by index if they aren't directly iterable, which is the most likely reason the first version came out empty.
  - Bug fixed along the way: `_grid_from_words` sorted `(y_centre, word_dict)` tuples, which raised `TypeError` when two words shared a y-centre. Now sorts by key only. This path (borderless text + OCR) is the only caller; the ruled path was unaffected, which is why the ICICI file never hit it.
- Falls back to the old `extract_tables()` if the smarter paths find nothing, so simple PDFs that already converted keep working.

Drag-and-drop works everywhere via byte-transfer (`Api.receive_dropped_file`), not by trying to detect a filesystem path from the browser side (unreliable across pywebview versions).

### Compression: the most-iterated feature, now has 4 modes

Low / Recommended / Extreme / **Custom %** (type a target, tool tries to hit it, backs off with a stated reason if it can't do so legibly).

`compress_pdf()` in `core/optimize.py` picks one of two fundamentally different strategies per file:
- **RECOMPRESS**: shrink embedded images in place + subset fonts. Used when images dominate the file (scans/photos), or when there's real selectable text worth keeping crisp.
- **RASTERIZE**: flatten every page to one JPEG. Used when a file has no real text AND images aren't dominant either, some report/statement generators draw every character as vector path outlines instead of text, producing huge, barely-compressible content streams with nothing to select in the first place, so rasterizing costs nothing that wasn't already lost. Also used at Extreme for ordinary text/vector PDFs once someone's explicitly asked for max compression.

Validated numbers (all independently checked with poppler + visually inspected for legibility):
| File type | Low | Recommended | Extreme |
|---|---|---|---|
| Scanned legal doc (image-dominant) | 63.5% | 90.4% | 96.5% |
| Vector-outline "fake text" (a real bank statement with zero selectable text) | 68.3% | 89.2% | 96.1% |
| Genuine text/vector (PowerPoint export) | 30.0% | 55.9% | 67.5% |

**Custom %** (`compress_pdf_custom()`) walks a 6-rung ladder per strategy (gentler interpolations between the same Low/Recommended/Extreme points, never more aggressive than Extreme already is) and checks legibility at each rung using **Windows' own built-in OCR engine** via the `winocr` package, no external OCR binary bundled, it's the same engine behind the Snipping Tool's "Extract text." Compares OCR'd words on 2-3 sample pages against ground truth (real text if the PDF has it, or an OCR pass on the original at 300 DPI if not) and requires >=85% word accuracy to accept a rung. Stops and reports why if the target can't be hit legibly. Degrades gracefully (best-effort, clearly labeled as unverified) on any machine where `winocr` isn't available, i.e. this Mac.

### Known limitations, not yet verified for real
- **Word/Excel/PowerPoint to PDF**: COM automation code is written correctly (uses `Dispatch()` not `EnsureDispatch()`, `pythoncom.CoInitialize()` per thread, both documented gotchas below) but GitHub's Windows runners don't have Office installed, so CI only proves the exe builds, not that these three tools work. Needs a real run on an office PC.
- **Custom % OCR verification**: same story, `winocr` needs Windows to even import. The search/ladder/backoff logic is fully tested on Mac (with OCR unavailable, so it exercises the graceful-degradation path); the actual OCR calls are unverified until she runs it on Windows.

## Non-obvious technical gotchas (learned the hard way, don't rediscover them)

1. **`doc.subset_fonts()` can segfault** (native crash, uncatchable by try/except) on PDFs whose fonts are already subsetted, the normal case for most Word/PowerPoint exports. Fixed by running it in an isolated `multiprocessing.Process` with a timeout (`_subset_fonts_isolated` in `core/optimize.py`); falls back to the unsubsetted file if it crashes or times out.
2. **Office COM automation**: must use `win32com.client.Dispatch()`, not `gencache.EnsureDispatch()`, in a frozen PyInstaller exe (`EnsureDispatch`'s gencache write is unreliable when frozen). Must call `pythoncom.CoInitialize()`/`CoUninitialize()` per thread since pywebview dispatches each API call on its own thread and COM needs explicit per-thread init. See `core/convert_to_pdf.py` docstring.
3. **`page.replace_image()`**, not manual xref/stream editing, is the only reliable way to swap embedded PDF images without corrupting the JPEG stream on save (learned this from real corruption, "no memory found" errors).
4. **`doc.save()` needs `deflate_images=True, deflate_fonts=True, use_objstms=True, compression_effort=100`** in addition to the more commonly-known `garbage=4, deflate=True, clean=True`, easy to miss, meaningfully affects compression ratio.
5. **Compression must be adaptive** (see RECOMPRESS vs RASTERIZE above), a single fixed strategy cannot work well across scan-heavy, vector-outline, and genuine-text PDFs. Detection: `_image_ratio()` and `_has_real_text()` in `core/optimize.py`.
6. **Whole-file safety net**: `compress_pdf` always compares output size to input and falls back to copying the original if "compressed" would be bigger. Never skip this.
7. **PDF to PowerPoint text fragmentation**: PyMuPDF's `get_text("dict")` "lines" grouping follows the PDF's internal text-showing operators, not visual position. Some generators (PowerPoint's own PDF export included) place every word at its own explicit coordinate, which PyMuPDF reports as one "line" per word. Must re-cluster spans by vertical bbox overlap to reconstruct real sentences, see `_visual_lines()` in `core/convert_from_pdf.py`.
8. **PUA bullet glyphs**: PowerPoint-exported PDFs often use Unicode Private Use Area codepoints (Wingdings-style symbol font glyphs) for bullets. Extracted as plain text with no font mapping, they render as broken boxes. Sanitize by replacing with a normal bullet character, see `_sanitize_text()`.
9. **Google Fonts can be fetched directly** via `curl` from the `fonts.googleapis.com` CSS API to get real `fonts.gstatic.com` file URLs, used to properly bundle Playfair Display (`ui/fonts/PlayfairDisplay-Black.ttf`, with its OFL license alongside).
10. **No em dashes anywhere**, explicit standing instruction, applies to UI copy AND code comments/docstrings/README. Already scrubbed clean as of this handover; keep it that way in anything new.
11. **PyInstaller build flags currently in use** (`.github/workflows/build-windows.yml`): `--collect-all pdfminer` (PDF to Excel needs pdfminer's data files, silent failure otherwise), `--hidden-import win32timezone` (common pywin32 + PyInstaller gotcha), `--collect-all winocr --collect-all winsdk` (for Custom % OCR).

## Design system (current, as approved)

- **Palette**: Navy, Green, Ivory. Warm ivory panels/background (not cool grey), navy as primary accent (`#0B4A85`-ish, sampled from the firm's real logo), green (`#4C9A3B`-ish, also sampled from the logo) used prominently, not just decoratively, currently the filter-pill-bar background. Gold (`#B8901E`) reserved for the small "100% Local Files" seal badge only.
- **Typography**: Playfair Display (bundled TTF, bold/900 weight) for the "PJS" / "PDF" stacked wordmark in the banner, no connecting swash/line (explicitly removed per feedback, "the way it is written is perfect, just remove the line"). Segoe UI system stack for all other UI text (this is a native Windows app; using the OS's own font is a deliberate, grounded choice, not a placeholder). Cascadia Mono for data/paths/the small tracked "STUDIO" and "LOCAL & SECURE" caption lines.
- **Fluid sizing**: banner title, STUDIO/subtitle lines, and filter pills all use CSS `clamp()` tied to viewport width, not fixed pixel sizes, so they visibly scale on a maximized/wide Windows screen instead of just leaving more empty margin. This was a real bug found via a screenshot from her actual Windows machine; don't regress it back to fixed px.
- **Icons**: single document glyph per tool (not the old two-file-plus-arrow motif) with a small rounded-square corner badge (PDF/W/X/P/IMG) marking the conversion format for the 8 conversion tools. Rounded-square icon backgrounds throughout (not circular).
- **Filter pills**: centered in their row (not left-aligned), fluid-sized.
- **Card borders**: 1.5px resting, 2px navy on hover/focus (replaced the old separate focus-outline-ring approach, which would have doubled up visually with a border-color change).
- **Firm logo**: real image (Piyush J. Shah & Co., embedded as base64 in `ui/index.html`), sits at the left of the banner, sourced from `/Users/charveepatel/Downloads/25 years 5.png` originally.
- **A second, more elaborate hand-lettered logo** (with a decorative swash) was shown to me as a pasted chat image, never as a file I could read, so I couldn't embed it directly, I built a coded interpretation instead (the current stacked PJS/PDF wordmark). If she wants the exact custom logo, get the actual image file path from her (drag it into a Finder window, or `@`-reference it like she's done with other files) and embed it directly rather than approximating in code again.

## Test files on hand (all on her Mac, already used extensively, characteristics known)

- `/Users/charveepatel/Downloads/Sales deed of Shop No 20.pdf`, 16 pages, scanned legal doc, 300 DPI grayscale scans, 99.8% image-dominant, exercises the RECOMPRESS path.
- `/Users/charveepatel/Downloads/Export Presentation.pdf`, 23 pages, PowerPoint export, real text/vector, exercises RECOMPRESS (Low/Recommended) and RASTERIZE (Extreme).
- `/Users/charveepatel/Downloads/Axis statement (1).pdf`, 21 pages, 20MB, bank statement with zero real text (vector-outline "fake text") and negligible images, exercises RASTERIZE at every tier. This one is the reason the adaptive-strategy rewrite happened, worth remembering why it's special.
- `/Users/charveepatel/Downloads/OpTransactionHistoryUX322-07-2026 (3).pdf`, 10 pages, ICICI bank statement with a real text layer and full gridlines. The reference file for the ruled-table PDF-to-Excel path (210 transactions, 9 columns).
- `/Users/charveepatel/Downloads/Ameris statement June.pdf`, 30 pages, 8.8MB, US bank statement that is a pure scanned image (44 to 54 image strips per page, zero selectable text). The reference file for the scanned/OCR PDF-to-Excel path, and the reason that path exists. Only convertible to Excel on Windows (winocr).
- Also provided but not yet deeply exercised: `Daphne 2 - BTS - Riva Advisors_v2.xlsx`, `Wall+Street+Prep+_+The+RedBook.pdf` (285 pages, 7.3MB), `GRE_Charvee Patel.pdf` (551KB), given as a broader OCR/compression test set for Custom %, worth running through if touching that feature again.

## Working style notes (how she wants this run)

- **No em dashes**, anywhere, ever, restated multiple times, treat as a hard rule.
- Wants rigorous, evidence-based answers: visual proof, independent validation (poppler, not just PyMuPDF reading back its own output), not just trusting size/percentage numbers.
- Values honesty about trade-offs (e.g., compression ratio vs. legibility) over optimistic claims.
- Prefers being asked when something is genuinely ambiguous rather than guessed at, but doesn't want excessive back-and-forth on things that have a clear right answer.
- Recently asked explicitly for "least amount of credits", bias toward efficient, direct execution over exploratory back-and-forth once a plan is agreed. Batch validation, avoid redundant re-testing of things already proven.
- Big asks tend to arrive as combined bug-report-plus-feature-request messages, often with real files attached (`@"/path"`) or pasted screenshots. Always investigate the actual attached file/screenshot before proposing a fix, several past "bugs" turned out to be genuine issues only visible once the real file was examined (e.g., the bank statement's vector-outline text, the segfault only triggering on certain font-subsetting states).
- She reviews mockups before real-app changes for anything visual/design-related; build a mockup as a Claude Artifact first, wait for explicit approval, then port into `ui/index.html`. Don't skip the mockup step for visual changes even under time pressure.
- When something is genuinely uncertain (like the OCR bundling risk was, before discovering `winocr`), say so plainly rather than papering over it, she has explicitly rewarded that kind of honesty in this conversation.
