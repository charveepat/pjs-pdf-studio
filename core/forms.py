"""Convert static PDFs to fillable AcroForms with per-cell auto-advance.

Three detection strategies cover the common Indian government/banking form types:

1. White-box forms (RTGS/NEFT, bank forms):
   Each character cell is a white-filled rectangle (~14x17 pt).

2. Line-grid forms (Form 93 PAN, passport):
   Each cell is reconstructed from short horizontal line segments that form
   a grid (top/bottom of cells drawn as individual ~15pt lines).

3. Black-bar forms (Form 134/135 TAN, income-tax):
   Each cell is the gap between thin black-filled bar rectangles.

Result: one text widget per cell, maxlength=1, with a PDF JavaScript
keystroke action that calls the next cell's setFocus() automatically.
Works in Adobe Reader (free). In Edge/Chrome, fields are still fillable;
user presses Tab to advance instead.
"""

from __future__ import annotations
from collections import defaultdict

import fitz


# ---------------------------------------------------------------------------
# Strategy 1: white-filled rect boxes
# ---------------------------------------------------------------------------

def _detect_white_boxes(page: fitz.Page) -> list[fitz.Rect]:
    cells = []
    for d in page.get_drawings():
        r = d.get("rect")
        if not r:
            continue
        items = d.get("items", [])
        if not items or items[0][0] != "re":
            continue
        w, h = r.width, r.height
        if not (6 <= w <= 30 and 6 <= h <= 28):
            continue
        fill = d.get("fill")
        if fill and len(fill) >= 3 and all(v > 0.78 for v in fill[:3]):
            cells.append(r)
    return cells


# ---------------------------------------------------------------------------
# Strategy 2: line-grid (short horizontal segments, ~15pt wide)
# ---------------------------------------------------------------------------

def _detect_line_grid(page: fitz.Page) -> list[fitz.Rect]:
    """Reconstruct character cells from short horizontal line segments."""
    h_segs: list[tuple[float, float, float]] = []  # (x0, y, x1)

    for d in page.get_drawings():
        items = d.get("items", [])
        if len(items) != 1 or items[0][0] != "l":
            continue
        p1, p2 = items[0][1], items[0][2]
        dx = abs(p2.x - p1.x)
        dy = abs(p2.y - p1.y)
        if dy < 1 and 8 <= dx <= 28:
            h_segs.append((min(p1.x, p2.x), (p1.y + p2.y) / 2, max(p1.x, p2.x)))

    if not h_segs:
        return []

    # Group by y (1pt tolerance)
    by_y: dict[int, list[tuple[float, float]]] = defaultdict(list)
    for x0, y, x1 in h_segs:
        by_y[round(y)].append((x0, x1))

    ys = sorted(by_y)
    cells = []
    used_tops: set[int] = set()

    for i, y_top in enumerate(ys):
        if y_top in used_tops:
            continue
        # Look for a matching bottom row within 8-22pt
        for y_bot in ys[i + 1:]:
            gap = y_bot - y_top
            if gap > 22:
                break
            if gap < 8:
                continue
            # Top and bottom segs should have matching x-spans
            tops = sorted(by_y[y_top])
            bots = sorted(by_y[y_bot])
            # Use top segments as cell definitions
            for x0, x1 in tops:
                cells.append(fitz.Rect(x0, y_top, x1, y_bot))
            used_tops.add(y_top)
            used_tops.add(y_bot)
            break

    return cells


# ---------------------------------------------------------------------------
# Strategy 3: black-bar gaps (TAN / income-tax forms)
# ---------------------------------------------------------------------------

def _detect_black_bar_gaps(page: fitz.Page) -> list[fitz.Rect]:
    """Find cells as the gaps between thin black bar rectangles."""
    v_bars: list[fitz.Rect] = []  # vertical dividers
    h_bars: list[fitz.Rect] = []  # top/bottom rails

    for d in page.get_drawings():
        r = d.get("rect")
        if not r:
            continue
        items = d.get("items", [])
        if not items or items[0][0] != "re":
            continue
        fill = d.get("fill")
        if not fill or not all(v < 0.15 for v in fill[:3]):
            continue  # must be black/dark
        w, h = r.width, r.height
        if 0.4 <= w <= 2.5 and 10 <= h <= 35:
            v_bars.append(r)
        elif 8 <= w <= 22 and 0.5 <= h <= 2.5:
            h_bars.append(r)

    if not v_bars:
        return []

    # For each adjacent pair of v_bars (right of one is close to left of next),
    # the gap between them at the same y-range = one cell.
    v_bars.sort(key=lambda r: (round(r.y0), r.x0))

    cells = []
    # Group v_bars by approximate y-row (y0 within 5pt)
    rows: dict[int, list[fitz.Rect]] = defaultdict(list)
    for vb in v_bars:
        rows[round(vb.y0 / 5) * 5].append(vb)

    for _, row_bars in rows.items():
        row_bars.sort(key=lambda r: r.x0)
        for i in range(len(row_bars) - 1):
            left = row_bars[i]
            right = row_bars[i + 1]
            gap = right.x0 - left.x1
            if 8 <= gap <= 25:
                # Cell is between the right edge of left bar and left edge of right bar
                y0 = min(left.y0, right.y0)
                y1 = max(left.y1, right.y1)
                cells.append(fitz.Rect(left.x1, y0 + 1, right.x0, y1 - 1))

    return cells


# ---------------------------------------------------------------------------
# Grouping helpers
# ---------------------------------------------------------------------------

def _group_into_fields(cells: list[fitz.Rect]) -> list[list[fitz.Rect]]:
    """Group individual character cells into logical fields.

    Cells on the same horizontal line (same y, close x) belong to one field.
    The gap threshold between cells in the same field is <= 6pt.
    """
    if not cells:
        return []

    # Sort top-to-bottom, left-to-right
    cells = sorted(cells, key=lambda r: (round(r.y0 / 3) * 3, r.x0))

    fields: list[list[fitz.Rect]] = []
    current: list[fitz.Rect] = [cells[0]]

    for cell in cells[1:]:
        prev = current[-1]
        same_row = abs(cell.y0 - prev.y0) < 4
        close_x = (cell.x0 - prev.x1) <= 8
        if same_row and close_x:
            current.append(cell)
        else:
            fields.append(current)
            current = [cell]

    fields.append(current)
    return fields


def _deduplicate_cells(cells: list[fitz.Rect]) -> list[fitz.Rect]:
    kept: list[fitz.Rect] = []
    for c in cells:
        area = c.width * c.height
        if area <= 0:
            continue
        overlap = False
        for k in kept:
            inter = c & k
            if not inter.is_empty and (inter.width * inter.height) / area > 0.4:
                overlap = True
                break
        if not overlap:
            kept.append(c)
    return kept


# ---------------------------------------------------------------------------
# Main entry point
# ---------------------------------------------------------------------------

_AUTO_ADVANCE_JS = """\
if (!event.willCommit) {{
  if (event.change !== null && event.change !== "") {{
    var nf = this.getField("{next_name}");
    if (nf) nf.setFocus();
  }}
}}"""


def pdf_to_fillable(path: str, save_path: str) -> dict:
    """Overlay AcroForm text fields on detected character cells.

    Returns {"ok": True, "fields": N, "cells": C} where N = widgets placed,
    C = individual character cells detected."""
    doc = fitz.open(path)
    total_widgets = 0
    total_cells = 0

    for pno, page in enumerate(doc):
        # Run all three detectors; use whichever finds more cells
        white = _detect_white_boxes(page)
        lines = _detect_line_grid(page)
        bars  = _detect_black_bar_gaps(page)

        # Pick the strategy with the most cells (or combine if non-overlapping)
        candidates = _deduplicate_cells(white + lines + bars)
        if not candidates:
            continue

        fields = _group_into_fields(candidates)
        total_cells += len(candidates)

        # Build sequential widget names across all fields on this page
        # so we can wire up auto-advance JS referencing the next widget name
        all_cells: list[tuple[str, fitz.Rect]] = []
        for fi, field_cells in enumerate(fields):
            for ci, cell in enumerate(field_cells):
                name = f"pg{pno+1}_f{fi:03d}_c{ci:03d}"
                all_cells.append((name, cell))

        for idx, (name, rect) in enumerate(all_cells):
            next_name = all_cells[idx + 1][0] if idx + 1 < len(all_cells) else None

            w = fitz.Widget()
            w.rect = rect
            w.field_type = fitz.PDF_WIDGET_TYPE_TEXT
            w.field_name = name
            w.field_value = ""
            w.text_fontsize = min(11, max(6, rect.height * 0.6))
            w.text_color = (0.05, 0.2, 0.6)
            w.fill_color = (0.93, 0.96, 1.0)
            w.border_color = (0.4, 0.55, 0.8)
            w.border_width = 0.3

            w.text_maxlen = 1
            if next_name:
                w.script_stroke = _AUTO_ADVANCE_JS.format(next_name=next_name)

            try:
                page.add_widget(w)
                total_widgets += 1
            except Exception:
                continue

    doc.save(save_path, garbage=4, deflate=True)
    doc.close()
    return {"ok": True, "fields": total_widgets, "cells": total_cells}
