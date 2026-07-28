"""Convert static PDFs to fillable AcroForms (PDF form fields).

Detection strategy:
1. Vector drawings: horizontal lines (form underlines) and empty rectangles
   are the most reliable indicator in designed and scanned forms.
2. Text underscores: sequences of underscores used in many template forms.

The resulting PDF opens in any AcroForm-capable viewer (Adobe Reader, Edge,
Chrome). The viewer handles saving filled data back to the same file."""

import fitz


def _horiz_lines_and_boxes(page: fitz.Page) -> list:
    results = []
    for d in page.get_drawings():
        items = d.get("items", [])
        rect = d.get("rect")
        if not rect or rect.is_empty or rect.width < 30:
            continue

        # Single horizontal line segment (form underline)
        if len(items) == 1 and items[0][0] == "l":
            p1, p2 = items[0][1], items[0][2]
            dy = abs(p1.y - p2.y)
            dx = abs(p2.x - p1.x)
            if dy < 3 and dx >= 40:
                y = (p1.y + p2.y) / 2
                field_rect = fitz.Rect(
                    min(p1.x, p2.x) + 1, y - 13,
                    max(p1.x, p2.x) - 1, y - 1,
                )
                if field_rect.is_valid and field_rect.height > 4:
                    results.append(("text", field_rect))

        # Rectangular box with no fill or white fill (input box)
        elif rect.height >= 12 and rect.height <= 60 and rect.width >= 50:
            fill = d.get("fill")
            if fill is None or (
                isinstance(fill, (list, tuple)) and len(fill) >= 3
                and all(v > 0.85 for v in fill[:3])
            ):
                inset = fitz.Rect(rect.x0 + 2, rect.y0 + 2, rect.x1 - 2, rect.y1 - 2)
                if inset.is_valid and inset.height > 4:
                    results.append(("text", inset))

    return results


def _underscore_fields(page: fitz.Page) -> list:
    results = []
    for rect in page.search_for("___"):
        field_rect = fitz.Rect(rect.x0, rect.y0 - 12, rect.x1, rect.y0 - 1)
        if field_rect.is_valid and field_rect.height > 3:
            results.append(("text", field_rect))
    return results


def _deduplicate(fields: list) -> list:
    kept = []
    for ft, rect in fields:
        area = rect.width * rect.height
        if area <= 0:
            continue
        overlaps = False
        for _, kr in kept:
            inter = rect & kr
            if not inter.is_empty and (inter.width * inter.height) / area > 0.5:
                overlaps = True
                break
        if not overlaps:
            kept.append((ft, rect))
    return kept


def _label_for(page: fitz.Page, field_rect: fitz.Rect) -> str:
    left = fitz.Rect(
        max(0, field_rect.x0 - 180), field_rect.y0 - 4,
        field_rect.x0 - 2, field_rect.y1 + 4,
    )
    text = page.get_textbox(left).strip().rstrip(":").strip()
    if text:
        return text[-40:]
    above = fitz.Rect(
        field_rect.x0, max(0, field_rect.y0 - 18),
        field_rect.x1, field_rect.y0 - 2,
    )
    text = page.get_textbox(above).strip().rstrip(":").strip()
    return text[-40:] if text else ""


def pdf_to_fillable(path: str, save_path: str) -> dict:
    """Overlay AcroForm text fields on detected blank areas in a static PDF.

    Returns {"ok": True, "fields": N} where N is total fields placed."""
    doc = fitz.open(path)
    total = 0

    for pno, page in enumerate(doc):
        candidates = _deduplicate(_horiz_lines_and_boxes(page) + _underscore_fields(page))

        for idx, (_, rect) in enumerate(candidates):
            label = _label_for(page, rect)
            safe_label = label.replace("/", "-").replace("\n", " ") or str(idx)
            name = f"p{pno+1}_{safe_label}"

            w = fitz.Widget()
            w.rect = rect
            w.field_type = fitz.PDF_WIDGET_TYPE_TEXT
            w.field_name = name
            w.field_value = ""
            w.text_fontsize = min(11, max(7, rect.height * 0.65))
            w.text_color = (0.05, 0.25, 0.65)
            w.fill_color = (0.94, 0.97, 1.0)
            w.border_color = (0.4, 0.55, 0.8)
            w.border_width = 0.5

            try:
                page.insert_widget(w)
                total += 1
            except Exception:
                continue

    doc.save(save_path, garbage=4, deflate=True)
    doc.close()
    return {"ok": True, "fields": total}
