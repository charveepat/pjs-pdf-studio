"""Redaction (auto-detect + apply) and password protection."""
import re

import fitz

PATTERNS = {
    "pan": re.compile(r"\b[A-Z]{5}[0-9]{4}[A-Z]\b"),
    "aadhaar": re.compile(r"\b\d{4}\s?\d{4}\s?\d{4}\b"),
    "bank": re.compile(r"\b\d{9,18}\b"),
    "email": re.compile(r"\b[\w.+-]+@[\w-]+\.[\w.-]+\b"),
    "phone": re.compile(r"(?:\+91[\s-]?)?\b[6-9]\d{9}\b"),
}


def scan_sensitive(path: str, pattern_keys: list[str]) -> list[dict]:
    """Regex-based first pass over each page's text layer. Bank-account and
    Aadhaar patterns are broad by design — the UI shows every hit with its
    own checkbox so a false positive just gets unchecked, not blacked out."""
    doc = fitz.open(path)
    results = []
    for pno in range(len(doc)):
        page = doc[pno]
        text = page.get_text("text")
        for key in pattern_keys:
            pattern = PATTERNS.get(key)
            if not pattern:
                continue
            for m in pattern.finditer(text):
                match_str = m.group(0)
                for rect in page.search_for(match_str):
                    results.append(
                        {
                            "type": key,
                            "page": pno + 1,
                            "text": match_str,
                            "rect": [rect.x0, rect.y0, rect.x1, rect.y1],
                        }
                    )
    doc.close()
    return results


def redact_pdf(path: str, boxes: list[dict], save_path: str) -> None:
    """boxes: [{"page": 1-indexed, "rect": [x0,y0,x1,y1]}, ...].
    Uses PyMuPDF's real redaction annotations, which strip the underlying
    text/image content on apply — not just a black rectangle drawn on top."""
    doc = fitz.open(path)
    for box in boxes:
        page = doc[box["page"] - 1]
        page.add_redact_annot(fitz.Rect(box["rect"]), fill=(0, 0, 0))
    for page in doc:
        page.apply_redactions()
    doc.save(save_path, garbage=4, deflate=True)
    doc.close()


def password_protect(path: str, password: str, save_path: str) -> None:
    doc = fitz.open(path)
    perm = (
        fitz.PDF_PERM_PRINT
        | fitz.PDF_PERM_COPY
        | fitz.PDF_PERM_ANNOTATE
        | fitz.PDF_PERM_ACCESSIBILITY
    )
    doc.save(
        save_path,
        encryption=fitz.PDF_ENCRYPT_AES_256,
        user_pw=password,
        owner_pw=password,
        permissions=perm,
    )
    doc.close()
