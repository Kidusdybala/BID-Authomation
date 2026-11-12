"""Small helpers shared by every module (no AI, no tokens)."""
import json
import re
from pathlib import Path


def norm(s):
    """Lowercase, letters+digits only. Used for model matching."""
    return re.sub(r"[^a-z0-9]", "", str(s or "").lower())


def slug(s, maxlen=48):
    s = re.sub(r"[^A-Za-z0-9]+", "-", str(s)).strip("-").lower()
    return s[:maxlen].strip("-") or "item"


def collapse(s):
    return re.sub(r"\s+", " ", str(s or "")).strip()


def load_json(p, default=None):
    p = Path(p)
    return json.loads(p.read_text("utf-8")) if p.exists() else default


def save_json(p, obj):
    p = Path(p)
    p.parent.mkdir(parents=True, exist_ok=True)
    p.write_text(json.dumps(obj, indent=2, ensure_ascii=False), "utf-8")


def model_tokens(s):
    """Tokens that look like model numbers: contain a digit, 4+ chars."""
    out = []
    for t in re.split(r"[\s,;/()+]+", str(s or "")):
        t = t.strip(".-")
        if len(norm(t)) >= 4 and re.search(r"\d", t) and t not in out:
            out.append(t)
    return out


def split_models(cell):
    """Split a messy BOM model cell into parts. Returns None if too complex to trust."""
    cell = str(cell or "")
    if len(collapse(cell)) > 220:
        return None
    parts = [collapse(p) for p in re.split(r"\s\+\s|;|\n", cell) if collapse(p)]
    return parts if len(parts) <= 6 else None


def read_text_any(path):
    """Plain text from txt/md/pdf/docx/xlsx (used for tender documents)."""
    path = Path(path)
    ext = path.suffix.lower()
    if ext in (".txt", ".md", ".csv"):
        return path.read_text("utf-8", errors="ignore")
    if ext == ".pdf":
        import pdfplumber
        out = []
        with pdfplumber.open(str(path)) as pdf:
            for i, pg in enumerate(pdf.pages, 1):
                out.append(f"[page {i}]\n" + (pg.extract_text() or ""))
        return "\n\n".join(out)
    if ext == ".docx":
        from docx import Document
        d = Document(str(path))
        out = [p.text for p in d.paragraphs if p.text.strip()]
        for t in d.tables:
            for r in t.rows:
                cells = []
                for c in r.cells:
                    if c.text.strip() and c.text.strip() not in cells:
                        cells.append(c.text.strip())
                out.append(" | ".join(cells))
        return "\n".join(out)
    if ext in (".xlsx", ".xlsm"):
        import openpyxl
        wb = openpyxl.load_workbook(str(path), data_only=True)
        out = []
        for ws in wb:
            for row in ws.iter_rows(values_only=True):
                cells = [str(c) for c in row if c is not None and str(c).strip()]
                if cells:
                    out.append(" | ".join(cells))
        return "\n".join(out)
    raise ValueError(f"Unsupported file type: {path}")


def grounded(quote, text):
    """True if `quote` appears (ignoring case, spacing and punctuation) inside `text`."""
    q = norm(quote)
    return len(q) >= 4 and q in norm(text)


def numbers_in(text):
    """Numbers as plain digit strings (commas/spaces removed), e.g. '4,000' -> '4000'."""
    import re as _re
    out = set()
    for m in _re.findall(r"\d[\d,]*(?:\.\d+)?", str(text or "")):
        out.add(m.replace(",", "").rstrip("."))
    return out


def content_words(s):
    stop = {"the", "and", "for", "with", "of", "to", "a", "an", "in", "on", "or", "be", "is", "are",
            "shall", "must", "should", "support", "supports", "minimum", "min", "per", "all", "any"}
    return {w for w in re.findall(r"[a-z0-9]+", str(s).lower()) if len(w) > 2 and w not in stop}
