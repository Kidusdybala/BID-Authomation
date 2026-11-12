"""Datasheet search engine: reads your PDF datasheets once, then finds the right pages
for a model with BM25 ranking. Runs locally and costs no tokens."""
import math
import re
from pathlib import Path
from .util import norm, load_json, save_json


def tok(s):
    return re.findall(r"[a-z0-9][a-z0-9\-\.]*", str(s).lower())


class BM25:
    def __init__(self, docs, k1=1.5, b=0.75):
        self.docs = [tok(d) for d in docs]
        self.k1, self.b = k1, b
        self.avg = (sum(len(d) for d in self.docs) / len(self.docs)) if self.docs else 0
        df = {}
        for d in self.docs:
            for t in set(d):
                df[t] = df.get(t, 0) + 1
        n = len(self.docs)
        self.idf = {t: math.log(1 + (n - c + 0.5) / (c + 0.5)) for t, c in df.items()}

    def score(self, query, i):
        d = self.docs[i]
        if not d:
            return 0.0
        tf = {}
        for t in d:
            tf[t] = tf.get(t, 0) + 1
        s = 0.0
        for q in set(tok(query)):
            if q in tf:
                s += self.idf.get(q, 0) * tf[q] * (self.k1 + 1) / (
                    tf[q] + self.k1 * (1 - self.b + self.b * len(d) / (self.avg or 1)))
        return s

    def top(self, query, k=5):
        sc = sorted(((self.score(query, i), i) for i in range(len(self.docs))), reverse=True)
        return [(s, i) for s, i in sc[:k] if s > 0]


class DatasheetIndex:
    def __init__(self, folder, cache_file, max_pages=6, max_chars=12000):
        self.max_pages, self.max_chars = max_pages, max_chars
        self.folder = Path(folder)
        self.cache_file = Path(cache_file)
        self.pages = []          # {file, page, text}
        self.warnings = []
        self._bm = None

    def build(self):
        cache = load_json(self.cache_file, {})
        out, new_cache = [], {}
        pdfs = sorted(self.folder.rglob("*.pdf")) if self.folder.exists() else []
        for pdf in pdfs:
            rel = str(pdf.relative_to(self.folder))
            sig = f"{pdf.stat().st_size}-{int(pdf.stat().st_mtime)}"
            if rel in cache and cache[rel]["sig"] == sig:
                pages = cache[rel]["pages"]
            else:
                pages = self._extract(pdf)
            new_cache[rel] = {"sig": sig, "pages": pages}
            if pages and not any(p.strip() for p in pages):
                self.warnings.append(f"{rel}: no text layer (scanned?) - needs OCR")
            for i, t in enumerate(pages, 1):
                out.append({"file": rel, "page": i, "text": t})
        save_json(self.cache_file, new_cache)
        self.pages = out
        self._bm = BM25([p["text"] for p in out]) if out else None
        return len(pdfs), len(out)

    @staticmethod
    def _extract(pdf):
        import pdfplumber
        pages = []
        with pdfplumber.open(str(pdf)) as doc:
            for pg in doc.pages:
                pages.append(pg.extract_text() or "")
        return pages

    def find(self, model, hint="", k=6, max_chars=12000):
        """Pages that actually mention the model (strict, so we never feed unrelated text to the AI)."""
        m = norm(model)
        if len(m) < 4 or not self.pages:
            return []
        cand = [i for i, p in enumerate(self.pages)
                if m in norm(p["text"]) or m in norm(p["file"])]
        if not cand:
            return []
        q = f"{model} {hint}"
        ranked = sorted(cand, key=lambda i: -self._bm.score(q, i))[:k]
        ranked.sort(key=lambda i: (self.pages[i]["file"], self.pages[i]["page"]))
        out, total = [], 0
        for i in ranked:
            p = self.pages[i]
            if total + len(p["text"]) > max_chars:
                p = dict(p, text=p["text"][:max(0, max_chars - total)])
            out.append(p)
            total += len(p["text"])
            if total >= max_chars:
                break
        return out
