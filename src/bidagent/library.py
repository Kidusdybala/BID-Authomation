"""Product library: one JSON per product, written once, reused in every bid.

library/
  products/<id>.json   {id, model, aliases, name, status, blocks:[{heading, description[],
                        features[{label,value,source,verified}], image, image_width_cm, caption}]}
  images/              product pictures
  style/examples.json  your own wording for intro sections (the writer agent imitates it)
  templates/           your sample .docx (cover, TOC, company background, table style)
"""
import difflib
from pathlib import Path
from .util import norm, load_json, save_json, slug


class Library:
    def __init__(self, root):
        self.root = Path(root)
        for d in ("products", "images", "style", "templates"):
            (self.root / d).mkdir(parents=True, exist_ok=True)
        self._cache = None

    def all(self):
        if self._cache is None:
            self._cache = [load_json(p) for p in sorted((self.root / "products").glob("*.json"))]
        return self._cache

    def get(self, pid):
        return next((p for p in self.all() if p["id"] == pid), None)

    def save(self, prod):
        save_json(self.root / "products" / f"{prod['id']}.json", prod)
        self._cache = None

    def image_path(self, name):
        return self.root / "images" / name if name else None

    def match(self, *texts):
        """Products whose model/alias appears inside the given BOM text."""
        blob = norm(" ".join(t for t in texts if t))
        hits = []
        for p in self.all():
            toks = [p["model"]] + p.get("aliases", [])
            if any(len(norm(t)) >= 4 and norm(t) in blob for t in toks):
                hits.append(p)
        # drop products whose model is just a fragment of another matched model
        keep = []
        for p in hits:
            m = norm(p["model"])
            if not any(q is not p and m != norm(q["model"]) and m in norm(q["model"]) for q in hits):
                keep.append(p)
        return keep

    def suggest(self, text, n=3):
        names = {norm(p["model"]): p["model"] for p in self.all()}
        return [names[k] for k in difflib.get_close_matches(norm(text), list(names), n=n, cutoff=0.6)]

    def style(self):
        return load_json(self.root / "style" / "examples.json", {})

    def save_style(self, obj):
        save_json(self.root / "style" / "examples.json", obj)

    @staticmethod
    def new_product(model, name="", aliases=None, status="draft"):
        return {"id": slug(model), "model": model, "name": name, "aliases": aliases or [],
                "status": status, "blocks": [], "source": {}}
