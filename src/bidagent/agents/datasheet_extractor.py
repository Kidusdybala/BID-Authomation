"""Datasheet extractor: for a product that is not in the library yet, reads the datasheet pages
that mention its model and builds a DRAFT library entry.

Anti-invention rule: every feature must come with a word-for-word quote from the datasheet.
Code checks the quote really is on the page. Features that fail the check are kept as
unverified and never used in the bid document or the compliance table."""
import re
from ..util import collapse, grounded, content_words, norm
from ..library import Library

SYSTEM = """[agent:datasheet]
You extract product facts from datasheet pages. Use ONLY the pages given. Never use outside knowledge.
Return ONE JSON object:
{
 "name": "product name/type as written, e.g. 'Network Camera'",
 "description": ["1-3 sentences describing the product, reusing the datasheet's own wording"],
 "features": [{"label": "short spec name", "value": "the value with its unit",
               "quote": "the exact words copied from the page that prove it (max 200 chars)",
               "page": <the [#n] number of the page>}]
}
Give 8-25 of the most important specifications. If the pages are about a different product than
the requested model, return {"name": "", "description": [], "features": []}."""


def _page_label(i, p):
    return f"[#{i}] {p['file']} p.{p['page']}"


def extract_product(llm, ds, model, item_hint=""):
    """Returns (product_dict | None, warnings)."""
    pages = ds.find(model, hint=item_hint, k=ds.max_pages, max_chars=ds.max_chars)
    if not pages:
        return None, [f"No datasheet page mentions '{model}'."]
    text = "\n\n".join(f"{_page_label(i, p)}\n{p['text']}" for i, p in enumerate(pages, 1))
    user = f"Requested model: {model}\nBOM item: {item_hint}\n\nDATASHEET PAGES:\n{text}"
    d = llm.complete_json("datasheet", SYSTEM, user, max_tokens=3000)
    if not isinstance(d, dict) or not d.get("features"):
        return None, [f"Datasheet for '{model}' found but nothing could be extracted."]

    prod = Library.new_product(model, name=collapse(d.get("name", "")) or item_hint, status="draft")
    prod["source"] = {"datasheets": sorted({p["file"] for p in pages})}
    all_text = " ".join(p["text"] for p in pages)
    feats, bad = [], 0
    for f in d["features"]:
        if not isinstance(f, dict) or not f.get("value"):
            continue
        try:
            pg = pages[int(f.get("page", 0)) - 1]
        except (ValueError, IndexError, TypeError):
            pg = None
        quote = collapse(f.get("quote", ""))
        ok = grounded(quote, pg["text"] if pg else all_text) or (pg is None and grounded(quote, all_text))
        if ok:   # the value itself must also be inside the quote (numbers especially)
            ok = grounded(f["value"], quote) or all(
                n in norm(quote) for n in re.findall(r"\d+(?:\.\d+)?", str(f["value"])))
        bad += 0 if ok else 1
        feats.append({"label": collapse(f.get("label", "")), "value": collapse(f["value"]),
                      "source": f"{pg['file']} p.{pg['page']}" if pg else pages[0]["file"],
                      "quote": quote, "verified": bool(ok)})
    # description sentences must be mostly made of words that exist in the datasheet
    pool = content_words(all_text)
    desc = []
    for s in d.get("description") or []:
        w = content_words(s)
        if w and len(w & pool) / len(w) >= 0.7:
            desc.append(collapse(s))
    heading = prod["name"] if prod["name"].lower().startswith(model.lower()[:4]) else f"{prod['name']} {model}".strip()
    prod["blocks"] = [{"heading": heading, "description": desc, "features": feats, "image": None,
                       "image_width_cm": None, "caption": heading}]
    warns = []
    if bad:
        warns.append(f"{model}: {bad} of {len(feats)} extracted specs could not be matched to the "
                     f"datasheet text and were left out (unverified).")
    if not desc:
        warns.append(f"{model}: no description could be taken from the datasheet; write one in the library.")
    return prod, warns
