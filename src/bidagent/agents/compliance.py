"""Compliance checker.

The AI only JUDGES (picks which library features answer a requirement, and says yes/partial/no).
The "Offered" column is built by code, copied word for word from the library (or from a quote
that code verified is on the datasheet page). The AI never writes a spec."""
import json
import re
from ..util import collapse, grounded, content_words, norm

STATUSES = ("complies", "partial", "not_complies", "not_found", "to_confirm")
LABEL = {"complies": "Complies", "partial": "Partially complies", "not_complies": "Does not comply",
         "not_found": "To be confirmed", "to_confirm": "To be confirmed"}

SYSTEM = """[agent:compliance]
You check whether a product meets tender requirements, using ONLY the numbered product features.
For each requirement return the numbers of the features that answer it and a verdict:
 "complies"     - the feature value meets or exceeds the requirement
 "partial"      - meets it only in part, or the value is close but lower / different
 "not_complies" - a feature clearly contradicts the requirement
 "not_found"    - no feature answers it (use an empty list)
Compare numbers and units carefully (e.g. 8 MP does not meet >= 12 MP).
Return ONE JSON object: {"results": [{"i": <requirement number>, "features": [<feature numbers>],
 "status": "complies|partial|not_complies|not_found", "note": "max 15 words"}]}"""

SYSTEM2 = """[agent:datasheet]
Find where the datasheet pages answer each requirement. Use ONLY the pages given.
For each requirement return the exact words copied from a page, and the verdict
(complies / partial / not_complies / not_found).
Return ONE JSON object: {"results": [{"i": <requirement number>, "found": true|false,
 "quote": "exact words copied from the page (max 200 chars)", "page": <[#n] number>,
 "status": "complies|partial|not_complies|not_found"}]}"""

_MIN = r"(?:>=|≥|at least|minimum(?: of)?|min\.?|not less than|no less than|over|above)"
_MAX = r"(?:<=|≤|at most|maximum(?: of)?|max\.?|not more than|no more than|up to|below|under)"
_NUM = r"(\d+(?:[.,]\d+)?)\s*([a-zA-Z%°]{1,6})"


def numeric_conflict(req, offered):
    """Cheap safety net for 'complies' verdicts: if the requirement says 'at least N unit' (or
    'at most') and the offered text has a smaller/larger number in the SAME unit, return a note."""
    for kind, pat in (("min", _MIN), ("max", _MAX)):
        for m in re.finditer(pat + r"\s*" + _NUM, req, flags=re.I):
            need, unit = float(m.group(1).replace(",", ".")), m.group(2).lower()
            haves = [(n, float(n.replace(",", "."))) for n, u in re.findall(_NUM, offered)
                     if u.lower() == unit]
            if haves and not any((h >= need) if kind == "min" else (h <= need) for _, h in haves):
                return (f"offered {haves[0][0]} {unit} vs required "
                        f"{'>=' if kind == 'min' else '<='} {m.group(1)} {unit}")
    return ""


def _feature_pool(products):
    pool = []
    for p in products:
        for b in p.get("blocks", []):
            for f in b.get("features", []):
                if f.get("verified", True) and f.get("value"):
                    txt = f"{f['label']}: {f['value']}" if f.get("label") else f["value"]
                    pool.append({"text": txt, "source": f.get("source", ""), "model": p["model"]})
    return pool


def _row(req, offered="", status="not_found", source="", note="", via="none"):
    return {"req": req, "offered": offered, "status": status, "label": LABEL[status],
            "source": source, "note": note, "via": via}


def check_item(llm, reqs, products, ds=None, second_pass=True):
    """One BOM line. Returns (rows, warnings)."""
    pool = _feature_pool(products)
    warnings, rows = [], []
    model = products[0]["model"] if products else ""
    if llm is None or not pool:
        results = {}
    else:
        user = json.dumps({"requirements": [{"i": i, "text": r} for i, r in enumerate(reqs)],
                           "features": [{"i": i, "text": f["text"]} for i, f in enumerate(pool)]},
                          ensure_ascii=False)
        try:
            d = llm.complete_json("compliance", SYSTEM, user, max_tokens=3000)
            results = {int(r["i"]): r for r in d.get("results", []) if str(r.get("i", "")).isdigit()}
        except Exception as e:
            warnings.append(f"{model}: compliance AI call failed ({e}); rows marked for manual check.")
            results = {}
    for i, req in enumerate(reqs):
        r = results.get(i)
        if r is None:
            rows.append(_offline_row(req, pool))
            continue
        idx = [k for k in (r.get("features") or []) if isinstance(k, int) and 0 <= k < len(pool)]
        status = r.get("status") if r.get("status") in STATUSES else "to_confirm"
        if status in ("complies", "partial", "not_complies") and not idx:
            status = "not_found"              # a verdict with no evidence is not a verdict
        offered = "; ".join(pool[k]["text"] for k in idx)
        src = ", ".join(dict.fromkeys(pool[k]["source"] for k in idx if pool[k]["source"]))
        note = collapse(r.get("note", ""))
        if status == "complies":
            c = numeric_conflict(req, offered)
            if c:
                status, note = "partial", f"Check: {c}"
        rows.append(_row(req, offered, status, src, note, "library" if idx else "none"))

    # second pass: unresolved requirements -> datasheet pages that mention this model
    todo = [i for i, r in enumerate(rows) if r["status"] in ("not_found", "to_confirm")]
    if second_pass and llm is not None and ds is not None and todo and model:
        pages = ds.find(model, hint=" ".join(reqs[i] for i in todo), k=ds.max_pages, max_chars=ds.max_chars)
        if pages:
            text = "\n\n".join(f"[#{n}] {p['file']} p.{p['page']}\n{p['text']}" for n, p in enumerate(pages, 1))
            user = json.dumps({"requirements": [{"i": i, "text": reqs[i]} for i in todo]}) + "\n\nPAGES:\n" + text
            try:
                d = llm.complete_json("datasheet", SYSTEM2, user, max_tokens=2500)
                for r in d.get("results", []):
                    i = r.get("i")
                    if i not in todo or not r.get("found"):
                        continue
                    try:
                        pg = pages[int(r.get("page")) - 1]
                    except (ValueError, TypeError, IndexError):
                        continue
                    quote = collapse(r.get("quote", ""))
                    st = r.get("status") if r.get("status") in STATUSES else "to_confirm"
                    if not grounded(quote, pg["text"]) or st == "not_found":
                        continue
                    note = ""
                    if st == "complies" and numeric_conflict(reqs[i], quote):
                        st, note = "partial", "Check: " + numeric_conflict(reqs[i], quote)
                    rows[i] = _row(reqs[i], quote, st, f"{pg['file']} p.{pg['page']}", note, "datasheet")
            except Exception as e:
                warnings.append(f"{model}: datasheet lookup failed ({e}).")
    return rows, warnings


def _offline_row(req, pool):
    """No-AI fallback: best word-overlap feature, always marked 'to be confirmed'."""
    rw = content_words(req)
    best, score = None, 0.0
    for f in pool:
        fw = content_words(f["text"])
        if fw and rw:
            s = len(fw & rw) / len(rw)
            if s > score:
                best, score = f, s
    if best and score >= 0.3:
        return _row(req, best["text"], "to_confirm", best["source"], "Closest library feature; please verify.", "library")
    return _row(req)
