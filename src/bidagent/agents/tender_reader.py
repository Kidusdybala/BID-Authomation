"""Tender reader: turns the client's tender document into structured facts + requirements."""
import re
from ..util import collapse, content_words, norm

SYSTEM = """[agent:tender]
You read tender documents for a bid team. Extract ONLY what the text says. Never guess or invent.
If a field is not in the text, use "" or []. Return ONE JSON object:
{
 "project_name": "", "client": "", "tender_no": "", "location": "", "deadline": "",
 "summary": "3-5 factual sentences on what the client wants",
 "purpose": ["why the client is doing this, as stated"],
 "objectives": ["client goals as stated"],
 "scope": ["what must be supplied/installed/done, as stated"],
 "requirements": [{"item_no": "BOM item number if the tender ties it to one, else \\"\\"",
                   "text": "one technical requirement, copied or tightly paraphrased"}],
 "general_requirements": ["requirements that apply to the whole bid (warranty, delivery, training...)"]
}
Keep every requirement short and keep numbers and units exactly as written."""


def _chunks(text, size, limit):
    paras, cur, out = re.split(r"\n\s*\n", text), "", []
    for p in paras:
        if len(cur) + len(p) > size and cur:
            out.append(cur)
            cur = ""
        while len(p) > size:                       # one huge paragraph
            out.append(p[:size])
            p = p[size:]
        cur += p + "\n\n"
    if cur.strip():
        out.append(cur)
    return out[:limit], max(0, len(out) - limit)


def _merge(parts):
    res = {"project_name": "", "client": "", "tender_no": "", "location": "", "deadline": "",
           "summary": "", "purpose": [], "objectives": [], "scope": [], "requirements": [],
           "general_requirements": []}
    seen = set()
    for d in parts:
        if not isinstance(d, dict):
            continue
        for k in ("project_name", "client", "tender_no", "location", "deadline", "summary"):
            if not res[k] and d.get(k):
                res[k] = collapse(d[k])
        for k in ("purpose", "objectives", "scope", "general_requirements"):
            for v in d.get(k) or []:
                v = collapse(v)
                if v and (k, norm(v)) not in seen:
                    seen.add((k, norm(v)))
                    res[k].append(v)
        for r in d.get("requirements") or []:
            if isinstance(r, str):
                r = {"item_no": "", "text": r}
            t = collapse(r.get("text", ""))
            if t and ("req", norm(t)) not in seen:
                seen.add(("req", norm(t)))
                res["requirements"].append({"item_no": collapse(r.get("item_no", "")), "text": t})
    return res


def read_tender(llm, text, bom_lines, chunk_chars=14000, max_chunks=8):
    """Returns (tender_dict, warnings)."""
    warnings = []
    items = "\n".join(f"{l['no']}: {l['item'][:70]}" for l in bom_lines if not l["is_group"])[:4000]
    chunks, skipped = _chunks(text, chunk_chars, max_chunks)
    if skipped:
        warnings.append(f"Tender is long: only the first {max_chunks} parts were read "
                        f"({skipped} skipped). Raise tender.max_chunks to read more.")
    parts = []
    for i, c in enumerate(chunks, 1):
        user = (f"BOM items (use these numbers for item_no):\n{items}\n\n"
                f"TENDER TEXT (part {i} of {len(chunks)}):\n{c}")
        try:
            parts.append(llm.complete_json("tender", SYSTEM, user, max_tokens=4096))
        except Exception as e:                       # one bad part should not kill the bid
            warnings.append(f"Tender part {i} could not be read by the AI: {e}")
    return _merge(parts), warnings


def read_tender_offline(text):
    """No-AI fallback: only the title line. Everything else is left for you."""
    first = next((collapse(l) for l in text.splitlines() if collapse(l) and not l.startswith("===")), "")
    t = _merge([{"project_name": first[:120]}])
    return t, ["AI unavailable: tender not analysed. Fill in project details by hand."]


def assign_requirements(tender, bom_lines):
    """Attach each requirement to a BOM line. Returns ({item_no: [text]}, [unassigned texts]).
    1) the item_no the tender gave (if it exists in the BOM), 2) word overlap with the item name."""
    leafs = [l for l in bom_lines if not l["is_group"]]
    nos = {l["no"] for l in leafs}
    out, loose = {}, []
    for r in tender.get("requirements", []):
        no, text = r.get("item_no", ""), r["text"]
        if no not in nos:
            no = ""
            rw = content_words(text)
            best, score = None, 0.0
            for l in leafs:
                iw = content_words(l["item"])
                if len(iw) >= 1 and rw:
                    s = len(iw & rw) / len(iw)
                    if s > score:
                        best, score = l, s
            if best and score >= 0.6:
                no = best["no"]
        if no:
            out.setdefault(no, []).append(text)
        else:
            loose.append(text)
    return out, loose
