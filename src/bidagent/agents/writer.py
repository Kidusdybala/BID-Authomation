"""Writer: drafts the bid-specific intro sections in YOUR style, from tender facts only."""
import json
from ..util import collapse, numbers_in

SYSTEM = """[agent:writer]
You write sections of a technical bid for a supplier, in the supplier's own voice.
RULES
- The STYLE EXAMPLES come from an older bid. Copy their tone, sentence length and structure ONLY.
  Never reuse their names, places, numbers or equipment.
- Every fact (client, place, numbers, equipment) must come from FACTS or BOM. Do not invent any.
- If a fact you need is missing, write [TO CONFIRM: what is missing] in its place.
- Plain paragraphs, no markdown, no bullet symbols. Write "we" for the supplier.
Return ONE JSON object: {"sections": {"<section title>": ["paragraph", "paragraph", ...]}}
using exactly the section titles you are given."""


def _trim(s, n):
    s = str(s or "")
    return s if len(s) <= n else s[:n].rsplit(" ", 1)[0] + " ..."


def write_sections(llm, style, tender, bom_lines, sections, meta):
    """sections: [{title, guidance}]. Returns ({title: [paragraphs]}, warnings)."""
    facts = {k: tender.get(k) for k in ("project_name", "client", "tender_no", "location", "summary",
                                         "purpose", "objectives", "scope", "general_requirements")}
    facts.update({k: v for k, v in meta.items() if v})
    groups = [f"{l['no']} {l['item']}" for l in bom_lines if l["level"] == 0]
    examples = {s["title"]: _trim(style.get(s["title"], ""), 1400) for s in sections if style.get(s["title"])}
    want = [{"title": s["title"], "guidance": s.get("guidance", "")} for s in sections]
    user = ("FACTS:\n" + json.dumps(facts, ensure_ascii=False) +
            "\n\nBOM main groups:\n" + "\n".join(groups[:40]) +
            "\n\nSTYLE EXAMPLES (tone only):\n" + json.dumps(examples, ensure_ascii=False) +
            "\n\nWRITE THESE SECTIONS:\n" + json.dumps(want, ensure_ascii=False))
    d = llm.complete_json("writer", SYSTEM, user, temperature=0.4, max_tokens=4096)
    raw = d.get("sections", d) if isinstance(d, dict) else {}
    out, warnings = {}, []
    src_numbers = numbers_in(user.split("STYLE EXAMPLES")[0]) | numbers_in(json.dumps(want))
    for s in sections:
        paras = raw.get(s["title"])
        if isinstance(paras, str):
            paras = [p for p in paras.split("\n") if p.strip()]
        paras = [collapse(p) for p in (paras or []) if collapse(p)]
        if not paras:
            warnings.append(f'Writer returned nothing for "{s["title"]}".')
            paras = [f'[TO WRITE: {s["title"]}]']
        extra = numbers_in(" ".join(paras)) - src_numbers
        extra = {n for n in extra if len(n) > 1}
        if extra:
            warnings.append(f'"{s["title"]}": numbers not found in the tender/BOM: '
                            f'{", ".join(sorted(extra))}. Check them.')
        out[s["title"]] = paras
    return out, warnings


def write_sections_offline(tender, sections):
    """No-AI fallback: use the tender's own words where we have them, otherwise a marker."""
    out = {}
    key_map = {"project overview": ["summary"], "purpose of the project": ["purpose"],
               "objectives of the project": ["objectives"], "scope of the project": ["scope"]}
    for s in sections:
        vals = []
        for k in key_map.get(s["title"].lower(), []):
            v = tender.get(k)
            vals += [v] if isinstance(v, str) and v else list(v or [])
        out[s["title"]] = vals or [f'[TO WRITE: {s["title"]}]']
    return out, ["AI unavailable: intro sections use the tender's own wording or placeholders."]
