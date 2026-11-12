"""Reviewer. Rule checks run first (free), then one AI call looks for things rules can't catch.
Everything ends up in one findings list that becomes review_report.md."""
import difflib
import json
import re
from pathlib import Path
import yaml
from ..util import norm, model_tokens, collapse

SEV = {"high": 0, "medium": 1, "low": 2, "info": 3}
_SPEC = re.compile(r"^\d+(?:\.\d+)?(?:mp|gb|tb|mb|mhz|ghz|w|kw|v|va|hz|k|p|mm|cm|m|in|inch|u|x|port|ports|"
                   r"fps|ah|nm|db|gbps|mbps|user|users|license|licenses|year|years)$")

SYSTEM = """[agent:reviewer]
You are a careful pre-submission reviewer for a technical bid. Look at the BOM lines and requirements.
Report only problems you are fairly sure about:
- the item name and the model disagree (different product, generation or number)
- a requirement says something the named product is known NOT to be (wrong sensor size, wrong port type...)
- the same product listed twice, or an odd quantity or unit for what the item is
Do NOT report style, spelling or things you cannot judge. Do NOT repeat the 'already found' list.
Return ONE JSON object: {"findings": [{"where": "Item 15", "severity": "high|medium|low",
 "issue": "what is wrong, 1-2 sentences", "fix": "what to check or change"}]}"""


def _item_model_tokens(text):
    return [t for t in model_tokens(text) if not _SPEC.match(norm(t))]


def mismatch(item, model):
    """Item text and model text both name a model number, but not the same one."""
    a, b = _item_model_tokens(item), _item_model_tokens(model)
    if not a or not b:
        return None
    na, nb = [norm(t) for t in a], [norm(t) for t in b]
    if any(x in y or y in x for x in na for y in nb):
        return None
    best = max((difflib.SequenceMatcher(None, x, y).ratio(), x, y) for x in na for y in nb)
    close = best[0] >= 0.6
    return {"item_tok": a[na.index(best[1])], "model_tok": b[nb.index(best[2])], "close": close}


def load_facts(path):
    p = Path(path)
    if not p.exists():
        return []
    return (yaml.safe_load(p.read_text("utf-8")) or {}).get("facts", [])


def static_checks(lines, line_parts, reqs_by_item, facts, draft_ids):
    """lines: list of BOM dicts. line_parts: {no: [{"text","products":[id...]}]}."""
    F = []

    def add(sev, where, issue, fix="", by="rule"):
        F.append({"sev": sev, "where": where, "issue": issue, "fix": fix, "by": by})

    leafs = [l for l in lines if not l["is_group"]]
    seen_no = {}
    for l in lines:
        if l["no"] in seen_no:
            add("medium", f"Item {l['no']}", f"Item number {l['no']} appears more than once (row {seen_no[l['no']]} and {l['row']}).",
                "Check the numbering in the Excel file (a number like 3.10 can turn into 3.1).")
        seen_no[l["no"]] = l["row"]

    for l in leafs:
        where = f"Item {l['no']} ({l['item'][:40]})"
        q = l["qty"]
        try:
            ok = float(str(q).replace(",", "")) > 0
        except ValueError:
            ok = False
        if not ok:
            add("high", where, f"Quantity is '{q}': empty, zero or not a number.", "Fix the quantity in the BOM.")
        if not l["unit"]:
            add("low", where, "No unit given.", "")
        if not l["model"]:
            add("medium", where, "No model/specification given.", "Without a model, no library or datasheet match is possible.")
        mm = mismatch(l["item"], l["model"])
        if mm:
            add("high" if mm["close"] else "medium", where,
                f"Item text says '{mm['item_tok']}' but the model column says '{mm['model_tok']}'.",
                "Check which one is right before submitting.")

    # duplicates
    by_full, by_model = {}, {}
    for l in leafs:
        by_full.setdefault((norm(l["item"]), norm(l["model"])), []).append(l)
        for t in _item_model_tokens(l["model"]):
            by_model.setdefault(norm(t), []).append(l)
    flagged = set()
    for k, grp in by_full.items():
        if len(grp) > 1 and k[0]:
            flagged.add(tuple(sorted(g["no"] for g in grp)))
            same_q = len({str(g["qty"]) for g in grp}) == 1
            add("high" if same_q else "medium", "Items " + ", ".join(g["no"] for g in grp),
                f"Same item and model listed {len(grp)} times" + (f", each with quantity {grp[0]['qty']}." if same_q else " with different quantities."),
                "Merge the lines or confirm both are intended.")
    for k, grp in by_model.items():
        nos = tuple(sorted({g["no"] for g in grp}))
        if len(nos) > 1 and nos not in flagged:
            add("low", "Items " + ", ".join(nos), f"The model number '{k}' appears in more than one item.",
                "Fine if intentional (e.g. different sites), otherwise a duplicate.")

    # library / datasheet state
    for l in leafs:
        for part in line_parts.get(l["no"], []):
            where = f"Item {l['no']} ({l['item'][:40]})"
            if not part["products"]:
                add("high", where, f"'{part['text'] or l['item']}' is not in the product library, and no datasheet entry was made for it.",
                    "Add the datasheet PDF to the datasheets folder and re-run, or add the product to the library.")
            for pid in part["products"]:
                if pid in draft_ids:
                    add("medium", where, f"Product '{pid}' was read from a datasheet and is still a draft.",
                        f"Read its library file, then run: python -m bidagent approve {pid}")

    # your own facts
    for l in leafs:
        blob = norm(" ".join([l["item"], l["model"]] + reqs_by_item.get(l["no"], [])))
        for f in facts:
            if any(norm(m) in blob for m in f.get("match", [])) and any(norm(c) in blob for c in f.get("contradicts", [])):
                add("high", f"Item {l['no']} ({l['item'][:40]})", f.get("note", "Conflicts with a known fact."),
                    "Verify and fix the requirement or the product.", by="facts.yaml")
    return F


def llm_review(llm, lines, reqs_by_item, already):
    leafs = [l for l in lines if not l["is_group"]][:120]
    rows = [{"no": l["no"], "item": l["item"][:100], "model": l["model"][:100], "unit": l["unit"],
             "qty": l["qty"], "requirements": reqs_by_item.get(l["no"], [])[:6]} for l in leafs]
    known = [f"{f['where']}: {f['issue'][:90]}" for f in already]
    user = "BOM:\n" + json.dumps(rows, ensure_ascii=False) + "\n\nAlready found (do not repeat):\n" + "\n".join(known[:40])
    d = llm.complete_json("reviewer", SYSTEM, user, max_tokens=2500)
    out = []
    for f in (d.get("findings") if isinstance(d, dict) else None) or []:
        if not isinstance(f, dict) or not f.get("issue"):
            continue
        sev = f.get("severity") if f.get("severity") in SEV else "medium"
        out.append({"sev": sev, "where": collapse(f.get("where", "")), "issue": collapse(f["issue"]) + " (AI opinion: verify)",
                    "fix": collapse(f.get("fix", "")), "by": "ai"})
    return out


def render_report(meta, findings, stats, warnings, usage):
    findings = sorted(findings, key=lambda f: (SEV[f["sev"]], f["where"]))
    n = {s: sum(1 for f in findings if f["sev"] == s) for s in SEV}
    L = [f"# Review report: {meta.get('project_name') or meta.get('name')}", "",
         f"**{n['high']} must-fix**, {n['medium']} to check, {n['low']} minor.", ""]
    L += ["| | |", "|---|---|"] + [f"| {k} | {v} |" for k, v in stats.items()] + [""]
    titles = {"high": "Must fix before sending", "medium": "Please check", "low": "Minor"}
    for s in ("high", "medium", "low"):
        rows = [f for f in findings if f["sev"] == s]
        if not rows:
            continue
        L += [f"## {titles[s]}", ""]
        for f in rows:
            tag = {"rule": "", "facts.yaml": " [your facts.yaml]", "ai": " [AI]"}.get(f["by"], "")
            L.append(f"- **{f['where']}**{tag}: {f['issue']}" + (f" -> {f['fix']}" if f["fix"] else ""))
        L.append("")
    if warnings:
        L += ["## Pipeline notes", ""] + [f"- {w}" for w in warnings] + [""]
    if usage:
        L += ["## AI usage this bid (all runs)", "", "| agent | calls | from cache | prompt tokens | answer tokens |", "|---|---|---|---|---|"]
        for a, u in usage.items():
            L.append(f"| {a} | {u['calls']} | {u['cached']} | {u['prompt']} | {u['completion']} |")
        L.append("")
    return "\n".join(L)
