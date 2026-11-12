"""Orchestrator: runs the steps in order, saves every step's result, and writes the review report.

Re-running a bid is free: each step is skipped when its inputs are unchanged, and every AI
answer is also cached on disk by the LLM client."""
import datetime
import hashlib
import re
import json
import os
from pathlib import Path
from . import bom as bommod
from .agents import compliance as comp
from .agents import datasheet_extractor, reviewer, tender_reader, writer
from . import web_research
from .assembler import assemble
from .datasheets import DatasheetIndex
from .library import Library
from .llm import LLM, LLMUnavailable, load_env, usage_summary
from .util import load_json, read_text_any, save_json, slug, split_models, model_tokens


def _sha(obj):
    return hashlib.sha256(json.dumps(obj, sort_keys=True, default=str, ensure_ascii=False).encode()).hexdigest()[:20]


class Pipeline:
    def __init__(self, cfg, use_llm=True, force=False, log=print):
        self.cfg, self.force, self.log = cfg, force, log
        P = cfg["_paths"]
        load_env(Path(cfg["_root"]) / ".env")
        self.lib = Library(P["library"])
        self.work_root = Path(P["work"])
        self.work_root.mkdir(parents=True, exist_ok=True)
        self.llm = LLM(cfg, self.work_root, enabled=use_llm)
        self.warnings = []

    # ---------------------------------------------------------------- step cache
    def step(self, name, key, fn):
        path = self.wd / "steps" / f"{name}.json"
        k = _sha(key)
        if not self.force and path.exists():
            d = load_json(path)
            if d.get("key") == k:
                self.log(f"  [{name}] unchanged, reused")
                return d["data"]
        self.log(f"  [{name}] running")
        data = fn()
        save_json(path, {"key": k, "data": data})
        return data

    def ai(self, agent):
        return self.llm.available(agent)

    def _lib_sig(self):
        return _sha([(p["id"], p["status"], p["model"], p["blocks"]) for p in self.lib.all()])

    def _google_search_model(self):
        if not self.cfg.get("web_research", {}).get("enabled", True):
            return None
        for step in self.llm.chain("datasheet"):
            if step.get("provider") != "gemini":
                continue
            provider = self.cfg.get("providers", {}).get("gemini", {})
            env = provider.get("api_key_env") or "GEMINI_API_KEY"
            if os.environ.get(env):
                return {"model": step["model"], "api_key": os.environ[env]}
        return None

    # ---------------------------------------------------------------- matching
    def match_lines(self, lines):
        """{no: [{"text": model part, "products": [ids]}]} for every BOM line that is not a group."""
        out = {}
        for l in lines:
            if l["is_group"]:
                continue
            parts = split_models(l["model"]) or ([l["model"]] if l["model"] else [])
            res = []
            for part in parts or [""]:
                hits = self.lib.match(part) or (self.lib.match(l["item"]) if len(parts) <= 1 else [])
                res.append({"text": part, "products": [h["id"] for h in hits]})
            out[l["no"]] = res
        return out

    # ---------------------------------------------------------------- the run
    def run(self, bom_path, tender_paths=(), name=None, client="", tender_no="", location="", out=None):
        cfg, P = self.cfg, self.cfg["_paths"]
        name = name or Path(bom_path).stem
        self.wd = self.work_root / slug(name)
        meta = {"name": name, "client": client, "tender_no": tender_no, "location": location,
                "date": datetime.date.today().strftime("%d %B %Y")}
        self.log(f"Bid: {name}")

        # 1 BOM (no AI)
        lines = bommod.to_dicts(bommod.parse_bom(bom_path))
        save_json(self.wd / "bom.json", lines)
        leafs = [l for l in lines if not l["is_group"]]
        self.log(f"  BOM: {len(lines)} rows, {len(leafs)} items to offer")

        # 2 datasheet index (no AI, cached per PDF)
        ds = DatasheetIndex(P["datasheets"], self.work_root / "datasheet_index.json",
                            cfg["datasheets"]["max_pages_per_model"], cfg["datasheets"]["max_chars"])
        n_pdf, n_pg = ds.build()
        self.warnings += ds.warnings
        self.log(f"  datasheets: {n_pdf} PDFs, {n_pg} pages indexed")

        # 3 tender
        tender = {"summary": "", "purpose": [], "objectives": [], "scope": [], "requirements": [],
                  "general_requirements": [], "project_name": "", "client": "", "tender_no": "", "location": ""}
        if tender_paths:
            text = "\n\n".join(f"=== {Path(t).name} ===\n{read_text_any(t)}" for t in tender_paths)
            use_ai = self.ai("tender")

            def do_tender():
                if use_ai:
                    t, w = tender_reader.read_tender(self.llm, text, lines, cfg["tender"]["chunk_chars"],
                                                     cfg["tender"]["max_chunks"])
                else:
                    t, w = tender_reader.read_tender_offline(text)
                return {"tender": t, "warnings": w}
            r = self.step("tender", {"text": text, "ai": use_ai, "lines": [l["no"] for l in lines],
                                      "chunk_chars": cfg["tender"]["chunk_chars"],
                                      "max_chunks": cfg["tender"]["max_chunks"]}, do_tender)
            tender = r["tender"]
            self.warnings += r["warnings"]
        for k in ("client", "tender_no", "location"):
            meta[k] = meta[k] or tender.get(k, "")
        meta["project_name"] = tender.get("project_name") or name
        reqs_by_item, loose = tender_reader.assign_requirements(tender, lines)
        if loose:
            self.warnings.append(f"{len(loose)} tender requirement(s) could not be tied to a BOM item "
                                 f"and are not in the compliance table: " + " | ".join(x[:80] for x in loose[:5]))

        # 4 match BOM -> library, 5 research BOM products and fill gaps
        parts = self.match_lines(lines)
        missing = sorted({(p["text"], l["item"]) for l in leafs for p in parts[l["no"]]
                          if not p["products"] and p["text"]})
        targets = sorted({(p["text"], l["item"]) for l in leafs for p in parts[l["no"]] if p["text"]})
        google_search = self._google_search_model()
        if targets and (google_search or (missing and ds.pages and self.ai("datasheet"))):
            def do_extract():
                made, warns = [], []
                for model, item in targets:
                    for tok in ([t for t in model_tokens(model) if not re.fullmatch(r"(19|20)\d\d", t)][:1] or [model]):
                        if (model, item) in missing and ds.pages and self.ai("datasheet"):
                            try:
                                prod, w = datasheet_extractor.extract_product(self.llm, ds, tok, item)
                            except LLMUnavailable as e:
                                w = [f"{tok}: AI unavailable ({e})"]
                            warns += w
                            if prod and self.lib.get(prod["id"]) is None:
                                self.lib.save(prod)
                                made.append(prod["id"])
                        if google_search:
                            reqs = []
                            for line in leafs:
                                if line["item"] == item and model in line["model"]:
                                    reqs.extend(reqs_by_item.get(line["no"], []))
                            web_prod, w = web_research.research_product(
                                google_search["api_key"], google_search["model"], tok, item,
                                reqs)
                            warns += w
                            if web_prod:
                                web_prod["id"] = "web-" + slug(tok, maxlen=42)
                                self.lib.save(web_prod)
                                made.append(web_prod["id"])
                return {"made": made, "warnings": warns}
            r = self.step("extract", {"targets": targets, "missing": missing, "pages": n_pg,
                                       "google_model": google_search["model"] if google_search else "",
                                       "requirements": reqs_by_item}, do_extract)
            self.warnings += r["warnings"]
            if r["made"]:
                self.log(f"  researched product entries: {', '.join(r['made'])}")
            parts = self.match_lines(lines)
        if targets and not google_search:
            self.warnings.append("Google Search research was not run. Add GEMINI_API_KEY to .env to research current BOM products online.")
            if missing:
                self.log(f"  {len(missing)} product(s) not in library; add a Gemini key for Google Search research")

        # 6 intro sections
        sections = list(cfg["sections"]) + list(cfg["extra_sections"])
        style = self.lib.style()

        def do_write():
            if self.ai("writer"):
                try:
                    s, w = writer.write_sections(self.llm, style, tender, lines, sections, meta)
                    return {"sections": s, "warnings": w}
                except Exception as e:
                    self.warnings.append(f"Writer failed ({e}); placeholders used.")
            s, w = writer.write_sections_offline(tender, sections)
            return {"sections": s, "warnings": w}
        r = self.step("write", {"t": tender, "m": meta, "s": sections, "style": style,
                                "ai": self.ai("writer"), "bom": [l["item"] for l in lines]}, do_write)
        intro = r["sections"]
        self.warnings += r["warnings"]

        # 7 compliance
        rows_by_item = {}
        if cfg["compliance"]["enabled"] and reqs_by_item:
            def do_comply():
                rows, warns = {}, []
                for l in leafs:
                    reqs = reqs_by_item.get(l["no"])
                    if not reqs:
                        continue
                    prods = [self.lib.get(pid) for p in parts[l["no"]] for pid in p["products"]]
                    prods = [p for p in prods if p]
                    ai = self.llm if self.ai("compliance") else None
                    rs, w = comp.check_item(ai, reqs, prods, ds if ds.pages else None,
                                            cfg["compliance"]["datasheet_second_pass"])
                    rows[l["no"]] = rs
                    warns += w
                return {"rows": rows, "warnings": warns}
            r = self.step("compliance", {"reqs": reqs_by_item, "parts": parts, "lib": self._lib_sig(),
                                         "ai": self.ai("compliance"), "ds": n_pg}, do_comply)
            rows_by_item = r["rows"]
            self.warnings += r["warnings"]

        # 8 review
        facts = reviewer.load_facts(P["facts"])
        drafts = {p["id"] for p in self.lib.all() if p.get("status") == "draft"}
        findings = reviewer.static_checks(lines, parts, reqs_by_item, facts, drafts)
        for no, rows in rows_by_item.items():            # compliance gaps are findings too
            l = next(x for x in lines if x["no"] == no)
            for x in rows:
                where = f"Item {no} ({l['item'][:40]})"
                if x["status"] == "not_complies":
                    findings.append({"sev": "high", "where": where, "by": "rule", "fix": "Offer another model or raise a deviation.",
                                     "issue": f"Does not meet: \"{x['req'][:110]}\" (offered: {x['offered'][:80]})"})
                elif x["status"] == "partial":
                    findings.append({"sev": "medium", "where": where, "by": "rule", "fix": "Check the datasheet.",
                                     "issue": f"Only partly meets: \"{x['req'][:110]}\" {x['note']}"})
                elif x["status"] in ("not_found", "to_confirm"):
                    findings.append({"sev": "low", "where": where, "by": "rule", "fix": "Confirm by hand.",
                                     "issue": f"No evidence found for: \"{x['req'][:110]}\""})
        if self.ai("reviewer"):
            def do_review():
                try:
                    return reviewer.llm_review(self.llm, lines, reqs_by_item, findings)
                except Exception as e:
                    self.warnings.append(f"AI review failed: {e}")
                    return []
            findings += self.step("review", {"lines": lines, "reqs": reqs_by_item, "n": len(findings)}, do_review)
        save_json(self.wd / "findings.json", findings)

        # 9 assemble the Word file
        out = Path(out) if out else Path(P["output"]) / f"{slug(name)}.docx"
        template = self.lib.root / "templates" / cfg["template"]["name"]
        if not template.exists():
            raise FileNotFoundError(f"No template at {template}. Run: python -m bidagent ingest <your finished bid.docx>")
        stats, w = assemble(cfg, self.lib, template, out, meta, lines, parts, intro, rows_by_item)
        self.warnings += w
        report = out.with_name(out.stem + "_review.md")
        summary = {"BOM items": len(leafs), "Products written": stats["products_written"],
                   "Missing-product placeholders": stats["placeholders"],
                   "Compliance rows": stats["compliance_rows"], "Draft products used": len(drafts)}
        report.write_text(reviewer.render_report(meta, findings, summary, self.warnings,
                                                 usage_summary(self.work_root)), "utf-8")
        self.log(f"  wrote {out}\n  wrote {report}")
        return {"docx": str(out), "report": str(report), "stats": summary, "findings": findings}
