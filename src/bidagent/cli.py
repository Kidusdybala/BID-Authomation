"""Command line.  python -m bidagent --help"""
import argparse
import json
import sys
from pathlib import Path
from .config import load_config
from .library import Library
from .llm import LLM, load_env, usage_summary


def main(argv=None):
    ap = argparse.ArgumentParser(prog="bidagent", description="Bid document builder")
    ap.add_argument("-c", "--config", default="config.yaml")
    sub = ap.add_subparsers(dest="cmd", required=True)

    p = sub.add_parser("ingest", help="learn from one of your finished bids (.docx)")
    p.add_argument("docx")
    p.add_argument("--map", help="JSON file {section heading: model} to fix product assignment")

    sub.add_parser("index", help="read the datasheet PDFs and show what was found")
    sub.add_parser("library", help="list products in the library")
    p = sub.add_parser("approve", help="mark draft products as approved")
    p.add_argument("ids", nargs="+")
    sub.add_parser("usage", help="AI calls and tokens used so far")
    sub.add_parser("doctor", help="check keys and which models each agent will use")

    p = sub.add_parser("build", help="build a bid from a BOM (and tender)")
    p.add_argument("--bom", required=True)
    p.add_argument("--tender", nargs="*", default=[], help="tender files (pdf/docx/txt/xlsx)")
    p.add_argument("--name")
    p.add_argument("--client", default="")
    p.add_argument("--tender-no", default="")
    p.add_argument("--location", default="")
    p.add_argument("--out")
    p.add_argument("--no-llm", action="store_true", help="run with no AI at all")
    p.add_argument("--force", action="store_true", help="ignore saved step results")

    a = ap.parse_args(argv)
    cfg = load_config(a.config)
    load_env(Path(cfg["_root"]) / ".env")
    P = cfg["_paths"]

    if a.cmd == "ingest":
        from .ingest import ingest_docx
        mapping = json.loads(Path(a.map).read_text("utf-8")) if a.map else None
        r = ingest_docx(a.docx, Library(P["library"]), mapping, cfg["template"]["name"])
        print(f"{r['products']} products, {r['blocks']} sections, template saved to {r['template']}")
        if r["unmapped"]:
            print("Sections not treated as products:", "; ".join(r["unmapped"]))
    elif a.cmd == "index":
        from .datasheets import DatasheetIndex
        ds = DatasheetIndex(P["datasheets"], Path(P["work"]) / "datasheet_index.json")
        Path(P["work"]).mkdir(parents=True, exist_ok=True)
        n, pg = ds.build()
        print(f"{n} PDFs, {pg} pages")
        for w in ds.warnings:
            print("WARNING:", w)
    elif a.cmd == "library":
        for p in Library(P["library"]).all():
            nf = sum(len(b["features"]) for b in p["blocks"])
            print(f"{p['status']:9} {p['id']:28} {nf:3} features  {p['name'][:40]}")
    elif a.cmd == "approve":
        lib = Library(P["library"])
        for i in a.ids:
            p = lib.get(i)
            if not p:
                print("not found:", i)
                continue
            p["status"] = "approved"
            lib.save(p)
            print("approved", i)
    elif a.cmd == "usage":
        for k, v in usage_summary(P["work"]).items():
            print(f"{k:12} calls={v['calls']} cached={v['cached']} prompt={v['prompt']} answer={v['completion']}")
    elif a.cmd == "doctor":
        llm = LLM(cfg, Path(P["work"]))
        for agent in ("tender", "datasheet", "writer", "compliance", "reviewer"):
            chain = llm.chain(agent)
            print(f"{agent:11} {'READY' if llm.available(agent) else 'NO KEY'}  " +
                  " -> ".join(f"{s['provider']}/{s['model']}" for s in chain))
    elif a.cmd == "build":
        from .pipeline import Pipeline
        pl = Pipeline(cfg, use_llm=not a.no_llm, force=a.force)
        r = pl.run(a.bom, a.tender, a.name, a.client, a.tender_no, a.location, a.out)
        n = {s: sum(1 for f in r["findings"] if f["sev"] == s) for s in ("high", "medium", "low")}
        print(f"\nDone. {n['high']} must-fix, {n['medium']} to check. Read the review report first:\n  {r['report']}")


if __name__ == "__main__":
    sys.exit(main())
