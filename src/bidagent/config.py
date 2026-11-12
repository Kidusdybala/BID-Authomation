"""Config loading. One YAML file, defaults filled in, paths resolved against the project folder."""
import copy
from pathlib import Path
import yaml

DEFAULTS = {
    "paths": {"library": "library", "datasheets": "datasheets", "work": "work",
              "output": "output", "facts": "knowledge/facts.yaml"},
    "template": {
        "name": "template.docx",
        # Top-level (Heading 1) sections of your template that are copied into every bid.
        # Everything before the first Heading 1 (cover page, contents) is always kept.
        "keep_sections": ["Company Background"],
        # Extra find/replace on the kept template text, e.g. {"Old Project Name": "{{PROJECT_NAME}}"}
        "replace": {},
        "table_style": None,         # None = reuse the style of the last table in your template
    },
    "titles": {"intro": "Project Introduction", "technical": "Technical Proposal",
               "compliance": "Compliance Statement", "bom": "Bill of Materials"},
    "sections": [
        {"title": "Project Overview",
         "guidance": "What the project is, who the client is, where, and what the bid delivers overall."},
        {"title": "Purpose of the Project",
         "guidance": "Why the client is doing this project (the problem or need), from the tender."},
        {"title": "Objectives of the Project",
         "guidance": "The client's goals, as 4-6 short sentences or a list written as sentences."},
        {"title": "Scope of the Project",
         "guidance": "What is included: the systems/equipment groups in the BOM, supply, install, "
                     "configuration, testing, training. Only what the tender or BOM states."},
    ],
    "extra_sections": [],            # same shape as `sections`: bid-specific text you want written
    "compliance": {"enabled": True, "datasheet_second_pass": True},
    "datasheets": {"max_pages_per_model": 6, "max_chars": 12000},
    "tender": {"chunk_chars": 14000, "max_chunks": 8},
    "draft_products": "include",     # include | skip  (products read from a datasheet, not yet approved)
    "providers": {},
    "agents": {},
}


def _merge(base, over):
    out = copy.deepcopy(base)
    for k, v in (over or {}).items():
        if isinstance(v, dict) and isinstance(out.get(k), dict) and k not in ("providers", "agents"):
            out[k] = _merge(out[k], v)
        else:
            out[k] = v
    return out


def load_config(path="config.yaml"):
    path = Path(path).resolve()
    raw = yaml.safe_load(path.read_text("utf-8")) if path.exists() else {}
    cfg = _merge(DEFAULTS, raw or {})
    root = path.parent
    cfg["_root"] = str(root)
    cfg["_paths"] = {k: (root / v) for k, v in cfg["paths"].items()}
    return cfg
