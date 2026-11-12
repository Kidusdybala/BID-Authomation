"""Learns from one of YOUR finished bid documents (no AI, no tokens):
  * every product section (heading, description, Key Features, picture) -> product library
  * your wording of Overview/Purpose/Objectives/Scope -> style examples for the writer agent
  * the .docx itself -> template (cover, TOC, company background, table look)
"""
import io
import re
import shutil
from pathlib import Path
from docx import Document
from PIL import Image
from .util import norm, slug, collapse, model_tokens
from .library import Library

STYLE_SECTIONS = ["Project Overview", "Purpose of the Project", "Objectives of the Project",
                  "Scope of the Project"]


def _sid(p):
    ppr = p.find("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pPr")
    if ppr is None:
        return ""
    ps = ppr.find("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pStyle")
    return ps.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val") if ps is not None else ""


def _bom_from_doc(doc):
    """(item, model) pairs from the last table of the document."""
    if not doc.tables:
        return []
    out = []
    for r in doc.tables[-1].rows:
        tcs = r._tr.tc_lst
        texts = []
        for tc in tcs:
            texts.append(collapse("".join(t.text or "" for t in tc.iter(
                "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t"))))
        if len(texts) >= 5 and texts[2] and texts[0] and texts[0].lower() != "no":
            out.append((texts[1], texts[2]))
    return out


def _model_token(model):
    toks = model_tokens(model)
    return toks[0] if toks else collapse(model)


def ingest_docx(docx_path, lib: Library, mapping=None, template_name="template.docx"):
    mapping = {k.lower(): v for k, v in (mapping or {}).items()}
    doc = Document(str(docx_path))
    body = list(doc.element.body.iterchildren())
    bom = _bom_from_doc(doc)
    models = [(item, _model_token(m), m) for item, m in bom]

    blocks, cur, mode = [], None, None
    h2 = ""
    texts_by_h2 = {}
    h1 = ""
    h1_text = {}
    from docx.text.paragraph import Paragraph
    for el in body:
        if not el.tag.endswith("}p"):
            continue
        p = Paragraph(el, doc)
        sid, txt = _sid(el), collapse(p.text)
        if sid == "Heading1":
            h1, cur = txt, None
            continue
        if sid == "Heading2":
            h2, cur = txt, None
            continue
        if sid == "Heading3":
            cur = {"heading": txt, "description": [], "features": [], "image": None,
                   "caption": "", "h2": h2}
            blocks.append(cur)
            mode = "desc"
            continue
        if txt and sid not in ("TOC1", "TOC2", "TOC3", "TableofFigures", "TOCHeading"):
            if h1:
                h1_text.setdefault(h1, []).append(txt)
            if h2 and cur is None:
                texts_by_h2.setdefault(h2, []).append(txt)
        if cur is None:
            continue
        blips = el.xpath(".//a:blip/@r:embed")
        if blips and cur["image"] is None:
            part = doc.part.related_parts[blips[0]]
            ext = (part.partname.ext or "png").lower()
            cx = el.xpath(".//wp:extent/@cx")
            cur["image"] = {"blob": part.blob, "ext": ext,
                            "width_cm": round(int(cx[0]) / 360000, 1) if cx else 12}
            continue
        if sid == "Caption":
            cur["caption"] = re.sub(r"^Figure\s*\d+\s*", "", txt).strip()
            continue
        if not txt:
            continue
        if txt.lower().startswith("key features"):
            mode = "features"
            continue
        if mode == "features" and (sid == "ListParagraph" or el.xpath(".//w:numPr")):
            label, _, value = txt.partition(":")
            if not value:
                label, value = "", txt
            cur["features"].append({"label": label.strip(), "value": value.strip(" ,;"),
                                    "source": f"ingested from {Path(docx_path).name}",
                                    "verified": True})
        elif mode == "desc":
            cur["description"].append(txt)

    # ---- assign blocks to products -------------------------------------------------
    created, unmapped = {}, []
    for b in blocks:
        text = norm(b["heading"] + " " + " ".join(b["description"]) + " " +
                    " ".join(f["value"] for f in b["features"]))
        target = None
        key = b["heading"].lower()
        if key in mapping:
            target = mapping[key]
        elif b["features"]:   # sections without "Key Features" are solution text, not products
            head = norm(b["heading"] + " " + " ".join(b["description"][:1]))
            best = None
            for item, tok_, full in models:
                pos = head.find(norm(tok_)) if len(norm(tok_)) >= 4 else -1
                if pos != -1 and (best is None or pos < best[0]):
                    best = (pos, tok_)
            target = best[1] if best else None
        if not target or not (b["description"] or b["features"]):
            unmapped.append(b["heading"])
            continue
        prod = created.get(target)
        if prod is None:
            item = next((i for i, t, f in models if norm(t) == norm(target)), "")
            prod = Library.new_product(target, name=item, status="approved")
            prod["source"] = {"ingested_from": Path(docx_path).name}
            created[target] = prod
        img_name = None
        if b["image"]:
            img_name = f"{prod['id']}-{len(prod['blocks']) + 1}.{b['image']['ext']}"
            _save_image(b["image"]["blob"], lib.image_path(img_name))
        prod["blocks"].append({"heading": b["heading"], "description": b["description"],
                               "features": b["features"], "image": img_name,
                               "image_width_cm": b["image"]["width_cm"] if b["image"] else None,
                               "caption": b["caption"] or b["heading"]})
    for prod in created.values():
        lib.save(prod)

    # ---- style examples + template -------------------------------------------------
    style = lib.style()
    for h in STYLE_SECTIONS:
        if texts_by_h2.get(h):
            style[h] = "\n\n".join(texts_by_h2[h])
    for h in list(texts_by_h2):
        if h not in STYLE_SECTIONS and sum(len(t) for t in texts_by_h2[h]) > 300 and "group_intro" not in style.get(h, ""):
            style.setdefault("group_intro_examples", {})[h] = "\n\n".join(texts_by_h2[h])[:1500]
    if h1_text.get("Company Background"):
        style["company_background"] = "\n\n".join(h1_text["Company Background"])[:2500]
    lib.save_style(style)
    dest = lib.root / "templates" / template_name
    if Path(docx_path).resolve() != dest.resolve():
        shutil.copyfile(docx_path, dest)
    return {"products": len(created), "blocks": sum(len(p["blocks"]) for p in created.values()),
            "unmapped": unmapped, "template": str(dest)}


def _save_image(blob, path, max_px=1400):
    try:
        im = Image.open(io.BytesIO(blob))
        im.load()
        if max(im.size) > max_px:
            im.thumbnail((max_px, max_px))
        fmt = (im.format or "PNG")
        if path.suffix.lower() in (".jpg", ".jpeg"):
            im.convert("RGB").save(path, "JPEG", quality=85)
        else:
            im.save(path)
    except Exception:
        path.write_bytes(blob)
