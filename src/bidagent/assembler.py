"""Builds the final Word file from YOUR template. No AI.

 kept from the template : cover/contents (everything before the first Heading 1) and the
                          Heading-1 sections named in template.keep_sections (Company Background)
 generated              : intro sections, Technical Proposal (library text + pictures),
                          Compliance Statement, Bill of Materials (last, as in your V04)
Your list bullets, caption style, heading styles and table style are reused from the template.
"""
import copy
import re
from pathlib import Path
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_COLOR_INDEX
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt
from .util import collapse, norm

STATUS_FILL = {"complies": "E2F0D9", "partial": "FFF2CC", "not_complies": "F8CBAD",
               "not_found": "E7E6E6", "to_confirm": "E7E6E6"}


# ------------------------------------------------------------------ helpers
def _sid(el):
    ppr = el.find(qn("w:pPr"))
    ps = ppr.find(qn("w:pStyle")) if ppr is not None else None
    return ps.get(qn("w:val")) if ps is not None else ""


def _text(el):
    return collapse("".join(t.text or "" for t in el.iter(qn("w:t"))))


def _has_style(doc, name):
    try:
        doc.styles[name]
        return True
    except KeyError:
        return False


def _fill(cell, hex_):
    tcPr = cell._tc.get_or_add_tcPr()
    for old in tcPr.findall(qn("w:shd")):
        tcPr.remove(old)
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_)
    tcPr.append(shd)


def _field(p, instr, result="", dirty=False):
    """Complex field as plain runs (so paragraph.text still shows the result)."""
    def run(child):
        r = OxmlElement("w:r")
        r.append(child)
        p._p.append(r)
        return r
    b = OxmlElement("w:fldChar")
    b.set(qn("w:fldCharType"), "begin")
    if dirty:
        b.set(qn("w:dirty"), "true")
    run(b)
    it = OxmlElement("w:instrText")
    it.set(qn("xml:space"), "preserve")
    it.text = f" {instr} "
    run(it)
    s = OxmlElement("w:fldChar")
    s.set(qn("w:fldCharType"), "separate")
    run(s)
    t = OxmlElement("w:t")
    t.set(qn("xml:space"), "preserve")
    t.text = result
    run(t)
    e = OxmlElement("w:fldChar")
    e.set(qn("w:fldCharType"), "end")
    run(e)


# ------------------------------------------------------------------ template handling
def _trim_template(doc, keep_titles, warnings):
    """Drop every Heading-1 section that is not in keep_titles. Returns the list/table styles found."""
    body = doc.element.body
    kids = [c for c in body.iterchildren() if c.tag != qn("w:sectPr")]
    has_h1 = any(c.tag == qn("w:p") and _sid(c) == "Heading1" for c in kids)
    if not has_h1:
        warnings.append("Template has no Heading 1 paragraphs: the whole template was kept and the "
                        "new content was added after it.")
        return
    keep_l = {k.lower() for k in keep_titles}
    found, state, drop = set(), "keep", []
    for el in kids:
        if el.tag == qn("w:p") and _sid(el) == "Heading1":
            title = _text(el).lower()
            state = "keep" if title in keep_l else "drop"
            if state == "keep":
                found.add(title)
        if state == "drop":
            drop.append(el)
    for k in keep_l - found:
        warnings.append(f'template.keep_sections: no Heading 1 called "{k}" in the template.')
    for el in drop:
        body.remove(el)
    # remove pictures that only the dropped sections used (keeps the file small)
    xml = body.xml
    used = set(re.findall(r'r:(?:embed|id|link)="(rId\d+)"', xml))
    for rid, rel in list(doc.part.rels.items()):
        if rel.reltype.endswith("/image") and rid not in used:
            doc.part.rels.pop(rid, None)


def _capture_list_ppr(doc):
    for el in doc.element.body.iter(qn("w:p")):
        if _sid(el) == "ListParagraph" and el.find(qn("w:pPr")) is not None \
                and el.find(qn("w:pPr")).find(qn("w:numPr")) is not None:
            return copy.deepcopy(el.find(qn("w:pPr")))
    return None


def _replace_in_paragraph(p_el, mapping):
    ts = list(p_el.iter(qn("w:t")))
    if not ts:
        return 0
    full = "".join(t.text or "" for t in ts)
    new = full
    for k, v in mapping.items():
        new = new.replace(k, v)
    if new == full:
        return 0
    ts[0].text = new
    ts[0].set(qn("xml:space"), "preserve")
    for t in ts[1:]:
        t.text = ""
    return 1


def _apply_placeholders(doc, mapping):
    n = 0
    for p in doc.element.body.iter(qn("w:p")):
        n += _replace_in_paragraph(p, mapping)
    for s in doc.sections:
        for part in (s.header, s.footer, s.first_page_header, s.first_page_footer):
            try:
                for p in part._element.iter(qn("w:p")):
                    n += _replace_in_paragraph(p, mapping)
            except Exception:
                pass
    return n


# ------------------------------------------------------------------ the builder
class Builder:
    def __init__(self, template, lib, cfg):
        self.doc = Document(str(template))
        self.lib, self.cfg = lib, cfg
        self.fig = 0
        self.warnings = []
        self.list_ppr = _capture_list_ppr(self.doc)
        ts = cfg["template"].get("table_style")
        if not ts and self.doc.tables:
            try:
                ts = self.doc.tables[-1].style.name
            except Exception:
                ts = None
        self.table_style = ts if ts and _has_style(self.doc, ts) else ("Table Grid" if _has_style(self.doc, "Table Grid") else None)

    # --- text blocks
    def heading(self, text, level, page_break=False):
        name = f"Heading {level}"
        if _has_style(self.doc, name):
            p = self.doc.add_paragraph(text, style=name)
        else:
            p = self.doc.add_paragraph()
            p.add_run(text).bold = True
        if page_break:
            p.paragraph_format.page_break_before = True
        return p

    def para(self, text, bold=False, italic=False, mark=False):
        p = self.doc.add_paragraph()
        r = p.add_run(text)
        r.bold, r.italic = bold, italic
        if mark or "[TO " in text:
            r.font.highlight_color = WD_COLOR_INDEX.YELLOW
        return p

    def bullet(self, label, value):
        if self.list_ppr is not None and _has_style(self.doc, "List Paragraph"):
            p = self.doc.add_paragraph(style="List Paragraph")
            ppr = p._p.find(qn("w:pPr"))
            if ppr is not None:
                p._p.remove(ppr)
            p._p.insert(0, copy.deepcopy(self.list_ppr))
        elif _has_style(self.doc, "List Bullet"):
            p = self.doc.add_paragraph(style="List Bullet")
        else:
            p = self.doc.add_paragraph()
            p.add_run("• ")
        if label:
            p.add_run(label + ": ").bold = True
        p.add_run(value)
        return p

    def picture(self, path, width_cm, caption):
        try:
            p = self.doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.keep_with_next = True
            p.add_run().add_picture(str(path), width=Cm(min(float(width_cm or 12), 15.5)))
        except Exception as e:
            self.warnings.append(f"Picture {Path(path).name} could not be inserted: {e}")
            return
        self.fig += 1
        cp = self.doc.add_paragraph(style="Caption") if _has_style(self.doc, "Caption") else self.doc.add_paragraph()
        cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        cp.add_run("Figure ")
        _field(cp, "SEQ Figure \\* ARABIC", str(self.fig))
        cp.add_run(" " + caption)

    def table(self, headers, rows, widths_cm, fills=None, bold_rows=()):
        t = self.doc.add_table(rows=1, cols=len(headers))
        if self.table_style:
            t.style = self.table_style
        t.autofit = False
        for i, h in enumerate(headers):
            c = t.rows[0].cells[i]
            c.text = ""
            r = c.paragraphs[0].add_run(h)
            r.bold, r.font.size = True, Pt(9)
            _fill(c, "D9E2F3")
        trPr = t.rows[0]._tr.get_or_add_trPr()
        trPr.append(OxmlElement("w:tblHeader"))
        for ri, row in enumerate(rows):
            cells = t.add_row().cells
            for i, v in enumerate(row):
                cells[i].text = ""
                r = cells[i].paragraphs[0].add_run("" if v is None else str(v))
                r.font.size = Pt(9)
                r.bold = ri in bold_rows
                if "[TO " in str(v):
                    r.font.highlight_color = WD_COLOR_INDEX.YELLOW
            if fills and fills[ri]:
                for ci, hexv in fills[ri].items():
                    _fill(cells[ci], hexv)
        for i, w in enumerate(widths_cm):
            t.columns[i].width = Cm(w)
        for row in t.rows:
            for i, w in enumerate(widths_cm):
                row.cells[i].width = Cm(w)
        self.doc.add_paragraph()
        return t

    # --- sections
    def toc_if_missing(self):
        has = any("TOC" in (i.text or "") and "\\o" in (i.text or "") for i in self.doc.element.body.iter(qn("w:instrText")))
        if has:
            return
        self.para("Table of Contents", bold=True)
        p = self.doc.add_paragraph()
        _field(p, 'TOC \\o "1-3" \\h \\z \\u', "Right-click here and choose Update Field to build the contents.", dirty=True)

    def intro(self, title, sections):
        self.heading(title, 1, page_break=True)
        for t, paras in sections.items():
            self.heading(t, 2)
            for x in paras:
                self.para(x)

    def product_block(self, b):
        self.heading(b["heading"], 3)
        for d in b.get("description", []):
            self.para(d)
        feats = [f for f in b.get("features", []) if f.get("verified", True)]
        if feats:
            self.para("Key Features", bold=True)
            for f in feats:
                self.bullet(f.get("label", ""), f.get("value", ""))
            sources = list(dict.fromkeys(f.get("source", "") for f in feats if f.get("source")))
            if sources:
                self.para("Sources", bold=True)
                for source in sources:
                    self.bullet("", source)
        img = self.lib.image_path(b.get("image"))
        if img and img.exists():
            self.picture(img, b.get("image_width_cm"), b.get("caption") or b["heading"])

    def technical(self, title, lines, line_parts, stats):
        self.heading(title, 1, page_break=True)
        done = set()
        for l in lines:
            if l["level"] == 0:
                self.heading(l["item"], 2)
            if l["is_group"]:
                continue
            for part in line_parts.get(l["no"], []):
                if not part["products"]:
                    self.heading(l["item"], 3)
                    self.para(f"[TO COMPLETE: no library entry or datasheet for "
                              f"{part['text'] or l['item']}. Add the datasheet and re-run.]")
                    stats["placeholders"] += 1
                    continue
                for pid in part["products"]:
                    if pid in done:
                        continue
                    done.add(pid)
                    prod = self.lib.get(pid)
                    if prod is None:
                        continue
                    if prod.get("status") == "draft" and self.cfg["draft_products"] == "skip":
                        self.heading(l["item"], 3)
                        self.para(f"[TO COMPLETE: {prod['model']} is a draft product, approve it in the library.]")
                        stats["placeholders"] += 1
                        continue
                    for b in prod["blocks"]:
                        self.product_block(b)
                    stats["products_written"] += 1

    def compliance(self, title, lines, rows_by_item, stats):
        if not rows_by_item:
            return
        self.heading(title, 1, page_break=True)
        by_no = {l["no"]: l for l in lines}
        for no, rows in rows_by_item.items():
            l = by_no[no]
            self.heading(f"{no}  {l['item']}", 2)
            data, fills = [], []
            for i, r in enumerate(rows, 1):
                ref = r["source"] + (f" ({r['note']})" if r["note"] else "")
                data.append([i, r["req"], r["offered"] or "[TO CONFIRM]", r["label"], ref])
                fills.append({3: STATUS_FILL.get(r["status"], "E7E6E6")})
            self.table(["No.", "Requirement", "Offered", "Compliance", "Reference"], data,
                       [1.0, 5.2, 5.6, 2.4, 2.8], fills)
            stats["compliance_rows"] += len(rows)

    def bom(self, title, lines):
        self.heading(title, 1, page_break=True)
        data = [[l["no"], l["item"], l["model"], l["unit"], l["qty"]] for l in lines]
        bold = [i for i, l in enumerate(lines) if l["is_group"]]
        self.table(["No", "Item", "Model", "Unit", "Quantity"], data, [1.3, 6.2, 6.2, 1.6, 1.7], bold_rows=bold)


def assemble(cfg, lib, template, out_path, meta, lines, line_parts, intro, rows_by_item):
    """Returns (stats, warnings)."""
    b = Builder(template, lib, cfg)
    doc = b.doc
    _trim_template(doc, cfg["template"]["keep_sections"], b.warnings)

    mapping = {"{{PROJECT_NAME}}": meta.get("project_name", ""), "{{CLIENT}}": meta.get("client", ""),
               "{{TENDER_NO}}": meta.get("tender_no", ""), "{{DATE}}": meta.get("date", ""),
               "{{LOCATION}}": meta.get("location", "")}
    for old, new in (cfg["template"].get("replace") or {}).items():
        for k, v in mapping.items():
            new = new.replace(k, v)
        mapping[old] = new
    n = _apply_placeholders(doc, mapping)
    if n == 0:
        b.warnings.append("Cover page was not updated: the template has no {{PROJECT_NAME}}-style "
                          "tokens and template.replace is empty. Add tokens to your template.")

    stats = {"products_written": 0, "placeholders": 0, "compliance_rows": 0}
    t = cfg["titles"]
    b.toc_if_missing()
    b.intro(t["intro"], intro)
    b.technical(t["technical"], lines, line_parts, stats)
    if cfg["compliance"]["enabled"]:
        b.compliance(t["compliance"], lines, rows_by_item, stats)
    b.bom(t["bom"], lines)

    # ask Word to refresh the contents + figure numbers when the file is opened
    st = doc.settings.element
    if st.find(qn("w:updateFields")) is None:
        u = OxmlElement("w:updateFields")
        u.set(qn("w:val"), "true")
        st.append(u)
    doc.core_properties.title = meta.get("project_name") or meta.get("name", "")
    Path(out_path).parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out_path))
    return stats, b.warnings
