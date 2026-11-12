"""BOM reader. Pure Python, no AI. Copes with different layouts:
   V04-style  : No | Item | Model | Unit | Quantity   (with 3 / 3.1 / 3.2 hierarchy)
   MOD-style  : No. | Required Item | <blank header = model/spec> | UoM | Qty.
"""
import re
from dataclasses import dataclass, asdict, field
import openpyxl
from .util import collapse


@dataclass
class BomLine:
    no: str
    item: str
    model: str
    unit: str
    qty: object
    level: int = 0
    is_group: bool = False
    row: int = 0
    parent: str = ""


def _cell(v):
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return collapse(v)


def parse_bom(path, sheet=None):
    wb = openpyxl.load_workbook(str(path), data_only=True)
    ws = wb[sheet] if sheet else wb.active
    rows = [list(r) for r in ws.iter_rows(values_only=True)]

    hdr = None
    for i, r in enumerate(rows[:40]):
        cells = [collapse(c).lower() for c in r]
        has_no = any(re.fullmatch(r"(no\.?|#|s/?n|item no\.?|sl\.? ?no\.?)", c) for c in cells)
        has_item = any(("item" in c or "description" in c) and not re.fullmatch(r"item no\.?", c)
                       for c in cells)
        if has_no and has_item:
            hdr = i
            break
    if hdr is None:
        raise ValueError("Could not find the BOM header row (need a 'No' column and an 'Item' column).")

    cells = [collapse(c).lower() for c in rows[hdr]]
    col = {}
    for j, c in enumerate(cells):
        if "no" not in col and re.fullmatch(r"(no\.?|#|s/?n|item no\.?|sl\.? ?no\.?)", c):
            col["no"] = j
        elif "item" not in col and ("item" in c or "description" in c):
            col["item"] = j
        elif "model" not in col and any(k in c for k in ("model", "spec", "part", "brand", "offered")):
            col["model"] = j
        elif "unit" not in col and any(k in c for k in ("uom", "unit")):
            col["unit"] = j
        elif "qty" not in col and any(k in c for k in ("qty", "quant")):
            col["qty"] = j
    if "model" not in col:  # blank-header column between Item and Unit
        lo, hi = col["item"], col.get("unit", col["item"] + 2)
        for j in range(lo + 1, hi):
            if not cells[j]:
                col["model"] = j
                break

    def get(r, k):
        j = col.get(k)
        return r[j] if j is not None and j < len(r) else None

    lines = []
    for i, r in enumerate(rows[hdr + 1:], start=hdr + 2):
        no = _cell(get(r, "no"))
        if not re.fullmatch(r"\d+(\.\d+)*", no):
            continue
        model = _cell(get(r, "model"))
        qty = get(r, "qty")
        if isinstance(qty, float) and qty.is_integer():
            qty = int(qty)
        lines.append(BomLine(no=no, item=_cell(get(r, "item")), model=model,
                             unit=_cell(get(r, "unit")), qty=qty if qty is not None else "",
                             level=no.count("."), row=i))
    nos = [l.no for l in lines]
    for l in lines:
        l.is_group = any(n.startswith(l.no + ".") for n in nos)
        if "." in l.no:
            l.parent = l.no.rsplit(".", 1)[0]
    return lines


def to_dicts(lines):
    return [asdict(l) for l in lines]


def from_dicts(ds):
    return [BomLine(**d) for d in ds]
