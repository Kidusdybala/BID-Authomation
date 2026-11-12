"""Builds stand-in inputs (a V04-style bid, BOMs, datasheet PDFs, a tender) into a folder.
Product names are fictional test data."""
import io
import sys
from pathlib import Path
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm
from PIL import Image, ImageDraw
import openpyxl
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas


def png(label, color):
    im = Image.new("RGB", (900, 600), color)
    ImageDraw.Draw(im).text((40, 40), label, fill="white")
    b = io.BytesIO()
    im.save(b, "PNG")
    b.seek(0)
    return b


def listpara(doc, text):
    p = doc.add_paragraph(style="List Paragraph")
    ppr = p._p.get_or_add_pPr()
    num = OxmlElement("w:numPr")
    for tag, v in (("w:ilvl", "0"), ("w:numId", "1")):
        e = OxmlElement(tag)
        e.set(qn("w:val"), v)
        num.append(e)
    ppr.append(num)
    p.add_run(text)


def make_v04(path):
    d = Document()
    d.add_paragraph("{{PROJECT_NAME}}", style="Title")
    d.add_paragraph("Prepared for {{CLIENT}}  |  {{TENDER_NO}}  |  {{DATE}}")
    toc = d.add_paragraph()
    for typ, txt in (("begin", None), (None, ' TOC \\o "1-3" \\h \\z \\u '), ("separate", None), ("t", "old toc"), ("end", None)):
        r = OxmlElement("w:r")
        if typ in ("begin", "separate", "end"):
            f = OxmlElement("w:fldChar"); f.set(qn("w:fldCharType"), typ); r.append(f)
        elif typ == "t":
            t = OxmlElement("w:t"); t.text = txt; r.append(t)
        else:
            i = OxmlElement("w:instrText"); i.text = txt; r.append(i)
        toc._p.append(r)
    d.add_heading("Company Background", 1)
    d.add_paragraph("Acme Systems Ltd supplies and installs security and IT systems across East Africa.")
    d.add_heading("Project Introduction", 1)
    d.add_heading("Project Overview", 2)
    d.add_paragraph("The OLD CLIENT project covers a campus security upgrade at Old Town with 12 sites.")
    d.add_heading("Purpose of the Project", 2)
    d.add_paragraph("The purpose is to improve safety. We will deliver a reliable system.")
    d.add_heading("Technical Proposal", 1)
    d.add_heading("Camera System", 2)
    d.add_heading("Acme NC-4K Network Camera NC4K200", 3)
    d.add_paragraph("The NC4K200 is a 4K network camera for indoor and outdoor surveillance.")
    d.add_paragraph("Key Features")
    for t in ("Sensor: 8 MP CMOS", "Lens: 2.8 mm fixed", "Power: PoE 802.3af, max 12 W", "Protection: IP67 and IK10"):
        listpara(d, t)
    d.add_picture(png("NC4K200", (30, 60, 120)), width=Cm(9))
    d.add_paragraph("Figure 1 Acme NC-4K camera", style="Caption")
    d.add_heading("Network Switch", 2)
    d.add_heading("Acme SW-24P PoE Switch SW24P100", 3)
    d.add_paragraph("The SW24P100 is a 24 port PoE switch for camera networks.")
    d.add_paragraph("Key Features")
    for t in ("Ports: 24 x 10/100/1000 Mbps PoE+", "Uplink: 4 x SFP", "PoE budget: 370 W"):
        listpara(d, t)
    d.add_heading("Data and Network Connection", 3)
    d.add_paragraph("All devices connect over the client LAN. This text is bid specific.")
    d.add_heading("Bill of Materials", 1)
    t = d.add_table(rows=1, cols=5)
    t.style = "Table Grid"
    for c, h in zip(t.rows[0].cells, ("No", "Item", "Model", "Unit", "Quantity")):
        c.text = h
    for row in (("1", "Camera System", "", "", ""), ("1.1", "Network Camera", "Acme NC4K200", "Nos", "10"),
                ("2", "Network Switch", "Acme SW24P100", "Nos", "2")):
        cs = t.add_row().cells
        for c, v in zip(cs, row):
            c.text = v
    d.save(path)


def make_bom_v04(path):
    wb = openpyxl.Workbook(); ws = wb.active
    ws.append(["Project BOM"]); ws.append([])
    ws.append(["No", "Item", "Model", "Unit", "Quantity"])
    for r in (("1", "Camera System", None, None, None), ("1.1", "Network Camera 4K", "Acme NC4K200", "Nos", 24),
              ("1.2", "Thermal Camera", "Zeta TH900X", "Nos", 2),
              ("2", "Network Switch 24-port PoE", "Acme SW24P100", "Nos", 3),
              ("3", "Rack Cabinet", "Generic 42U Rack RK42U", "Nos", 1)):
        ws.append(list(r))
    wb.save(path)


def make_bom_mod(path):
    wb = openpyxl.Workbook(); ws = wb.active
    ws.append(["No.", "Required Item", None, "UoM", "Qty."])
    rows = [("1", "Desktop PC Dell OptiPlex 3090", "Dell OptiPlex 3096", "Nos", 10),
            ("2", "Printer HP LaserJet Pro M401dn", "HP LaserJet Pro 4003dn", "Nos", 4),
            ("3", "Full-frame camera", "Canon EOS R10", "Nos", 2),
            ("4", "Antivirus license", "AV Suite Pro 2026", "License", 4000),
            ("5", "Antivirus license", "AV Suite Pro 2026", "License", 4000),
            ("6", "Cable", "Cat6 305m", "Roll", 0)]
    for r in rows:
        ws.append(list(r))
    wb.save(path)


def make_pdf(path, pages):
    c = canvas.Canvas(str(path), pagesize=A4)
    for lines in pages:
        y = 800
        for ln in lines:
            c.drawString(50, y, ln); y -= 18
        c.showPage()
    c.save()


def make_all(root):
    root = Path(root)
    (root / "datasheets").mkdir(parents=True, exist_ok=True)
    make_v04(root / "V04.docx")
    make_bom_v04(root / "bom_v04.xlsx")
    make_bom_mod(root / "bom_mod.xlsx")
    make_pdf(root / "datasheets" / "zeta_th900x.pdf", [
        ["Zeta TH900X Thermal Camera", "Zeta TH900X is a thermal imaging camera for perimeter detection.",
         "Resolution: 640 x 480 thermal", "Detection range: 1500 m vehicle", "Power: PoE 802.3at, max 20 W",
         "Protection: IP66", "Operating temperature: -40 to 70 C"],
        ["Zeta TH900X ordering information", "Waterproof lens included."]])
    make_pdf(root / "datasheets" / "unrelated.pdf", [["Some other product", "Nothing relevant here 1234"]])
    (root / "tender.txt").write_text(
        "Project: Riverside Campus Security Upgrade\nClient: Riverside University\nTender No: RU/2026/045\n"
        "Location: Riverside, Kenya\n\nThe university wants to improve campus safety.\n\n"
        "REQ 1.1: Camera resolution at least 12 MP\nREQ 1.1: Power over Ethernet 802.3af\n"
        "REQ 1.2: Detection range at least 1000 m\nREQ 1.2: Protection rating IP66\nREQ 1.2: Built-in speaker\n"
        "REQ 2: PoE budget at least 300 W\nREQ 2: At least 4 SFP uplink ports\n", "utf-8")


if __name__ == "__main__":
    make_all(sys.argv[1])
