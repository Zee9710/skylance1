#!/usr/bin/env python3
"""Editable Word companion to the Skylance Technical Submittal, matching the
company house style (header band with logo + CR, footer contact bar, black
section bars with cyan accent, cyan sub-headings, branded tables)."""
import os
from docx import Document
from docx.shared import Pt, RGBColor, Cm, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.enum.section import WD_SECTION
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

ASSET = os.path.join(os.path.dirname(os.path.abspath(__file__)), "brand_assets")
HEADER_IMG = os.path.join(ASSET, "header_cobrand.png")  # Skylance + SELA co-branded band
FOOTER_IMG = os.path.join(ASSET, "crop_footer.png")

CYAN = RGBColor(0x26, 0xB8, 0xEF)
BLACK = RGBColor(0x11, 0x11, 0x11)
INK = RGBColor(0x1F, 0x1F, 0x1F)
GREY = RGBColor(0x3A, 0x3A, 0x3A)
SOFT = RGBColor(0x88, 0x88, 0x88)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
FILL_BLACK = "111111"
FILL_ROW = "F2F2F2"
HAIRHEX = "CFCFCF"

DOC_NO = "TS-SL-260629"
REV = "Rev A – For Design Review"
FONT = "Calibri"

doc = Document()
A4_W = Cm(21.0)
PAGE_W = Cm(21.0)

normal = doc.styles["Normal"]
normal.font.name = FONT
normal.font.size = Pt(10.5)
normal.font.color.rgb = INK
normal.paragraph_format.line_spacing = 1.12
normal.paragraph_format.space_after = Pt(6)


def shade(cell, hexc):
    tcPr = cell._tc.get_or_add_tcPr()
    sh = OxmlElement("w:shd")
    sh.set(qn("w:val"), "clear"); sh.set(qn("w:color"), "auto"); sh.set(qn("w:fill"), hexc)
    tcPr.append(sh)


def cell_text(cell, text, bold=False, color=INK, size=10, align=WD_ALIGN_PARAGRAPH.LEFT):
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    p = cell.paragraphs[0]; p.alignment = align
    p.paragraph_format.space_after = Pt(2); p.paragraph_format.space_before = Pt(2)
    r = p.add_run(text); r.font.bold = bold; r.font.color.rgb = color
    r.font.size = Pt(size); r.font.name = FONT


def cell_margins(cell, t=70, b=70, l=120, r=120):
    tcPr = cell._tc.get_or_add_tcPr(); m = OxmlElement("w:tcMar")
    for tag, v in (("top", t), ("bottom", b), ("start", l), ("end", r)):
        e = OxmlElement(f"w:{tag}"); e.set(qn("w:w"), str(v)); e.set(qn("w:type"), "dxa"); m.append(e)
    tcPr.append(m)


def no_borders(table):
    tblPr = table._tbl.tblPr; b = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        e = OxmlElement(f"w:{edge}"); e.set(qn("w:val"), "none"); b.append(e)
    tblPr.append(b)


def grid(table, color=HAIRHEX, sz=6):
    tblPr = table._tbl.tblPr; b = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        e = OxmlElement(f"w:{edge}"); e.set(qn("w:val"), "single"); e.set(qn("w:sz"), str(sz))
        e.set(qn("w:space"), "0"); e.set(qn("w:color"), color); b.append(e)
    tblPr.append(b)


def pborder(p, color, size=18, where="bottom", space=4):
    pPr = p._p.get_or_add_pPr(); pbdr = OxmlElement("w:pBdr")
    e = OxmlElement(f"w:{where}"); e.set(qn("w:val"), "single"); e.set(qn("w:sz"), str(size))
    e.set(qn("w:space"), str(space)); e.set(qn("w:color"), color); pbdr.append(e); pPr.append(pbdr)


# ---------------- header / footer with brand bands ----------------
def add_field(p, field):
    r = p.add_run()
    f1 = OxmlElement("w:fldChar"); f1.set(qn("w:fldCharType"), "begin")
    it = OxmlElement("w:instrText"); it.set(qn("xml:space"), "preserve"); it.text = field
    f2 = OxmlElement("w:fldChar"); f2.set(qn("w:fldCharType"), "end")
    r._r.append(f1); r._r.append(it); r._r.append(f2)
    r.font.size = Pt(8); r.font.color.rgb = SOFT
    return r


def build_header(section, full_w):
    h = section.header
    h.is_linked_to_previous = False
    p = h.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.space_before = Pt(0)
    # co-branded band already includes CONTRACTOR/CLIENT labels + cyan rule
    p.add_run().add_picture(HEADER_IMG, width=full_w)


def build_footer(section, full_w):
    f = section.footer
    f.is_linked_to_previous = False
    pg = f.paragraphs[0]
    pg.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pg.paragraph_format.space_after = Pt(1)
    r = pg.add_run("Page "); r.font.size = Pt(7.5); r.font.color.rgb = SOFT
    add_field(pg, "PAGE")
    r2 = pg.add_run(f"  |  {DOC_NO}  |  {REV}  |  Confidential")
    r2.font.size = Pt(7.5); r2.font.color.rgb = SOFT
    img = f.add_paragraph()
    img.alignment = WD_ALIGN_PARAGRAPH.CENTER
    img.paragraph_format.space_after = Pt(0); img.paragraph_format.space_before = Pt(0)
    img.add_run().add_picture(FOOTER_IMG, width=full_w)


# ---------------- content helpers ----------------
def h1(num, title):
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    no_borders(tbl)
    c = tbl.rows[0].cells[0]
    shade(c, FILL_BLACK); cell_margins(c, t=90, b=90, l=140)
    # cyan left accent border on the cell
    tcPr = c._tc.get_or_add_tcPr(); b = OxmlElement("w:tcBorders")
    le = OxmlElement("w:start"); le.set(qn("w:val"), "single"); le.set(qn("w:sz"), "22")
    le.set(qn("w:space"), "0"); le.set(qn("w:color"), "26B8EF"); b.append(le); tcPr.append(b)
    p = c.paragraphs[0]
    r = p.add_run(f"{num}.  {title.upper()}")
    r.font.bold = True; r.font.size = Pt(11.5); r.font.color.rgb = WHITE; r.font.name = FONT
    sp = doc.add_paragraph(); sp.paragraph_format.space_after = Pt(2)
    return tbl


def h2(text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10); p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(text); r.font.bold = True; r.font.size = Pt(11); r.font.color.rgb = CYAN
    return p


def applic(text):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(6)
    r = p.add_run("Applicable to:  "); r.font.italic = True; r.font.bold = True
    r.font.size = Pt(9.5); r.font.color.rgb = CYAN
    r2 = p.add_run(text); r2.font.italic = True; r2.font.size = Pt(9.5); r2.font.color.rgb = GREY
    return p


def micro(text):
    p = doc.add_paragraph(); p.paragraph_format.space_before = Pt(4); p.paragraph_format.space_after = Pt(2)
    r = p.add_run(text); r.font.bold = True; r.font.size = Pt(10); r.font.color.rgb = INK
    return p


def body(text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    r = p.add_run(text); r.font.size = Pt(10.5); r.font.color.rgb = INK
    return p


def bullet(text):
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.left_indent = Cm(0.8); p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.line_spacing = 1.1
    r = p.add_run(text); r.font.size = Pt(10.5); r.font.color.rgb = INK


def qty_table(headers, rows, total):
    tbl = doc.add_table(rows=1, cols=2); tbl.alignment = WD_TABLE_ALIGNMENT.LEFT; grid(tbl)
    hd = tbl.rows[0].cells
    for i, h in enumerate(headers):
        shade(hd[i], FILL_BLACK); cell_margins(hd[i])
        cell_text(hd[i], h, bold=True, color=WHITE, size=10,
                  align=WD_ALIGN_PARAGRAPH.LEFT if i == 0 else WD_ALIGN_PARAGRAPH.CENTER)
    for ri, row in enumerate(rows):
        cs = tbl.add_row().cells
        for i, v in enumerate(row):
            if ri % 2 == 1: shade(cs[i], FILL_ROW)
            cell_margins(cs[i])
            cell_text(cs[i], v, bold=(i == 0), color=INK, size=10,
                      align=WD_ALIGN_PARAGRAPH.LEFT if i == 0 else WD_ALIGN_PARAGRAPH.CENTER)
    cs = tbl.add_row().cells
    for i, v in enumerate(total):
        shade(cs[i], FILL_BLACK); cell_margins(cs[i])
        cell_text(cs[i], v, bold=True, color=CYAN, size=10,
                  align=WD_ALIGN_PARAGRAPH.LEFT if i == 0 else WD_ALIGN_PARAGRAPH.CENTER)
    for r in tbl.rows:
        r.cells[0].width = Cm(8.5); r.cells[1].width = Cm(8.5)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)


# ============== SECTION SETUP ==============
SKYLANCE_IMG = os.path.join(ASSET, "skylance_logo.png")
SELA_IMG = os.path.join(ASSET, "sela_logo.png")


def cell_right_border(cell, color=HAIRHEX, sz=6):
    tcPr = cell._tc.get_or_add_tcPr(); b = OxmlElement("w:tcBorders")
    e = OxmlElement("w:end"); e.set(qn("w:val"), "single"); e.set(qn("w:sz"), str(sz))
    e.set(qn("w:space"), "0"); e.set(qn("w:color"), color); b.append(e); tcPr.append(b)


def cover_lockup():
    """Standalone Skylance + SELA logos side by side for the cover white space."""
    t = doc.add_table(rows=2, cols=3); t.alignment = WD_TABLE_ALIGNMENT.CENTER
    no_borders(t)
    w = [Cm(6.0), Cm(1.0), Cm(6.0)]
    for row in t.rows:
        for i, cw in enumerate(w):
            row.cells[i].width = cw
    # logos row
    c0, cd, c2 = t.rows[0].cells
    for c in (c0, cd, c2):
        c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    p0 = c0.paragraphs[0]; p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p0.add_run().add_picture(SKYLANCE_IMG, height=Cm(1.75))
    cell_right_border(c0)
    p2 = c2.paragraphs[0]; p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p2.add_run().add_picture(SELA_IMG, height=Cm(0.78))
    # labels row
    l0, ld, l2 = t.rows[1].cells
    for cell, lbl, val in ((l0, "CONTRACTOR", "Skylance Technical Services Co."),
                           (l2, "CLIENT", "Jeddah Yacht Club & Marina")):
        pp = cell.paragraphs[0]; pp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pp.paragraph_format.space_before = Pt(4)
        r = pp.add_run(lbl); r.font.bold = True; r.font.size = Pt(8); r.font.color.rgb = CYAN
        r2 = pp.add_run("\n" + val); r2.font.size = Pt(7.5); r2.font.color.rgb = GREY
    doc.add_paragraph().paragraph_format.space_after = Pt(2)


sec0 = doc.sections[0]
sec0.top_margin = Cm(1.5); sec0.bottom_margin = Cm(2.2)
sec0.left_margin = Cm(2.0); sec0.right_margin = Cm(2.0)
sec0.header_distance = Cm(0.4); sec0.footer_distance = Cm(0.4)
FULLW = Cm(21.0)  # full-bleed band width = A4 width
# cover uses a standalone logo lockup (below), not the co-branded band header
build_footer(sec0, FULLW)
sec0.different_first_page_header_footer = False

# ============== COVER ==============
cover_lockup()
hrc = doc.add_paragraph(); hrc.paragraph_format.space_after = Pt(0)
pborder(hrc, "26B8EF", size=14, where="bottom", space=4)
for _ in range(5):
    sp = doc.add_paragraph(); sp.paragraph_format.space_after = Pt(10)
hr = doc.add_paragraph(); hr.alignment = WD_ALIGN_PARAGRAPH.CENTER
pborder(hr, HAIRHEX, size=6, where="bottom", space=6)
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(16); p.paragraph_format.space_after = Pt(8)
r = p.add_run("TECHNICAL SUBMITTAL"); r.font.bold = True; r.font.size = Pt(28); r.font.color.rgb = BLACK
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.space_after = Pt(4)
r = p.add_run("Marina Infrastructure Maintenance and Rehabilitation Works")
r.font.bold = True; r.font.size = Pt(13); r.font.color.rgb = CYAN
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.space_after = Pt(14)
r = p.add_run("Jeddah Yacht Club & Marina (JYC)"); r.font.size = Pt(11.5); r.font.color.rgb = GREY
hr2 = doc.add_paragraph(); hr2.alignment = WD_ALIGN_PARAGRAPH.CENTER
hr2.paragraph_format.space_after = Pt(20)
pborder(hr2, "26B8EF", size=14, where="bottom", space=6)

ctl = [
    ("Project", "Jeddah Yacht Club & Marina (JYC)"),
    ("Client / Consultant", "SELA"),
    ("Document No.", DOC_NO),
    ("Revision", REV),
    ("Prepared By", "SKYLANCE TECHNICAL SERVICES"),
    ("Date", "29 Jun 2026"),
    ("Classification", "Technical Submittal – For Review & Approval"),
]
t = doc.add_table(rows=len(ctl), cols=2); t.alignment = WD_TABLE_ALIGNMENT.CENTER; grid(t)
for i, (k, v) in enumerate(ctl):
    c0, c1 = t.rows[i].cells
    shade(c0, FILL_BLACK); cell_margins(c0); cell_margins(c1)
    cell_text(c0, k, bold=True, color=WHITE, size=9.5)
    if i % 2 == 1: shade(c1, FILL_ROW)
    cell_text(c1, v, bold=False, color=INK, size=9.5)
for r in t.rows:
    r.cells[0].width = Cm(5.0); r.cells[1].width = Cm(12.0)
doc.add_paragraph().paragraph_format.space_after = Pt(10)
for line in [
    "This document is prepared by Skylance Technical Services for submission to the Client / Consultant for review and approval.",
    "All specified methods and materials are subject to final approval. Actual performance may vary depending on site conditions and maintenance practices.",
]:
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.space_after = Pt(3)
    r = p.add_run(line); r.font.italic = True; r.font.size = Pt(8.5); r.font.color.rgb = SOFT

doc.add_section(WD_SECTION.NEW_PAGE)
sec1 = doc.sections[1]
sec1.top_margin = Cm(2.9); sec1.bottom_margin = Cm(2.2)
sec1.left_margin = Cm(2.0); sec1.right_margin = Cm(2.0)
sec1.header_distance = Cm(0.4); sec1.footer_distance = Cm(0.4)
build_header(sec1, FULLW); build_footer(sec1, FULLW)

# ============== CONTENT ==============
h1("1", "Introduction")
body("This document has been developed to describe the technical details of the Jeddah Yacht Club "
     "& Marina (JYC) Marina Infrastructure Maintenance and Rehabilitation Works, covering the "
     "inspection, repair, refurbishment, replacement and protective maintenance of marina floating "
     "assets, gangways, wave breaker structures, quay wall surfaces, navigation poles and associated "
     "marine assets.")
body("Jeddah Yacht Club & Marina (JYC) has invited qualified marine infrastructure contractors to "
     "undertake these works. Skylance Technical Services presents this Technical Submittal in response, "
     "setting out the structure-wise scope of work, methods and material standards to be applied across "
     "the marina in accordance with applicable Saudi Arabian regulations and recognised marine industry "
     "standards.")

h1("2", "Objective")
body("The objective of this scope of work is to restore structural integrity, improve operational "
     "safety, prevent further corrosion and extend the service life of marina assets through "
     "comprehensive maintenance works.")
h2("In delivering the works, the Contractor shall:")
for t in [
    "Conduct a detailed site survey and verification of quantities prior to commencement.",
    "Supply all labour, supervision, tools, equipment, materials, consumables and marine access equipment required.",
    "Dispose of waste materials in accordance with local environmental regulations.",
    "Maintain uninterrupted marina operations where reasonably practicable.",
    "Comply with all applicable Saudi Arabian safety regulations and marine industry standards.",
    "Submit method statements, risk assessments, work schedules and quality control procedures before mobilization.",
    "Offer warranty on all works completed.",
]: bullet(t)

h1("3", "Scope of Works")
body("The scope of work includes the following structures and marine assets, with detailed scope "
     "prescribed structure-wise in the sections that follow:")
h2("Scope Includes")
for t in [
    "Floating Docks A, B, C, D and K, and the Slipway Pontoon.",
    "Main Wave Breaker structure.",
    "Navigation Poles.",
    "Quay Wall infrastructure.",
    "Gangways serving all docks and the slipway pontoon.",
]: bullet(t)
body("The detailed, structure-wise scope of work for each of the above assets is set out in "
     "Sections 4 through 11 of this submittal.")

h1("4", "Gangway Rehabilitation Works")
applic("Dock A, B, C, D, K & Slipway Pontoon")
h2("4.1  Timber Works")
for t in ["Clean, prepare, sand and apply marine-grade protective coating to all gangway timber planks.",
          "Replace any damaged timber sections identified during execution."]: bullet(t)
h2("4.2  Roller System")
for t in ["Remove and replace damaged gangway rollers.",
          "Supply additional PVC rollers as operational spare stock."]: bullet(t)
h2("4.3  Teflon Wear Protection")
for t in ["Remove damaged gangway flip protection pads.",
          "Supply and install new marine-grade Teflon wear pads."]: bullet(t)
h2("4.4  Structural Steel Rehabilitation")
body("Where applicable, the Contractor shall:")
for t in ["Clean steel surfaces by mechanical preparation or approved blasting method.",
          "Remove corrosion and contaminants.",
          "Apply approved marine-grade anti-corrosion coating system.",
          "Repaint gangway H-beam support structures."]: bullet(t)
h2("4.5  Anchor Bolt Replacement")
body("Where applicable, the Contractor shall:")
for t in ["Remove corroded foundation bolts and nuts.",
          "Supply and install new stainless steel or approved galvanized anchor bolts.",
          "Upgrade bolt sizing where specified to improve structural stability."]: bullet(t)
h2("4.6  Electrical Safety Improvements")
body("The Contractor shall:")
for t in ["Install weatherproof protection for all exposed dock gate electrical components.",
          "Provide suitable enclosures, sealing systems, covers and fixing arrangements.",
          "Inspect and rectify electrical grounding (earthing) systems for all gangways.",
          "Test and certify grounding resistance values following completion."]: bullet(t)
h2("4.7  Cable Tray Works (Docks D and K)")
body("The Contractor shall:")
for t in ["Secure under-gangway cable trays using approved marine-grade support systems.",
          "Supply and install all required fixing accessories.",
          "Ensure cable tray alignment and support integrity."]: bullet(t)
body("Based on requirements identified across all dock gangways.")

h1("5", "Pontoon Joint Plate Repairs")
applic("Dock A, B, C, D, K")
h2("5.1  Preferred Method (Option A)")
for t in ["Remove damaged joint plates.",
          "Drill two additional fixing holes on site.",
          "Install new fixings into pontoon concrete structure using approved plastic grips and marine-grade bolts.",
          "Seal all abandoned holes with approved marine repair compound."]: bullet(t)
h2("5.2  Alternative Method (Option B)")
bullet("Weld four support bolts to joint plates where approved by JYC.")
micro("Joint Plate Quantities")
qty_table(["Dock", "Quantity"],
          [["Dock A", "16"], ["Dock B", "30"], ["Dock C", "18"], ["Dock D", "24"], ["Dock K", "16"]],
          ["Total", "104 Joint Plates"])

h1("6", "Pontoon Surface and Pile Refurbishment")
applic("Dock A, B, C, K & Slipway Pontoon")
body("The Contractor shall:")
for t in ["Prepare all pontoon surfaces.", "Remove loose coatings and corrosion.",
          "Apply approved marine coating system.", "Repaint all pontoon surfaces.",
          "Prepare and repaint marina piles.",
          "Provide coating system data sheets and warranty information."]: bullet(t)

h1("7", "Fiberglass Pile Cap Replacement")
body("The Contractor shall supply and install new fiberglass pile caps as follows:")
qty_table(["Location", "Quantity"],
          [["Dock A", "9"], ["Dock B", "13"], ["Dock C", "9"], ["Dock D", "9"], ["Dock K", "5"]],
          ["Total", "45 Pile Caps"])
body("The Contractor shall:")
for t in ["Remove existing damaged pile caps.", "Supply marine-grade fiberglass pile caps.",
          "Install and secure pile caps in accordance with manufacturer recommendations."]: bullet(t)

h1("8", "Pile Bracket and Wear Pad Repairs")
applic("Dock D & K")
body("The Contractor shall:")
for t in ["Remove damaged wear pads.", "Supply and install replacement wear pads.",
          "Adjust pile bracket fixing hole alignment.",
          "Rectify bracket positioning to prevent premature wear.",
          "Supply spare wear pads for future maintenance."]: bullet(t)
micro("Minimum Quantity")
bullet("10 wear pads per dock, inclusive of spare stock.")

h1("9", "Main Wave Breaker and Navigation Pole Rehabilitation")
body("The Contractor shall carry out the following works:")
h2("9.1  Wave Breaker")
for t in ["Prepare exposed steel surfaces.", "Remove corrosion.",
          "Apply approved marine anti-corrosion coating system.",
          "Repaint complete wave breaker structure."]: bullet(t)
h2("9.2  Navigation Poles")
for t in ["Remove corrosion.", "Repaint navigation poles.",
          "Replace heavily corroded foundation anchor bolts."]: bullet(t)
h2("9.3  Manhole Repair")
for t in ["Remove temporary plywood cover.",
          "Supply and install permanent marine-grade manhole cover.",
          "Reinstate surrounding area."]: bullet(t)

h1("10", "Quay Wall Surface Repainting")
body("The Contractor shall:")
for t in ["Prepare top surface of quay wall surrounding the marina basin.",
          "Remove deteriorated coatings.", "Apply approved marine coating system.",
          "Repaint quay wall top surfaces only."]: bullet(t)

h1("11", "Slipway Pontoon Structural Repairs")
body("The Contractor shall:")
for t in ["Inspect and repair stainless steel H-beam attached to quay wall.",
          "Reinstall and secure foundation anchor bolts.",
          "Repair damaged / spalled concrete around foundation areas.",
          "Reinstate concrete using approved marine repair materials.",
          "Ensure structural integrity and long-term durability of the installation."]: bullet(t)

ep = doc.add_paragraph(); ep.alignment = WD_ALIGN_PARAGRAPH.CENTER
ep.paragraph_format.space_before = Pt(12)
r = ep.add_run("— End of Technical Submittal —")
r.font.italic = True; r.font.size = Pt(9); r.font.color.rgb = SOFT

out = "Technical Submittal - JYC Marina Infrastructure (Skylance).docx"
doc.save(out)
print("saved:", out)
