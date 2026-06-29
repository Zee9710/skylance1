#!/usr/bin/env python3
"""Skylance Technical Submittal - JYC Marina Infrastructure.
Rendered in the Skylance house style using brand assets cropped from the
company template (header band with logo + CR, footer contact bar, watermark)."""
import os
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_JUSTIFY
from reportlab.platypus import (
    BaseDocTemplate, PageTemplate, Frame, Paragraph, Spacer, Table, TableStyle,
    ListFlowable, ListItem, NextPageTemplate, PageBreak, HRFlowable, Image,
)
from reportlab.lib.styles import ParagraphStyle as PS
from reportlab.lib.utils import ImageReader

ASSET = os.path.join(os.path.dirname(os.path.abspath(__file__)), "brand_assets")
HEADER_IMG = os.path.join(ASSET, "header_cobrand.png")  # Skylance + SELA co-branded band
FOOTER_IMG = os.path.join(ASSET, "crop_footer.png")
WM_IMG = os.path.join(ASSET, "crop_watermark.png")

# ---------------- palette (sampled from template) ----------------
CYAN = colors.HexColor("#26B8EF")
BLACK = colors.HexColor("#111111")
INK = colors.HexColor("#1F1F1F")
GREY = colors.HexColor("#3A3A3A")
SOFTGREY = colors.HexColor("#888888")
ROWGREY = colors.HexColor("#F2F2F2")
HAIR = colors.HexColor("#CFCFCF")

PAGE_W, PAGE_H = A4
LM = RM = 1.9 * cm

# co-branded header band geometry (image aspect 2481 x 442)
HDR_W = PAGE_W
HDR_H = PAGE_W * (442.0 / 2481.0)
# footer band geometry (image 2481 x 141)
FTR_W = PAGE_W
FTR_H = PAGE_W * (141.0 / 2481.0)

# content frame
FRAME_TOP = PAGE_H - HDR_H - 6           # band already includes labels + cyan rule
FRAME_BOTTOM = FTR_H + 20                 # footer band + page-no line
TM = PAGE_H - FRAME_TOP
BM = FRAME_BOTTOM

DOC_NO = "TS-SL-260629"
REV = "Rev A – For Design Review"
FOOTER_LINE = f"TS-SL-260629  |  {REV}  |  Confidential"


def S(name, **kw):
    return PS(name, **kw)


body = S("body", fontName="Helvetica", fontSize=10, leading=14.5, textColor=INK,
         alignment=TA_JUSTIFY, spaceAfter=7)
lead = S("lead", parent=body, spaceAfter=4)
applic = S("applic", fontName="Helvetica-Oblique", fontSize=9.5, leading=13,
           textColor=GREY, spaceBefore=1, spaceAfter=7)
bullet = S("bullet", fontName="Helvetica", fontSize=10, leading=13.5, textColor=INK,
           spaceAfter=0)


def h1(num, title):
    """Black full-width bar, white bold text, cyan left accent."""
    cell = Paragraph(
        f'<font color="white"><b>{num}.&nbsp;&nbsp;{title.upper()}</b></font>',
        S("h1", fontName="Helvetica-Bold", fontSize=11.5, leading=15, textColor=colors.white))
    t = Table([[cell]], colWidths=[PAGE_W - LM - RM])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), BLACK),
        ("LINEBEFORE", (0, 0), (0, -1), 3.2, CYAN),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("LEFTPADDING", (0, 0), (-1, -1), 10),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
    ]))
    t.spaceBefore = 14
    t.spaceAfter = 8
    return t


def h2(text):
    return Paragraph(f'<b>{text}</b>',
                     S("h2", fontName="Helvetica-Bold", fontSize=11, leading=14,
                       textColor=CYAN, spaceBefore=10, spaceAfter=4))


def applic_p(text):
    return Paragraph(
        f'<font color="#26B8EF"><b><i>Applicable to:&nbsp;&nbsp;</i></b></font>'
        f'<i><font color="#3A3A3A">{text}</font></i>', applic)


def micro(text):
    return Paragraph(f'<b>{text}</b>',
                     S("micro", fontName="Helvetica-Bold", fontSize=10, leading=13,
                       textColor=INK, spaceBefore=4, spaceAfter=3))


def blist(items):
    li = [ListItem(Paragraph(t, bullet), leftIndent=14, spaceAfter=5,
                   bulletColor=BLACK) for t in items]
    return ListFlowable(li, bulletType="bullet", start="•", bulletFontSize=8,
                        bulletColor=BLACK, leftIndent=18, bulletOffsetY=1,
                        spaceBefore=2, spaceAfter=7)


def qty_table(headers, rows, total):
    data = [headers] + rows + [total]
    w = PAGE_W - LM - RM
    t = Table(data, colWidths=[w * 0.5, w * 0.5], hAlign="LEFT")
    style = [
        ("BACKGROUND", (0, 0), (-1, 0), BLACK),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 9.5),
        ("ALIGN", (1, 0), (1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("FONTNAME", (0, 1), (0, -1), "Helvetica-Bold"),
        ("TEXTCOLOR", (0, 1), (-1, -2), INK),
        ("GRID", (0, 0), (-1, -1), 0.6, HAIR),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("LEFTPADDING", (0, 0), (-1, -1), 9),
        ("BACKGROUND", (0, -1), (-1, -1), BLACK),
        ("TEXTCOLOR", (0, -1), (-1, -1), CYAN),
        ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"),
    ]
    for i in range(1, len(rows) + 1):
        if i % 2 == 0:
            style.append(("BACKGROUND", (0, i), (-1, i), ROWGREY))
    t.setStyle(TableStyle(style))
    t.spaceAfter = 6
    return t


def coat_table(headers, rows):
    data = [headers] + rows
    w = PAGE_W - LM - RM
    cw = [w * 0.22, w * 0.28, w * 0.50]
    cells = [[Paragraph(f'<b><font color="white">{c}</font></b>',
                        S("th", fontName="Helvetica-Bold", fontSize=9.5, leading=12))
              for c in headers]]
    for r in rows:
        cells.append([Paragraph(x, S("td", fontName="Helvetica", fontSize=9.5,
                                     leading=12, textColor=INK)) for x in r])
    t = Table(cells, colWidths=cw, hAlign="LEFT")
    style = [
        ("BACKGROUND", (0, 0), (-1, 0), BLACK),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("GRID", (0, 0), (-1, -1), 0.6, HAIR),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
    ]
    for i in range(1, len(rows) + 1):
        if i % 2 == 0:
            style.append(("BACKGROUND", (0, i), (-1, i), ROWGREY))
    t.setStyle(TableStyle(style))
    t.spaceAfter = 6
    return t


# ---------------- page furniture ----------------
def draw_brand(canvas, doc, cover=False):
    canvas.saveState()
    # watermark (centered in content area, faint)
    try:
        wm = ImageReader(WM_IMG)
        iw, ih = wm.getSize()
        tw = PAGE_W * 0.62
        th = tw * ih / iw
        canvas.drawImage(wm, (PAGE_W - tw) / 2, (PAGE_H - th) / 2 - 10,
                         width=tw, height=th, mask="auto")
    except Exception:
        pass
    # co-branded header band: Skylance logo + CR (left), SELA logo (right),
    # CONTRACTOR / CLIENT labels and cyan rule (all baked into the image).
    # The cover page uses a standalone logo lockup instead of the band.
    if not cover:
        canvas.drawImage(HEADER_IMG, 0, PAGE_H - HDR_H, width=HDR_W, height=HDR_H, mask="auto")
    # footer band
    canvas.drawImage(FOOTER_IMG, 0, 0, width=FTR_W, height=FTR_H, mask="auto")
    # page-number line above footer band
    canvas.setFont("Helvetica", 7.5)
    canvas.setFillColor(SOFTGREY)
    pno = canvas.getPageNumber()
    canvas.drawCentredString(PAGE_W / 2, FTR_H + 6,
                             f"Page {pno}  |  {FOOTER_LINE}")
    canvas.restoreState()


def on_cover(c, d):
    draw_brand(c, d, cover=True)


def on_body(c, d):
    draw_brand(c, d, cover=False)


doc = BaseDocTemplate(
    "Technical Submittal - JYC Marina Infrastructure (Skylance).pdf",
    pagesize=A4, leftMargin=LM, rightMargin=RM, topMargin=TM, bottomMargin=BM,
    title="Technical Submittal - JYC Marina Infrastructure",
    author="Skylance Technical Services",
)
frame = Frame(LM, FRAME_BOTTOM, PAGE_W - LM - RM, FRAME_TOP - FRAME_BOTTOM, id="f")
COVER_TOP = PAGE_H - 1.5 * cm
cover_frame = Frame(LM, FRAME_BOTTOM, PAGE_W - LM - RM, COVER_TOP - FRAME_BOTTOM, id="cf")
doc.addPageTemplates([
    PageTemplate(id="Cover", frames=[cover_frame], onPage=on_cover),
    PageTemplate(id="Body", frames=[frame], onPage=on_body),
])

SKYLANCE_IMG = os.path.join(ASSET, "skylance_logo.png")
SELA_IMG = os.path.join(ASSET, "sela_logo.png")


def cover_logo_lockup():
    """Standalone Skylance + SELA logos side by side for the cover white space."""
    sky_h = 62.0
    sky = Image(SKYLANCE_IMG, width=sky_h * 346.0 / 294.0, height=sky_h)
    sela_h = 27.0
    sela = Image(SELA_IMG, width=sela_h * 479.0 / 158.0, height=sela_h)
    lbl = S("lbl", fontName="Helvetica-Bold", fontSize=8, leading=11,
            textColor=CYAN, alignment=TA_CENTER)
    sub = S("lblsub", fontName="Helvetica", fontSize=7.5, leading=10,
            textColor=GREY, alignment=TA_CENTER)
    cont = Paragraph('CONTRACTOR<br/><font color="#3A3A3A">Skylance Technical Services Co.</font>', lbl)
    clnt = Paragraph('CLIENT<br/><font color="#3A3A3A">Jeddah Yacht Club &amp; Marina</font>', lbl)
    t = Table([[sky, "", sela], [cont, "", clnt]],
              colWidths=[210, 36, 210], hAlign="CENTER")
    t.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, 0), "MIDDLE"),
        ("VALIGN", (0, 1), (-1, 1), "TOP"),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("LINEAFTER", (0, 0), (0, 0), 0.8, HAIR),
        ("TOPPADDING", (0, 1), (-1, 1), 12),
        ("BOTTOMPADDING", (0, 0), (-1, 0), 4),
    ]))
    return t

E = []

# =================== COVER ===================
E.append(Spacer(1, 18))
E.append(cover_logo_lockup())
E.append(HRFlowable(width="100%", thickness=1.4, color=CYAN, spaceBefore=14, spaceAfter=0))
E.append(Spacer(1, 120))
E.append(HRFlowable(width="86%", thickness=0.8, color=HAIR, hAlign="CENTER", spaceAfter=18))
E.append(Paragraph("TECHNICAL SUBMITTAL",
                   S("ttl", fontName="Helvetica-Bold", fontSize=29, leading=34,
                     textColor=BLACK, alignment=TA_CENTER, spaceAfter=10)))
E.append(Paragraph("Marina Infrastructure Maintenance and Rehabilitation Works",
                   S("sub", fontName="Helvetica-Bold", fontSize=13, leading=17,
                     textColor=CYAN, alignment=TA_CENTER, spaceAfter=5)))
E.append(Paragraph("Jeddah Yacht Club &amp; Marina (JYC)",
                   S("sub2", fontName="Helvetica", fontSize=11.5, leading=15,
                     textColor=GREY, alignment=TA_CENTER, spaceAfter=18)))
E.append(HRFlowable(width="86%", thickness=1.4, color=CYAN, hAlign="CENTER", spaceAfter=26))

ctl = [
    ("Project", "Jeddah Yacht Club & Marina (JYC)"),
    ("Client / Consultant", "SELA"),
    ("Document No.", DOC_NO),
    ("Revision", REV),
    ("Prepared By", "SKYLANCE TECHNICAL SERVICES"),
    ("Date", "29 Jun 2026"),
    ("Classification", "Technical Submittal – For Review & Approval"),
]
tbl = Table(
    [[Paragraph(f'<b><font color="white">{k}</font></b>',
                S("k", fontName="Helvetica-Bold", fontSize=9.5, leading=12)),
      Paragraph(v, S("v", fontName="Helvetica", fontSize=9.5, leading=12, textColor=INK))]
     for k, v in ctl],
    colWidths=[(PAGE_W - LM - RM) * 0.30, (PAGE_W - LM - RM) * 0.70], hAlign="CENTER")
cs = [
    ("BACKGROUND", (0, 0), (0, -1), BLACK),
    ("GRID", (0, 0), (-1, -1), 0.6, HAIR),
    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ("TOPPADDING", (0, 0), (-1, -1), 7),
    ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
    ("LEFTPADDING", (0, 0), (-1, -1), 9),
]
for i in range(len(ctl)):
    if i % 2 == 1:
        cs.append(("BACKGROUND", (1, i), (1, i), ROWGREY))
tbl.setStyle(TableStyle(cs))
E.append(tbl)
E.append(Spacer(1, 22))
E.append(Paragraph(
    "This document is prepared by Skylance Technical Services for submission to the "
    "Client / Consultant for review and approval.",
    S("disc", fontName="Helvetica-Oblique", fontSize=8.5, leading=12, textColor=SOFTGREY,
      alignment=TA_CENTER, spaceAfter=3)))
E.append(Paragraph(
    "All specified methods and materials are subject to final approval. Actual performance "
    "may vary depending on site conditions and maintenance practices.",
    S("disc2", fontName="Helvetica-Oblique", fontSize=8.5, leading=12, textColor=SOFTGREY,
      alignment=TA_CENTER)))

E.append(NextPageTemplate("Body"))
E.append(PageBreak())

# =================== TABLE OF CONTENTS ===================
toc_rows = [
    ("1", "Introduction", "3"),
    ("2", "Objective", "3"),
    ("3", "Scope of Works", "3"),
    ("4", "Gangway Rehabilitation Works", "4"),
    ("5", "Pontoon Joint Plate Repairs", "5"),
    ("6", "Pontoon Surface and Pile Refurbishment", "5"),
    ("7", "Fiberglass Pile Cap Replacement", "6"),
    ("8", "Pile Bracket and Wear Pad Repairs", "6"),
    ("9", "Main Wave Breaker and Navigation Pole Rehabilitation", "6"),
    ("10", "Quay Wall Surface Repainting", "7"),
    ("11", "Slipway Pontoon Structural Repairs", "7"),
]
hdr = Table([[Paragraph('<b><font color="white">TABLE OF CONTENTS</font></b>',
                        S("toch", fontName="Helvetica-Bold", fontSize=11, leading=14))]],
            colWidths=[PAGE_W - LM - RM])
hdr.setStyle(TableStyle([
    ("BACKGROUND", (0, 0), (-1, -1), BLACK),
    ("LINEBEFORE", (0, 0), (0, -1), 3.2, CYAN),
    ("TOPPADDING", (0, 0), (-1, -1), 6), ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
    ("LEFTPADDING", (0, 0), (-1, -1), 10),
]))
E.append(hdr)
E.append(Spacer(1, 4))
w = PAGE_W - LM - RM
rows = []
for n, t, pg in toc_rows:
    rows.append([
        Paragraph(f'<b>{n}</b>', S("tn", fontName="Helvetica-Bold", fontSize=9.5, textColor=INK)),
        Paragraph(t, S("tt", fontName="Helvetica", fontSize=9.5, textColor=GREY)),
        Paragraph(pg, S("tp", fontName="Helvetica", fontSize=9.5, textColor=GREY)),
    ])
toc = Table(rows, colWidths=[w * 0.07, w * 0.83, w * 0.10])
ts = [
    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ("LINEBELOW", (0, 0), (-1, -1), 0.5, HAIR),
    ("ALIGN", (2, 0), (2, -1), "RIGHT"),
    ("LINEBEFORE", (1, 0), (1, -1), 2.2, CYAN),
    ("TOPPADDING", (0, 0), (-1, -1), 7), ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
    ("LEFTPADDING", (1, 0), (1, -1), 10),
]
for i in range(len(rows)):
    if i % 2 == 1:
        ts.append(("BACKGROUND", (0, i), (-1, i), ROWGREY))
toc.setStyle(TableStyle(ts))
E.append(toc)
E.append(PageBreak())

# =================== 1. INTRODUCTION ===================
E.append(h1("1", "Introduction"))
E.append(Paragraph(
    "This document has been developed to describe the technical details of the Jeddah Yacht Club "
    "&amp; Marina (JYC) Marina Infrastructure Maintenance and Rehabilitation Works, covering the "
    "inspection, repair, refurbishment, replacement and protective maintenance of marina floating "
    "assets, gangways, wave breaker structures, quay wall surfaces, navigation poles and associated "
    "marine assets.", body))
E.append(Paragraph(
    "Jeddah Yacht Club &amp; Marina (JYC) has invited qualified marine infrastructure contractors to "
    "undertake these works. Skylance Technical Services presents this Technical Submittal in response, "
    "setting out the structure-wise scope of work, methods and material standards to be applied across "
    "the marina in accordance with applicable Saudi Arabian regulations and recognised marine industry "
    "standards.", body))

# =================== 2. OBJECTIVE ===================
E.append(h1("2", "Objective"))
E.append(Paragraph(
    "The objective of this scope of work is to restore structural integrity, improve operational "
    "safety, prevent further corrosion and extend the service life of marina assets through "
    "comprehensive maintenance works.", body))
E.append(h2("In delivering the works, the Contractor shall:"))
E.append(blist([
    "Conduct a detailed site survey and verification of quantities prior to commencement.",
    "Supply all labour, supervision, tools, equipment, materials, consumables and marine access equipment required.",
    "Dispose of waste materials in accordance with local environmental regulations.",
    "Maintain uninterrupted marina operations where reasonably practicable.",
    "Comply with all applicable Saudi Arabian safety regulations and marine industry standards.",
    "Submit method statements, risk assessments, work schedules and quality control procedures before mobilization.",
    "Offer warranty on all works completed.",
]))

# =================== 3. SCOPE OF WORKS ===================
E.append(h1("3", "Scope of Works"))
E.append(Paragraph(
    "The scope of work includes the following structures and marine assets, with detailed scope "
    "prescribed structure-wise in the sections that follow:", lead))
E.append(h2("Scope Includes"))
E.append(blist([
    "Floating Docks A, B, C, D and K, and the Slipway Pontoon.",
    "Main Wave Breaker structure.",
    "Navigation Poles.",
    "Quay Wall infrastructure.",
    "Gangways serving all docks and the slipway pontoon.",
]))
E.append(Paragraph(
    "The detailed, structure-wise scope of work for each of the above assets is set out in "
    "Sections 4 through 11 of this submittal.", body))

# =================== 4. GANGWAY REHABILITATION ===================
E.append(h1("4", "Gangway Rehabilitation Works"))
E.append(applic_p("Dock A, B, C, D, K &amp; Slipway Pontoon"))
E.append(h2("4.1  Timber Works"))
E.append(blist([
    "Clean, prepare, sand and apply marine-grade protective coating to all gangway timber planks.",
    "Replace any damaged timber sections identified during execution.",
]))
E.append(h2("4.2  Roller System"))
E.append(blist([
    "Remove and replace damaged gangway rollers.",
    "Supply additional PVC rollers as operational spare stock.",
]))
E.append(h2("4.3  Teflon Wear Protection"))
E.append(blist([
    "Remove damaged gangway flip protection pads.",
    "Supply and install new marine-grade Teflon wear pads.",
]))
E.append(h2("4.4  Structural Steel Rehabilitation"))
E.append(Paragraph("Where applicable, the Contractor shall:", lead))
E.append(blist([
    "Clean steel surfaces by mechanical preparation or approved blasting method.",
    "Remove corrosion and contaminants.",
    "Apply approved marine-grade anti-corrosion coating system.",
    "Repaint gangway H-beam support structures.",
]))
E.append(h2("4.5  Anchor Bolt Replacement"))
E.append(Paragraph("Where applicable, the Contractor shall:", lead))
E.append(blist([
    "Remove corroded foundation bolts and nuts.",
    "Supply and install new stainless steel or approved galvanized anchor bolts.",
    "Upgrade bolt sizing where specified to improve structural stability.",
]))
E.append(h2("4.6  Electrical Safety Improvements"))
E.append(Paragraph("The Contractor shall:", lead))
E.append(blist([
    "Install weatherproof protection for all exposed dock gate electrical components.",
    "Provide suitable enclosures, sealing systems, covers and fixing arrangements.",
    "Inspect and rectify electrical grounding (earthing) systems for all gangways.",
    "Test and certify grounding resistance values following completion.",
]))
E.append(h2("4.7  Cable Tray Works (Docks D and K)"))
E.append(Paragraph("The Contractor shall:", lead))
E.append(blist([
    "Secure under-gangway cable trays using approved marine-grade support systems.",
    "Supply and install all required fixing accessories.",
    "Ensure cable tray alignment and support integrity.",
]))
E.append(Paragraph("Based on requirements identified across all dock gangways.", body))

# =================== 5. PONTOON JOINT PLATE REPAIRS ===================
E.append(h1("5", "Pontoon Joint Plate Repairs"))
E.append(applic_p("Dock A, B, C, D, K"))
E.append(h2("5.1  Preferred Method (Option A)"))
E.append(blist([
    "Remove damaged joint plates.",
    "Drill two additional fixing holes on site.",
    "Install new fixings into pontoon concrete structure using approved plastic grips and marine-grade bolts.",
    "Seal all abandoned holes with approved marine repair compound.",
]))
E.append(h2("5.2  Alternative Method (Option B)"))
E.append(blist(["Weld four support bolts to joint plates where approved by JYC."]))
E.append(micro("Joint Plate Quantities"))
E.append(qty_table(
    ["Dock", "Quantity"],
    [["Dock A", "16"], ["Dock B", "30"], ["Dock C", "18"], ["Dock D", "24"], ["Dock K", "16"]],
    ["Total", "104 Joint Plates"]))

# =================== 6. PONTOON SURFACE & PILE REFURB ===================
E.append(h1("6", "Pontoon Surface and Pile Refurbishment"))
E.append(applic_p("Dock A, B, C, K &amp; Slipway Pontoon"))
E.append(Paragraph("The Contractor shall:", lead))
E.append(blist([
    "Prepare all pontoon surfaces.",
    "Remove loose coatings and corrosion.",
    "Apply approved marine coating system.",
    "Repaint all pontoon surfaces.",
    "Prepare and repaint marina piles.",
    "Provide coating system data sheets and warranty information.",
]))

# =================== 7. FIBERGLASS PILE CAP ===================
E.append(h1("7", "Fiberglass Pile Cap Replacement"))
E.append(Paragraph("The Contractor shall supply and install new fiberglass pile caps as follows:", lead))
E.append(qty_table(
    ["Location", "Quantity"],
    [["Dock A", "9"], ["Dock B", "13"], ["Dock C", "9"], ["Dock D", "9"], ["Dock K", "5"]],
    ["Total", "45 Pile Caps"]))
E.append(Paragraph("The Contractor shall:", lead))
E.append(blist([
    "Remove existing damaged pile caps.",
    "Supply marine-grade fiberglass pile caps.",
    "Install and secure pile caps in accordance with manufacturer recommendations.",
]))

# =================== 8. PILE BRACKET & WEAR PAD ===================
E.append(h1("8", "Pile Bracket and Wear Pad Repairs"))
E.append(applic_p("Dock D &amp; K"))
E.append(Paragraph("The Contractor shall:", lead))
E.append(blist([
    "Remove damaged wear pads.",
    "Supply and install replacement wear pads.",
    "Adjust pile bracket fixing hole alignment.",
    "Rectify bracket positioning to prevent premature wear.",
    "Supply spare wear pads for future maintenance.",
]))
E.append(micro("Minimum Quantity"))
E.append(blist(["10 wear pads per dock, inclusive of spare stock."]))

# =================== 9. WAVE BREAKER & NAV POLE ===================
E.append(h1("9", "Main Wave Breaker and Navigation Pole Rehabilitation"))
E.append(Paragraph("The Contractor shall carry out the following works:", lead))
E.append(h2("9.1  Wave Breaker"))
E.append(blist([
    "Prepare exposed steel surfaces.",
    "Remove corrosion.",
    "Apply approved marine anti-corrosion coating system.",
    "Repaint complete wave breaker structure.",
]))
E.append(h2("9.2  Navigation Poles"))
E.append(blist([
    "Remove corrosion.",
    "Repaint navigation poles.",
    "Replace heavily corroded foundation anchor bolts.",
]))
E.append(h2("9.3  Manhole Repair"))
E.append(blist([
    "Remove temporary plywood cover.",
    "Supply and install permanent marine-grade manhole cover.",
    "Reinstate surrounding area.",
]))

# =================== 10. QUAY WALL ===================
E.append(h1("10", "Quay Wall Surface Repainting"))
E.append(Paragraph("The Contractor shall:", lead))
E.append(blist([
    "Prepare top surface of quay wall surrounding the marina basin.",
    "Remove deteriorated coatings.",
    "Apply approved marine coating system.",
    "Repaint quay wall top surfaces only.",
]))

# =================== 11. SLIPWAY PONTOON STRUCTURAL ===================
E.append(h1("11", "Slipway Pontoon Structural Repairs"))
E.append(Paragraph("The Contractor shall:", lead))
E.append(blist([
    "Inspect and repair stainless steel H-beam attached to quay wall.",
    "Reinstall and secure foundation anchor bolts.",
    "Repair damaged / spalled concrete around foundation areas.",
    "Reinstate concrete using approved marine repair materials.",
    "Ensure structural integrity and long-term durability of the installation.",
]))
E.append(Spacer(1, 14))
E.append(HRFlowable(width="38%", thickness=1.0, color=CYAN, hAlign="CENTER", spaceAfter=6))
E.append(Paragraph("— End of Technical Submittal —",
                   S("end", fontName="Helvetica-Oblique", fontSize=9, textColor=SOFTGREY,
                     alignment=TA_CENTER)))

doc.build(E)
print("PDF built.")
