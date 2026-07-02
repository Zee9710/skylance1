#!/usr/bin/env python3
"""Rebuild third-party product sheets onto the Skylance letterhead.

  1. 11,000 lb Floating Lift — Technical Specifications  (Basta branding removed)
  2. Rotomoulded Products — Technical Data               (Marina Performance branding removed)

Content is recreated in the Skylance house style (letterhead header band,
footer contact bar, light watermark); original vendor branding is dropped.
"""
import os
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_JUSTIFY
from reportlab.platypus import (
    BaseDocTemplate, PageTemplate, Frame, Paragraph, Spacer, Table, TableStyle,
    ListFlowable, ListItem, HRFlowable, Image,
)
from reportlab.lib.styles import ParagraphStyle as PS
from reportlab.lib.utils import ImageReader

HERE = os.path.dirname(os.path.abspath(__file__))
ASSET = os.path.join(HERE, "brand_assets")
HEADER_IMG = os.path.join(ASSET, "crop_header.png")
FOOTER_IMG = os.path.join(ASSET, "crop_footer.png")
WM_IMG = os.path.join(ASSET, "crop_watermark.png")
COLLAGE = os.path.join(HERE, "product_assets", "rotomoulded_collage.png")

CYAN = colors.HexColor("#26B8EF")
BLACK = colors.HexColor("#111111")
INK = colors.HexColor("#1F1F1F")
GREY = colors.HexColor("#3A3A3A")
SOFTGREY = colors.HexColor("#888888")
ROWGREY = colors.HexColor("#F2F2F2")
HAIR = colors.HexColor("#CFCFCF")

PAGE_W, PAGE_H = A4
LM = RM = 1.9 * cm
HDR_W = PAGE_W
HDR_H = PAGE_W * (322.0 / 2481.0)
FTR_W = PAGE_W
FTR_H = PAGE_W * (141.0 / 2481.0)
FRAME_TOP = PAGE_H - HDR_H - 20
FRAME_BOTTOM = FTR_H + 20
TM = PAGE_H - FRAME_TOP
BM = FRAME_BOTTOM

FOOTER_LINE = "Skylance Technical Services  |  Product Data  |  Confidential"


def S(name, **kw):
    return PS(name, **kw)


body = S("body", fontName="Helvetica", fontSize=10, leading=14.5, textColor=INK,
         alignment=TA_JUSTIFY, spaceAfter=7)
lead = S("lead", parent=body, spaceAfter=4, keepWithNext=1)


def h1(title):
    cell = Paragraph(f'<font color="white"><b>{title.upper()}</b></font>',
                     S("h1", fontName="Helvetica-Bold", fontSize=11.5, leading=15, textColor=colors.white))
    t = Table([[cell]], colWidths=[PAGE_W - LM - RM])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), BLACK),
        ("LINEBEFORE", (0, 0), (0, -1), 3.2, CYAN),
        ("TOPPADDING", (0, 0), (-1, -1), 6), ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("LEFTPADDING", (0, 0), (-1, -1), 10), ("RIGHTPADDING", (0, 0), (-1, -1), 8),
    ]))
    t.spaceBefore = 4
    t.spaceAfter = 8
    t.keepWithNext = 1
    return t


def h2(text):
    p = Paragraph(f'<b>{text}</b>', S("h2", fontName="Helvetica-Bold", fontSize=11, leading=14,
                                      textColor=CYAN, spaceBefore=8, spaceAfter=4, keepWithNext=1))
    p.keepWithNext = 1
    return p


def spec_table(rows):
    w = PAGE_W - LM - RM
    data = [[Paragraph('<b><font color="white">Parameter</font></b>',
                       S("th", fontName="Helvetica-Bold", fontSize=9.5, leading=12)),
             Paragraph('<b><font color="white">Specification</font></b>',
                       S("th2", fontName="Helvetica-Bold", fontSize=9.5, leading=12))]]
    for k, v in rows:
        data.append([Paragraph(k, S("k", fontName="Helvetica-Bold", fontSize=10, leading=13, textColor=INK)),
                     Paragraph(v, S("v", fontName="Helvetica", fontSize=10, leading=13, textColor=INK))])
    t = Table(data, colWidths=[w * 0.55, w * 0.45], hAlign="LEFT")
    style = [
        ("BACKGROUND", (0, 0), (-1, 0), BLACK),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("GRID", (0, 0), (-1, -1), 0.6, HAIR),
        ("TOPPADDING", (0, 0), (-1, -1), 6), ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("LEFTPADDING", (0, 0), (-1, -1), 9),
    ]
    for i in range(1, len(rows) + 1):
        if i % 2 == 0:
            style.append(("BACKGROUND", (0, i), (-1, i), ROWGREY))
    t.setStyle(TableStyle(style))
    return t


def draw_brand(canvas, doc):
    canvas.saveState()
    try:
        wm = ImageReader(WM_IMG); iw, ih = wm.getSize()
        tw = PAGE_W * 0.62; th = tw * ih / iw
        canvas.drawImage(wm, (PAGE_W - tw) / 2, (PAGE_H - th) / 2 - 10, width=tw, height=th, mask="auto")
    except Exception:
        pass
    canvas.drawImage(HEADER_IMG, 0, PAGE_H - HDR_H, width=HDR_W, height=HDR_H, mask="auto")
    ry = PAGE_H - HDR_H - 12
    canvas.setStrokeColor(CYAN); canvas.setLineWidth(1.4)
    canvas.line(LM, ry, PAGE_W - RM, ry)
    canvas.drawImage(FOOTER_IMG, 0, 0, width=FTR_W, height=FTR_H, mask="auto")
    canvas.setFont("Helvetica", 7.5); canvas.setFillColor(SOFTGREY)
    canvas.drawCentredString(PAGE_W / 2, FTR_H + 6, f"Page {canvas.getPageNumber()}  |  {FOOTER_LINE}")
    canvas.restoreState()


def build(out, story):
    doc = BaseDocTemplate(out, pagesize=A4, leftMargin=LM, rightMargin=RM,
                          topMargin=TM, bottomMargin=BM, author="Skylance Technical Services")
    frame = Frame(LM, FRAME_BOTTOM, PAGE_W - LM - RM, FRAME_TOP - FRAME_BOTTOM, id="f")
    doc.addPageTemplates([PageTemplate(id="P", frames=[frame], onPage=draw_brand)])
    doc.build(story)
    print("built:", out)


# =====================================================================
#  DOC 1 — 11,000 lb Floating Lift, Technical Specifications
# =====================================================================
E = []
E.append(h1("11,000 lb Floating Lift — Technical Specifications"))
E.append(Paragraph("Standard specifications for the 11,000 lb capacity floating boat lift.", lead))
E.append(spec_table([
    ("Lifting Capacity", "11,000 lbs"),
    ("Minimum Slip Width", "12′"),
    ("Launch Speed", "40 sec"),
    ("Solar Charging", "12 V"),
    ("Remote Control Fobs", "2"),
    ("Stainless Steel Cylinders", "4"),
    ("Standard Hose Length", "27″  |  823 cm"),
    ("Submerged Hose Ends", "Stainless Steel"),
    ("Connecting Hardware", "Stainless Steel"),
    ("Standard Max Boat Beam", "108″  |  275 cm"),
    ("Custom Max Boat Beam", "132″"),
    ("Minimum Water Depth", "27″"),
]))
E.append(Spacer(1, 6))
E.append(Paragraph("* Contact Skylance Technical Services for additional specifications.",
                   S("note", fontName="Helvetica-Oblique", fontSize=8.5, leading=12, textColor=SOFTGREY)))
build(os.path.join(HERE, "11,000 lb Floating Lift - Technical Specifications (Skylance).pdf"), E)


# =====================================================================
#  DOC 2 — Rotomoulded Products, Technical Data
# =====================================================================
E = []
E.append(h1("Rotomoulded Products — Boat Lifts, Jet Ski Docks & Modular Docks"))
# photo collage (vendor branding already excluded from the crop)
cw = PAGE_W - LM - RM
col_w = cw
col_h = col_w * 1480.0 / 1820.0
E.append(Image(COLLAGE, width=col_w, height=col_h))
E.append(Spacer(1, 8))

pt = S("pt", fontName="Helvetica", fontSize=9.5, leading=13, textColor=INK, alignment=TA_JUSTIFY)


def numbered(items):
    li = []
    for i, txt in enumerate(items, 1):
        li.append(ListItem(Paragraph(txt, pt), value=i, spaceAfter=6))
    return ListFlowable(li, bulletType="1", bulletFormat="%s.", bulletFontName="Helvetica-Bold",
                        bulletFontSize=10, bulletColor=CYAN, leftIndent=20, spaceBefore=2)


pts = [
    "Boat Lifts, Jet ski docks and Superior Modular Docks (SMDs) are all designed for non-deterioration "
    "in the marine environment and can all easily connect to any pontoon or marina structure.",
    "The D-Series Boat Lift utilises drive-on pontoon technology; the vessel can be driven straight on to "
    "the lift to provide berthing out of the water.",
    "Stylish jet ski docks ensure that all craft are stored completely out of the water. Multiple docks can "
    "be connected side by side to form a continuous dock for commercial use.",
    "Polyethylene SMD modules are tough, UV resistant and versatile for arranging in any configuration.",
]
E.append(numbered(pts))
build(os.path.join(HERE, "Rotomoulded Products - Technical Data (Skylance).pdf"), E)
