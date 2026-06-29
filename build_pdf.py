#!/usr/bin/env python3
"""Generate the Skylance / SELA Technical Submittal as a polished PDF."""
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm, mm
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_JUSTIFY
from reportlab.platypus import (
    BaseDocTemplate, PageTemplate, Frame, Paragraph, Spacer, Table, TableStyle,
    ListFlowable, ListItem, KeepTogether, NextPageTemplate, PageBreak, HRFlowable,
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

# ---------------- palette ----------------
NAVY = colors.HexColor("#0B1F3A")
STEEL = colors.HexColor("#1C4E80")
ACCENT = colors.HexColor("#0E7690")
GREY = colors.HexColor("#3C434C")
LIGHT = colors.HexColor("#7A8694")
HAIR = colors.HexColor("#D5DEE7")
ROW = colors.HexColor("#EAF1F7")
SOFT = colors.HexColor("#F2F6FA")

PAGE_W, PAGE_H = A4
LM = RM = 2.0 * cm
TM = 2.7 * cm
BM = 1.9 * cm

styles = getSampleStyleSheet()


def S(name, **kw):
    return ParagraphStyle(name, **kw)


body = S("body", fontName="Helvetica", fontSize=10, leading=14.5, textColor=GREY,
         alignment=TA_JUSTIFY, spaceAfter=7)
intro = S("intro", parent=body, alignment=TA_LEFT)
applic = S("applic", fontName="Helvetica-Oblique", fontSize=9.5, leading=13,
           textColor=GREY, spaceBefore=1, spaceAfter=7)
micro = S("micro", fontName="Helvetica-Bold", fontSize=10, leading=13, textColor=NAVY,
          spaceBefore=4, spaceAfter=3)
lead = S("lead", parent=body, spaceAfter=4)
bullet = S("bullet", fontName="Helvetica", fontSize=10, leading=13.5, textColor=GREY,
           spaceAfter=0)


def h1(num, title):
    t = Table(
        [[Paragraph(f'<font color="#0E7690">{num}.</font>&nbsp;&nbsp;'
                    f'<font color="#0B1F3A">{title.upper()}</font>',
                    S("h1", fontName="Helvetica-Bold", fontSize=13, leading=16))]],
        colWidths=[PAGE_W - LM - RM],
    )
    t.setStyle(TableStyle([
        ("LINEBELOW", (0, 0), (-1, -1), 1.1, STEEL),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
    ]))
    return t


def h2(num, title):
    return Paragraph(
        f'<font color="#1C4E80"><b>{num}&nbsp;&nbsp;{title}</b></font>',
        S("h2", fontName="Helvetica-Bold", fontSize=11, leading=14,
          spaceBefore=10, spaceAfter=3, textColor=STEEL))


def applic_p(text):
    return Paragraph(
        f'<font color="#0E7690"><b><i>Applicable to:&nbsp;&nbsp;</i></b></font>'
        f'<i>{text}</i>', applic)


def blist(items):
    li = [ListItem(Paragraph(t, bullet), leftIndent=14, spaceAfter=4,
                   value=None, bulletColor=ACCENT) for t in items]
    return ListFlowable(
        li, bulletType="bullet", start="•", bulletFontSize=8,
        bulletColor=ACCENT, leftIndent=16, bulletOffsetY=1, spaceBefore=1, spaceAfter=6,
    )


def qty_table(headers, rows, total):
    data = [headers] + rows + [total]
    w = (PAGE_W - LM - RM)
    t = Table(data, colWidths=[w * 0.55, w * 0.45], hAlign="LEFT")
    style = [
        ("BACKGROUND", (0, 0), (-1, 0), NAVY),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 9.5),
        ("ALIGN", (1, 0), (1, -1), "CENTER"),
        ("ALIGN", (0, 0), (0, -1), "LEFT"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("FONTNAME", (0, 1), (0, -1), "Helvetica-Bold"),
        ("FONTNAME", (1, 1), (1, -2), "Helvetica"),
        ("TEXTCOLOR", (0, 1), (-1, -2), GREY),
        ("GRID", (0, 0), (-1, -1), 0.6, HAIR),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        # total row
        ("BACKGROUND", (0, -1), (-1, -1), STEEL),
        ("TEXTCOLOR", (0, -1), (-1, -1), colors.white),
        ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"),
    ]
    for i in range(1, len(rows) + 1):
        if i % 2 == 0:
            style.append(("BACKGROUND", (0, i), (-1, i), ROW))
    t.setStyle(TableStyle(style))
    return t


# ---------------- page furniture ----------------
def header_footer(canvas, doc, cover=False):
    canvas.saveState()
    if not cover:
        y = PAGE_H - 1.35 * cm
        canvas.setFont("Helvetica-Bold", 14)
        canvas.setFillColor(NAVY)
        canvas.drawString(LM, y, "SELA")
        canvas.setFont("Helvetica-Bold", 14)
        canvas.setFillColor(STEEL)
        canvas.drawRightString(PAGE_W - RM, y, "SKYLANCE")
        canvas.setFont("Helvetica", 6)
        canvas.setFillColor(LIGHT)
        canvas.drawString(LM, y - 9, "Client / Consultant")
        canvas.drawRightString(PAGE_W - RM, y - 9, "Marine Infrastructure Contractor")
        canvas.setStrokeColor(STEEL)
        canvas.setLineWidth(0.8)
        canvas.line(LM, y - 14, PAGE_W - RM, y - 14)
    # footer
    fy = 1.25 * cm
    canvas.setStrokeColor(HAIR)
    canvas.setLineWidth(0.5)
    canvas.line(LM, fy + 6, PAGE_W - RM, fy + 6)
    canvas.setFont("Helvetica", 7)
    canvas.setFillColor(LIGHT)
    canvas.drawString(LM, fy - 2,
                      "Technical Submittal  |  Marina Infrastructure Maintenance & Rehabilitation Works")
    canvas.drawRightString(PAGE_W - RM, fy - 2, f"Page {canvas.getPageNumber()}")
    canvas.restoreState()


def on_cover(canvas, doc):
    header_footer(canvas, doc, cover=True)


def on_body(canvas, doc):
    header_footer(canvas, doc, cover=False)


# ---------------- document ----------------
doc = BaseDocTemplate(
    "Technical Submittal - JYC Marina Infrastructure (Skylance).pdf",
    pagesize=A4, leftMargin=LM, rightMargin=RM, topMargin=TM, bottomMargin=BM,
    title="Technical Submittal - JYC Marina Infrastructure",
    author="Skylance",
)
frame_cover = Frame(LM, BM, PAGE_W - LM - RM, PAGE_H - 1.6 * cm - BM, id="cover")
frame_body = Frame(LM, BM, PAGE_W - LM - RM, PAGE_H - TM - BM, id="body")
doc.addPageTemplates([
    PageTemplate(id="Cover", frames=[frame_cover], onPage=on_cover),
    PageTemplate(id="Body", frames=[frame_body], onPage=on_body),
])

E = []

# ---------- COVER ----------
brand = Table([[
    Paragraph('<font color="#0B1F3A" size="22"><b>SELA</b></font><br/>'
              '<font color="#7A8694" size="8">Client / Consultant</font>',
              S("bl", fontName="Helvetica", leading=24, alignment=TA_LEFT)),
    Paragraph('<font color="#1C4E80" size="22"><b>SKYLANCE</b></font><br/>'
              '<font color="#7A8694" size="8">Marine Infrastructure Contractor</font>',
              S("br", fontName="Helvetica", leading=24, alignment=2)),
]], colWidths=[(PAGE_W - LM - RM) / 2] * 2)
brand.setStyle(TableStyle([
    ("VALIGN", (0, 0), (-1, -1), "TOP"),
    ("LEFTPADDING", (0, 0), (-1, -1), 0),
    ("RIGHTPADDING", (0, 0), (-1, -1), 0),
]))
E.append(brand)
E.append(Spacer(1, 6))
E.append(HRFlowable(width="100%", thickness=1.6, color=ACCENT, spaceAfter=10))
E.append(Spacer(1, 90))

E.append(Paragraph('<b>JEDDAH YACHT CLUB &amp; MARINA&nbsp;&nbsp;(JYC)</b>',
                   S("c1", fontName="Helvetica-Bold", fontSize=12, leading=16,
                     textColor=ACCENT, alignment=TA_CENTER, spaceAfter=4)))
E.append(Paragraph('Marina Infrastructure Maintenance and Rehabilitation Works',
                   S("c2", fontName="Helvetica", fontSize=13, leading=17,
                     textColor=GREY, alignment=TA_CENTER, spaceAfter=26)))
E.append(Paragraph('TECHNICAL SUBMITTAL',
                   S("ttl", fontName="Helvetica-Bold", fontSize=33, leading=38,
                     textColor=NAVY, alignment=TA_CENTER, spaceAfter=8)))
E.append(Paragraph('Inspection &middot; Repair &middot; Refurbishment &middot; '
                   'Replacement &middot; Protective Maintenance',
                   S("c3", fontName="Helvetica-Oblique", fontSize=10, leading=14,
                     textColor=LIGHT, alignment=TA_CENTER)))
E.append(Spacer(1, 80))

ctl_rows = [
    ("Project", "Marina Infrastructure Maintenance & Rehabilitation Works"),
    ("Client", "Jeddah Yacht Club & Marina (JYC) / SELA"),
    ("Contractor", "Skylance"),
    ("Document", "Technical Submittal"),
    ("Document Ref.", "SKL-JYC-TS-001  ·  Rev. 0"),
    ("Date", "June 2026"),
]
ctl = Table(
    [[Paragraph(f'<b>{k}</b>', S("k", fontName="Helvetica-Bold", fontSize=9.5,
                                 textColor=colors.white, leading=12)),
      Paragraph(v, S("v", fontName="Helvetica", fontSize=9.5, textColor=GREY, leading=12))]
     for k, v in ctl_rows],
    colWidths=[(PAGE_W - LM - RM) * 0.28, (PAGE_W - LM - RM) * 0.72], hAlign="CENTER",
)
cstyle = [
    ("BACKGROUND", (0, 0), (0, -1), NAVY),
    ("GRID", (0, 0), (-1, -1), 0.6, HAIR),
    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ("TOPPADDING", (0, 0), (-1, -1), 6),
    ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
    ("LEFTPADDING", (0, 0), (-1, -1), 9),
]
for i in range(len(ctl_rows)):
    if i % 2 == 1:
        cstyle.append(("BACKGROUND", (1, i), (1, i), SOFT))
ctl.setStyle(TableStyle(cstyle))
E.append(ctl)

E.append(NextPageTemplate("Body"))
E.append(PageBreak())

# ---------- 1. INTRODUCTION ----------
E.append(h1("1", "Introduction"))
E.append(Paragraph(
    "This document has been developed to describe the technical details of the Jeddah Yacht Club "
    "&amp; Marina (JYC) Marina Infrastructure Maintenance and Rehabilitation Works, covering the "
    "inspection, repair, refurbishment, replacement and protective maintenance of marina floating "
    "assets, gangways, wave breaker structures, quay wall surfaces, navigation poles and associated "
    "marine assets.", body))
E.append(Paragraph(
    "Jeddah Yacht Club &amp; Marina (JYC) has invited qualified marine infrastructure contractors to "
    "undertake these works. Skylance presents this Technical Submittal in response, setting out the "
    "structure-wise scope of work, methods and material standards to be applied across the marina in "
    "accordance with applicable Saudi Arabian regulations and recognised marine industry standards.",
    body))

# ---------- 2. OBJECTIVE ----------
E.append(h1("2", "Objective"))
E.append(Paragraph(
    "The objective of this scope of work is to restore structural integrity, improve operational "
    "safety, prevent further corrosion and extend the service life of marina assets through "
    "comprehensive maintenance works.", body))
E.append(Paragraph("In delivering the works, the Contractor shall:", micro))
E.append(blist([
    "Conduct a detailed site survey and verification of quantities prior to commencement.",
    "Supply all labour, supervision, tools, equipment, materials, consumables and marine access equipment required.",
    "Dispose of waste materials in accordance with local environmental regulations.",
    "Maintain uninterrupted marina operations where reasonably practicable.",
    "Comply with all applicable Saudi Arabian safety regulations and marine industry standards.",
    "Submit method statements, risk assessments, work schedules and quality control procedures before mobilization.",
    "Offer warranty on all works completed.",
]))

# ---------- 3. SCOPE OF WORKS ----------
E.append(h1("3", "Scope of Works"))
E.append(Paragraph(
    "The scope of work includes the following structures and marine assets, with detailed scope "
    "prescribed structure-wise in the sections that follow:", lead))
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

# ---------- 4. GANGWAY REHABILITATION ----------
E.append(h1("4", "Gangway Rehabilitation Works"))
E.append(applic_p("Dock A, B, C, D, K &amp; Slipway Pontoon"))
E.append(h2("4.1", "Timber Works"))
E.append(blist([
    "Clean, prepare, sand and apply marine-grade protective coating to all gangway timber planks.",
    "Replace any damaged timber sections identified during execution.",
]))
E.append(h2("4.2", "Roller System"))
E.append(blist([
    "Remove and replace damaged gangway rollers.",
    "Supply additional PVC rollers as operational spare stock.",
]))
E.append(h2("4.3", "Teflon Wear Protection"))
E.append(blist([
    "Remove damaged gangway flip protection pads.",
    "Supply and install new marine-grade Teflon wear pads.",
]))
E.append(h2("4.4", "Structural Steel Rehabilitation"))
E.append(Paragraph("Where applicable, the Contractor shall:", lead))
E.append(blist([
    "Clean steel surfaces by mechanical preparation or approved blasting method.",
    "Remove corrosion and contaminants.",
    "Apply approved marine-grade anti-corrosion coating system.",
    "Repaint gangway H-beam support structures.",
]))
E.append(h2("4.5", "Anchor Bolt Replacement"))
E.append(Paragraph("Where applicable, the Contractor shall:", lead))
E.append(blist([
    "Remove corroded foundation bolts and nuts.",
    "Supply and install new stainless steel or approved galvanized anchor bolts.",
    "Upgrade bolt sizing where specified to improve structural stability.",
]))
E.append(h2("4.6", "Electrical Safety Improvements"))
E.append(Paragraph("The Contractor shall:", lead))
E.append(blist([
    "Install weatherproof protection for all exposed dock gate electrical components.",
    "Provide suitable enclosures, sealing systems, covers and fixing arrangements.",
    "Inspect and rectify electrical grounding (earthing) systems for all gangways.",
    "Test and certify grounding resistance values following completion.",
]))
E.append(h2("4.7", "Cable Tray Works (Docks D and K)"))
E.append(Paragraph("The Contractor shall:", lead))
E.append(blist([
    "Secure under-gangway cable trays using approved marine-grade support systems.",
    "Supply and install all required fixing accessories.",
    "Ensure cable tray alignment and support integrity.",
]))
E.append(Paragraph("Based on requirements identified across all dock gangways.", body))

# ---------- 5. PONTOON JOINT PLATE REPAIRS ----------
E.append(h1("5", "Pontoon Joint Plate Repairs"))
E.append(applic_p("Dock A, B, C, D, K"))
E.append(h2("5.1", "Preferred Method (Option A)"))
E.append(blist([
    "Remove damaged joint plates.",
    "Drill two additional fixing holes on site.",
    "Install new fixings into pontoon concrete structure using approved plastic grips and marine-grade bolts.",
    "Seal all abandoned holes with approved marine repair compound.",
]))
E.append(h2("5.2", "Alternative Method (Option B)"))
E.append(blist(["Weld four support bolts to joint plates where approved by JYC."]))
E.append(Paragraph("Joint Plate Quantities", micro))
E.append(qty_table(
    ["Dock", "Quantity"],
    [["Dock A", "16"], ["Dock B", "30"], ["Dock C", "18"], ["Dock D", "24"], ["Dock K", "16"]],
    ["Total", "104 Joint Plates"],
))

# ---------- 6. PONTOON SURFACE & PILE REFURB ----------
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

# ---------- 7. FIBERGLASS PILE CAP ----------
E.append(h1("7", "Fiberglass Pile Cap Replacement"))
E.append(Paragraph("The Contractor shall supply and install new fiberglass pile caps as follows:", lead))
E.append(qty_table(
    ["Location", "Quantity"],
    [["Dock A", "9"], ["Dock B", "13"], ["Dock C", "9"], ["Dock D", "9"], ["Dock K", "5"]],
    ["Total", "45 Pile Caps"],
))
E.append(Paragraph("The Contractor shall:", lead))
E.append(blist([
    "Remove existing damaged pile caps.",
    "Supply marine-grade fiberglass pile caps.",
    "Install and secure pile caps in accordance with manufacturer recommendations.",
]))

# ---------- 8. PILE BRACKET & WEAR PAD ----------
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
E.append(Paragraph("Minimum Quantity", micro))
E.append(blist(["10 wear pads per dock, inclusive of spare stock."]))

# ---------- 9. WAVE BREAKER & NAV POLE ----------
E.append(h1("9", "Main Wave Breaker and Navigation Pole Rehabilitation"))
E.append(Paragraph("The Contractor shall carry out the following works:", lead))
E.append(h2("9.1", "Wave Breaker"))
E.append(blist([
    "Prepare exposed steel surfaces.",
    "Remove corrosion.",
    "Apply approved marine anti-corrosion coating system.",
    "Repaint complete wave breaker structure.",
]))
E.append(h2("9.2", "Navigation Poles"))
E.append(blist([
    "Remove corrosion.",
    "Repaint navigation poles.",
    "Replace heavily corroded foundation anchor bolts.",
]))
E.append(h2("9.3", "Manhole Repair"))
E.append(blist([
    "Remove temporary plywood cover.",
    "Supply and install permanent marine-grade manhole cover.",
    "Reinstate surrounding area.",
]))

# ---------- 10. QUAY WALL ----------
E.append(h1("10", "Quay Wall Surface Repainting"))
E.append(Paragraph("The Contractor shall:", lead))
E.append(blist([
    "Prepare top surface of quay wall surrounding the marina basin.",
    "Remove deteriorated coatings.",
    "Apply approved marine coating system.",
    "Repaint quay wall top surfaces only.",
]))

# ---------- 11. SLIPWAY PONTOON STRUCTURAL ----------
E.append(h1("11", "Slipway Pontoon Structural Repairs"))
E.append(Paragraph("The Contractor shall:", lead))
E.append(blist([
    "Inspect and repair stainless steel H-beam attached to quay wall.",
    "Reinstall and secure foundation anchor bolts.",
    "Repair damaged / spalled concrete around foundation areas.",
    "Reinstate concrete using approved marine repair materials.",
    "Ensure structural integrity and long-term durability of the installation.",
]))

E.append(Spacer(1, 12))
E.append(HRFlowable(width="40%", thickness=0.8, color=ACCENT, hAlign="CENTER", spaceAfter=6))
E.append(Paragraph("— End of Technical Submittal —",
                   S("end", fontName="Helvetica-Oblique", fontSize=9, textColor=LIGHT,
                     alignment=TA_CENTER)))

doc.build(E)
print("PDF built.")
