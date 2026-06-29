#!/usr/bin/env python3
"""Build the single MERGED Skylance Technical Submittal for JYC Marina.

Integrates three source submittals (organised by work scope):
  - JYC Marina Infrastructure (structured scope)
  - Marine Structure Rehabilitation (coating systems + product data sheets)
  - Gangway & Cable Tray Repairs (gangway components + Hilti + drawings)

The written body is generated in the Skylance house style; the original
manufacturer data sheets and gangway drawings are embedded as Appendices.
"""
import os
import fitz  # PyMuPDF for final assembly
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_JUSTIFY
from reportlab.platypus import (
    BaseDocTemplate, PageTemplate, Frame, Paragraph, Spacer, Table, TableStyle,
    ListFlowable, ListItem, NextPageTemplate, PageBreak, HRFlowable, Image, KeepTogether,
)
from reportlab.lib.styles import ParagraphStyle as PS
from reportlab.lib.utils import ImageReader

HERE = os.path.dirname(os.path.abspath(__file__))
ASSET = os.path.join(HERE, "brand_assets")
HEADER_IMG = os.path.join(ASSET, "crop_header.png")
FOOTER_IMG = os.path.join(ASSET, "crop_footer.png")
WM_IMG = os.path.join(ASSET, "crop_watermark.png")
SKYLANCE_IMG = os.path.join(ASSET, "skylance_logo.png")
SELA_IMG = os.path.join(ASSET, "sela_logo.png")

SRC_MARINE = os.path.join(HERE, "Marine Structure Rehabilitation - Technical Submittal.pdf")
SRC_GANGWAY = os.path.join(HERE, "Technical Submittal Skylance- Gangway Cable Tray Repairs Sela JYC.pdf")
SRC_CABLE16 = os.path.join(HERE, "1x16 mm² NCU-PVC 450-750 V (TD).pdf")
SRC_CABLE70 = os.path.join(HERE, "1x70 mm² NCU-PVC 450-750 V (TD).pdf")

OUT = os.path.join(HERE, "Technical Submittal - Marina Infrastructure Maintenance and Rehabilitation Works (Skylance).pdf")
BODY = os.path.join(HERE, "_body_main.pdf")

DOC_NO = "TS-SL-260629-M"
REV = "Rev A – For Review & Approval"
FOOTER_LINE = f"{DOC_NO}  |  {REV}  |  Confidential"

# ---------- palette ----------
CYAN = colors.HexColor("#26B8EF")
BLACK = colors.HexColor("#111111")
INK = colors.HexColor("#1F1F1F")
GREY = colors.HexColor("#3A3A3A")
SOFTGREY = colors.HexColor("#888888")
ROWGREY = colors.HexColor("#F2F2F2")
HAIR = colors.HexColor("#CFCFCF")
SOFT = colors.HexColor("#F2F6FA")

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


def S(name, **kw):
    return PS(name, **kw)


body = S("body", fontName="Helvetica", fontSize=10, leading=14.5, textColor=INK,
         alignment=TA_JUSTIFY, spaceAfter=7)
lead = S("lead", parent=body, spaceAfter=4, keepWithNext=1)
applic = S("applic", fontName="Helvetica-Oblique", fontSize=9.5, leading=13,
           textColor=GREY, spaceBefore=1, spaceAfter=7, keepWithNext=1)
bullet = S("bullet", fontName="Helvetica", fontSize=10, leading=13.5, textColor=INK, spaceAfter=0)


def h1(num, title):
    cell = Paragraph(f'<font color="white"><b>{num}.&nbsp;&nbsp;{title.upper()}</b></font>',
                     S("h1", fontName="Helvetica-Bold", fontSize=11.5, leading=15, textColor=colors.white))
    t = Table([[cell]], colWidths=[PAGE_W - LM - RM])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), BLACK),
        ("LINEBEFORE", (0, 0), (0, -1), 3.2, CYAN),
        ("TOPPADDING", (0, 0), (-1, -1), 6), ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("LEFTPADDING", (0, 0), (-1, -1), 10), ("RIGHTPADDING", (0, 0), (-1, -1), 8),
    ]))
    t.spaceBefore = 14
    t.spaceAfter = 8
    t.keepWithNext = 1
    return t


def h2(text):
    p = Paragraph(f'<b>{text}</b>', S("h2", fontName="Helvetica-Bold", fontSize=11, leading=14,
                                      textColor=CYAN, spaceBefore=10, spaceAfter=4, keepWithNext=1))
    p.keepWithNext = 1
    return p


def applic_p(text):
    p = Paragraph(f'<font color="#26B8EF"><b><i>Applicable to:&nbsp;&nbsp;</i></b></font>'
                  f'<i><font color="#3A3A3A">{text}</font></i>', applic)
    p.keepWithNext = 1
    return p


def micro(text):
    p = Paragraph(f'<b>{text}</b>', S("micro", fontName="Helvetica-Bold", fontSize=10, leading=13,
                                      textColor=INK, spaceBefore=4, spaceAfter=3, keepWithNext=1))
    p.keepWithNext = 1
    return p


def blist(items):
    li = [ListItem(Paragraph(t, bullet), leftIndent=14, spaceAfter=4, bulletColor=BLACK) for t in items]
    return ListFlowable(li, bulletType="bullet", start="•", bulletFontSize=8, bulletColor=BLACK,
                        leftIndent=18, bulletOffsetY=1, spaceBefore=1, spaceAfter=6)


def qty_table(headers, rows, total):
    data = [headers] + rows + [total]
    w = PAGE_W - LM - RM
    t = Table(data, colWidths=[w * 0.5, w * 0.5], hAlign="LEFT")
    style = [
        ("BACKGROUND", (0, 0), (-1, 0), BLACK), ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"), ("FONTSIZE", (0, 0), (-1, -1), 9.5),
        ("ALIGN", (1, 0), (1, -1), "CENTER"), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("FONTNAME", (0, 1), (0, -1), "Helvetica-Bold"), ("TEXTCOLOR", (0, 1), (-1, -2), INK),
        ("GRID", (0, 0), (-1, -1), 0.6, HAIR), ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5), ("LEFTPADDING", (0, 0), (-1, -1), 9),
        ("BACKGROUND", (0, -1), (-1, -1), BLACK), ("TEXTCOLOR", (0, -1), (-1, -1), CYAN),
        ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"),
    ]
    for i in range(1, len(rows) + 1):
        if i % 2 == 0:
            style.append(("BACKGROUND", (0, i), (-1, i), ROWGREY))
    t.setStyle(TableStyle(style))
    t.spaceAfter = 6
    return t


def spec_table(headers, rows, widths):
    w = PAGE_W - LM - RM
    cw = [w * x for x in widths]
    cells = [[Paragraph(f'<b><font color="white">{c}</font></b>',
                        S("th", fontName="Helvetica-Bold", fontSize=9.5, leading=12)) for c in headers]]
    for r in rows:
        cells.append([Paragraph(x, S("td", fontName="Helvetica", fontSize=9.5, leading=12, textColor=INK)) for x in r])
    t = Table(cells, colWidths=cw, hAlign="LEFT")
    style = [
        ("BACKGROUND", (0, 0), (-1, 0), BLACK), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("GRID", (0, 0), (-1, -1), 0.6, HAIR), ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5), ("LEFTPADDING", (0, 0), (-1, -1), 8),
    ]
    for i in range(1, len(rows) + 1):
        if i % 2 == 0:
            style.append(("BACKGROUND", (0, i), (-1, i), ROWGREY))
    t.setStyle(TableStyle(style))
    t.spaceAfter = 6
    return t


# ---------- page furniture ----------
def draw_brand(canvas, doc, cover=False):
    canvas.saveState()
    try:
        wm = ImageReader(WM_IMG); iw, ih = wm.getSize()
        tw = PAGE_W * 0.62; th = tw * ih / iw
        canvas.drawImage(wm, (PAGE_W - tw) / 2, (PAGE_H - th) / 2 - 10, width=tw, height=th, mask="auto")
    except Exception:
        pass
    if not cover:
        canvas.drawImage(HEADER_IMG, 0, PAGE_H - HDR_H, width=HDR_W, height=HDR_H, mask="auto")
        ry = PAGE_H - HDR_H - 12
        canvas.setStrokeColor(CYAN); canvas.setLineWidth(1.4)
        canvas.line(LM, ry, PAGE_W - RM, ry)
    canvas.drawImage(FOOTER_IMG, 0, 0, width=FTR_W, height=FTR_H, mask="auto")
    canvas.setFont("Helvetica", 7.5); canvas.setFillColor(SOFTGREY)
    canvas.drawCentredString(PAGE_W / 2, FTR_H + 6, f"Page {canvas.getPageNumber()}  |  {FOOTER_LINE}")
    canvas.restoreState()


def on_cover(c, d):
    draw_brand(c, d, cover=True)


def on_body(c, d):
    draw_brand(c, d, cover=False)


def cover_logo_lockup():
    sky_h = 62.0
    sky = Image(SKYLANCE_IMG, width=sky_h * 346.0 / 294.0, height=sky_h)
    sela_h = 27.0
    sela = Image(SELA_IMG, width=sela_h * 479.0 / 158.0, height=sela_h)
    lbl = S("lbl", fontName="Helvetica-Bold", fontSize=8, leading=11, textColor=CYAN, alignment=TA_CENTER)
    cont = Paragraph('CONTRACTOR<br/><font color="#3A3A3A">Skylance Technical Services Co.</font>', lbl)
    clnt = Paragraph('CLIENT<br/><font color="#3A3A3A">Jeddah Yacht Club &amp; Marina</font>', lbl)
    t = Table([[sky, "", sela], [cont, "", clnt]], colWidths=[210, 36, 210], hAlign="CENTER")
    t.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, 0), "MIDDLE"), ("VALIGN", (0, 1), (-1, 1), "TOP"),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"), ("LINEAFTER", (0, 0), (0, 0), 0.8, HAIR),
        ("TOPPADDING", (0, 1), (-1, 1), 12), ("BOTTOMPADDING", (0, 0), (-1, 0), 4),
    ]))
    return t


# =====================================================================
#  CONTENT  (sections 1-15)
# =====================================================================
# (num, title) used for both TOC and section-page detection
SECTIONS = [
    ("1", "Introduction"),
    ("2", "Objective"),
    ("3", "Scope of Works"),
    ("4", "Gangway and Cable Tray Rehabilitation Works"),
    ("5", "Pontoon Joint Plate Repairs"),
    ("6", "Pontoon Surface and Pile Refurbishment"),
    ("7", "Fiberglass Pile Cap Replacement"),
    ("8", "Pile Bracket and Wear Pad Repairs"),
    ("9", "Main Wave Breaker and Navigation Pole Rehabilitation"),
    ("10", "Quay Wall Repair and Coating"),
    ("11", "Slipway Pontoon Structural Repairs"),
    ("12", "Concrete Repair Works"),
    ("13", "Protective Coating and Material Systems"),
    ("14", "Deliverables"),
    ("15", "Appendices"),
]


def content_flowables():
    E = []

    # 1. INTRODUCTION
    E.append(h1("1", "Introduction"))
    E.append(Paragraph(
        "This Technical Submittal has been developed by Skylance Technical Services to describe the "
        "technical details of the Jeddah Yacht Club &amp; Marina (JYC) Marina Infrastructure Maintenance "
        "and Rehabilitation Works. It consolidates, into a single document, the inspection findings and "
        "the proposed repair, refurbishment, replacement, protective-coating and maintenance methodology "
        "for the marina floating assets, gangways and cable tray systems, wave breaker structure, quay "
        "wall, navigation poles, steel piles, concrete elements and associated marine assets.", body))
    E.append(Paragraph(
        "The submittal presents the structure-wise scope of work together with the proposed coating and "
        "repair material systems. The original manufacturer product data sheets and gangway drawings "
        "referenced throughout are reproduced in full under the Appendices. All works shall be carried "
        "out in accordance with applicable Saudi Arabian regulations, the project specifications and the "
        "coating / material manufacturers' recommendations.", body))

    # 2. OBJECTIVE
    E.append(h1("2", "Objective"))
    E.append(Paragraph(
        "The objective of this scope of work is to restore structural integrity, improve operational "
        "safety and appearance, prevent further corrosion and deterioration, and extend the service life "
        "of the marina assets through comprehensive maintenance, repair and protective-coating works.", body))
    E.append(h2("In delivering the works, the Contractor will:"))
    E.append(blist([
        "Conduct a detailed site survey and verification of quantities prior to commencement.",
        "Supply all labour, supervision, tools, equipment, materials, consumables and marine access equipment required.",
        "Dispose of waste materials in accordance with local environmental regulations.",
        "Maintain uninterrupted marina operations where reasonably practicable.",
        "Comply with all applicable Saudi Arabian safety regulations and marine industry standards.",
        "Submit method statements, risk assessments, work schedules and quality control procedures before mobilization.",
        "Offer warranty on all works completed.",
    ]))

    # 3. SCOPE OF WORKS
    E.append(h1("3", "Scope of Works"))
    E.append(Paragraph("The scope of work covers the following structures, marine assets and work types, "
                       "with the detailed scope prescribed in the sections that follow:", lead))
    E.append(h2("Scope Includes"))
    E.append(blist([
        "Gangway and cable tray rehabilitation across Docks A, B, C, D, K and the Slipway Pontoon.",
        "Floating pontoon surface, pile and pile-cap refurbishment and protective epoxy coating.",
        "Pontoon joint plate repairs, pile bracket and wear pad repairs.",
        "Main wave breaker and navigation pole rehabilitation and coating.",
        "Quay wall concrete repair, surface preparation and epoxy / protective coating.",
        "Slipway pontoon structural repairs.",
        "Concrete repair works and protective coating / material systems across all assets.",
    ]))
    E.append(Paragraph("The detailed scope is set out in Sections 4 through 13; supporting product data "
                       "sheets and drawings are provided under Section 15 (Appendices).", body))

    # 4. GANGWAY & CABLE TRAY
    E.append(h1("4", "Gangway and Cable Tray Rehabilitation Works"))
    E.append(applic_p("Dock A, B, C, D, K &amp; Slipway Pontoon"))

    E.append(h2("4.1  Timber Works and Varnishing"))
    E.append(blist([
        "Clean, prepare, sand and apply marine-grade protective coating / varnish to all gangway timber planks.",
        "Remove loose, peeling and deteriorated existing coatings and sand to a smooth, uniform finish.",
        "Replace any damaged timber sections identified during execution.",
    ]))
    E.append(Paragraph("Proposed timber coating material (or equivalent):", lead))
    E.append(spec_table(["Product", "Manufacturer", "Function"],
                        [["Woodshield Stain Exterior Gloss", "Jotun",
                          "Decorative and protective exterior wood coating providing weather resistance and enhanced appearance"],
                         ["Hempel's Polyurethane Gloss Varnish (or equivalent)", "Hempel",
                          "Protective polyurethane gloss varnish for exterior timber"]],
                        [0.34, 0.16, 0.50]))

    E.append(h2("4.2  Gangway Rollers"))
    E.append(Paragraph(
        "The gangway rollers were found worn, seized or showing surface corrosion, affecting smooth travel "
        "of the gangway along its track. The proposed scope comprises removal of the existing worn rollers, "
        "supply of new replacement rollers matching the original OEM specification (material, diameter and "
        "bearing type), and reinstallation complete with lubrication to restore free and even rolling movement. "
        "Additional PVC rollers shall be supplied as operational spare stock. Refer to the "
        "Pontoon Roller drawing in Appendix B.", body))

    E.append(h2("4.3  Gangway Ramp Rubbing Plate"))
    E.append(Paragraph(
        "The rubbing plate fitted to the gangway ramp, forming the wear-resistant landing surface at the "
        "shore / vessel interface, was found worn and locally deformed due to repeated contact during landing. "
        "The proposed scope comprises removal of the damaged plate, fabrication and supply of a new rubbing "
        "plate of matching material and thickness, and re-fixing in accordance with the original design to "
        "reinstate a smooth, protected landing surface. Refer to the Gangway Pontoon Ramp and "
        "Shore Ramp drawings in Appendix B.", body))

    E.append(h2("4.4  Gangway Shore Pivot and Pin"))
    E.append(Paragraph(
        "The shore-end pivot bracket and pivot pin, which allow the gangway to articulate at the shore "
        "connection, were found to exhibit wear, elongation and corrosion resulting in excessive play. The "
        "proposed scope comprises removal of the existing pivot pin and bushing, inspection and repair of the "
        "pivot bracket as required, and supply and installation of a new pivot pin and bushing to restore "
        "proper, secure articulation. Refer to the Gangway Shore Pivot drawing in Appendix B.", body))

    E.append(h2("4.5  Teflon Wear Protection"))
    E.append(blist([
        "Remove damaged gangway flip protection pads.",
        "Supply and install new marine-grade Teflon wear pads.",
    ]))

    E.append(h2("4.6  Structural Steel Rehabilitation and Coating"))
    E.append(Paragraph("Where applicable, the Contractor will clean steel surfaces by mechanical "
                       "preparation or approved blasting method, remove corrosion and contaminants, repaint "
                       "the gangway H-beam support structures and apply the approved marine-grade "
                       "anti-corrosion coating system below (or equivalent):", lead))
    E.append(spec_table(["Coat", "Product", "Function"],
                        [["1st Coat", "Hempaprime Multi 500", "Epoxy primer providing adhesion and corrosion protection"],
                         ["2nd Coat", "Hempaprime Multi 500", "High-build epoxy barrier coat for enhanced durability and protection"],
                         ["Top Coat", "Hempathane HS 55610", "Polyurethane finish providing UV resistance, colour retention and weather protection"]],
                        [0.18, 0.30, 0.52]))

    E.append(h2("4.7  Anchor Bolt Replacement and Gangway Fixing Anchors"))
    E.append(Paragraph(
        "The anchor bolts securing the gangway base / support structure to the supporting deck or platform "
        "were found loose, corroded and insufficient in holding capacity. The proposed scope comprises removal "
        "of the existing corroded foundation bolts and nuts, preparation of new anchor points, and installation "
        "of new stainless steel or approved galvanized anchor bolts, with bolt sizing upgraded where specified. "
        "Chemical / adhesive anchors shall be installed in accordance with the Hilti HIT-HY 200-R V3 anchor "
        "system technical data sheet provided in Appendix C.", body))

    E.append(h2("4.8  Electrical Safety Improvements"))
    E.append(blist([
        "Install weatherproof protection for all exposed dock gate electrical components.",
        "Provide suitable enclosures, sealing systems, covers and fixing arrangements.",
        "Inspect and rectify electrical grounding (earthing) systems for all gangways.",
        "Test and certify grounding resistance values following completion.",
    ]))

    E.append(h2("4.9  Cable Tray Works"))
    E.append(Paragraph(
        "The cable tray routed along the gangway structure, carrying power and control cabling, was found "
        "corroded and deformed with damaged or missing supports (notably Docks D and K). The proposed scope "
        "comprises removal of the affected cable tray sections and supports, supply and installation of new "
        "marine-grade cable tray and approved support systems of matching size and material, and reinstatement "
        "of the cabling with adequate securing to restore safe, protected cable routing and alignment. "
        "Cable tray fabrication data and drawings are provided in Appendix D, and the proposed "
        "replacement power cable data sheets (1×16 mm² and 1×70 mm² NCU-PVC 450/750 V to "
        "IEC 60227) in Appendix E.", body))

    # 5. JOINT PLATES
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
    E.append(qty_table(["Dock", "Quantity"],
                       [["Dock A", "16"], ["Dock B", "30"], ["Dock C", "18"], ["Dock D", "24"], ["Dock K", "16"]],
                       ["Total", "104 Joint Plates"]))

    # 6. PONTOON SURFACE & PILE
    E.append(h1("6", "Pontoon Surface and Pile Refurbishment"))
    E.append(applic_p("Dock A, B, C, K &amp; Slipway Pontoon"))
    E.append(Paragraph("The Contractor will prepare all pontoon surfaces, remove loose coatings and "
                       "corrosion, repaint all pontoon surfaces, and prepare and repaint the marina piles, "
                       "providing coating system data sheets and warranty information.", body))

    E.append(h2("6.1  Floating Pontoon Epoxy Coating"))
    E.append(Paragraph("Refurbishment of existing epoxy-coated floating pontoon surfaces through cleaning, "
                       "localized repair of damaged coating areas, surface preparation and application of a "
                       "protective epoxy top coat.", lead))
    E.append(micro("Surface Preparation"))
    E.append(blist([
        "Wash and clean the pontoon surface to remove dirt, grease, salts and marine contaminants.",
        "Remove loose, flaking or damaged coating by mechanical means and repair localized defects.",
        "Lightly abrade sound existing epoxy surfaces to promote adhesion where needed.",
        "Ensure the surface is clean, dry and free from contaminants before coating application.",
    ]))
    E.append(spec_table(["Coat", "Product", "Function"],
                        [["Top Coat", "Hempafloor Durable 300",
                          "High-performance epoxy finish providing abrasion resistance, durability and long-term protection"]],
                        [0.18, 0.30, 0.52]))

    E.append(h2("6.2  Steel Pile Protective Epoxy Coating"))
    E.append(Paragraph("Refurbishment and protection of exposed steel piles through removal of deteriorated "
                       "paint, surface preparation and application of a high-performance protective coating "
                       "system suitable for marine conditions. Underwater / wet areas are not included.", lead))
    E.append(micro("Surface Preparation"))
    E.append(blist([
        "Remove existing loose and deteriorated paint by mechanical means.",
        "Remove rust, scale, dirt, grease, marine growth and other contaminants.",
        "Prepare steel surfaces to the required cleanliness standard per project specification and manufacturer recommendations.",
    ]))
    E.append(spec_table(["Coat", "Product", "Function"],
                        [["Base Coat", "Hempadur Multi-Strength GF 35870",
                          "High-build epoxy providing excellent adhesion, corrosion resistance and barrier protection"],
                         ["Top Coat", "Hempathane HS 55610",
                          "Durable polyurethane finish providing UV resistance, colour retention and long-term weather protection"]],
                        [0.18, 0.32, 0.50]))

    # 7. FIBERGLASS PILE CAP
    E.append(h1("7", "Fiberglass Pile Cap Replacement"))
    E.append(Paragraph("The Contractor will supply and install new fiberglass pile caps as follows:", lead))
    E.append(qty_table(["Location", "Quantity"],
                       [["Dock A", "9"], ["Dock B", "13"], ["Dock C", "9"], ["Dock D", "9"], ["Dock K", "5"]],
                       ["Total", "45 Pile Caps"]))
    E.append(Paragraph("The Contractor will:", lead))
    E.append(blist([
        "Remove existing damaged pile caps.",
        "Supply marine-grade fiberglass pile caps.",
        "Install and secure pile caps in accordance with manufacturer recommendations.",
    ]))

    # 8. PILE BRACKET & WEAR PAD
    E.append(h1("8", "Pile Bracket and Wear Pad Repairs"))
    E.append(applic_p("Dock D &amp; K"))
    E.append(Paragraph("The Contractor will:", lead))
    E.append(blist([
        "Remove damaged wear pads.",
        "Supply and install replacement wear pads.",
        "Adjust pile bracket fixing hole alignment.",
        "Rectify bracket positioning to prevent premature wear.",
        "Supply spare wear pads for future maintenance.",
    ]))
    E.append(micro("Minimum Quantity"))
    E.append(blist(["10 wear pads per dock, inclusive of spare stock."]))

    # 9. WAVE BREAKER & NAV POLE
    E.append(h1("9", "Main Wave Breaker and Navigation Pole Rehabilitation"))
    E.append(h2("9.1  Wave Breaker"))
    E.append(blist([
        "Prepare exposed steel surfaces and remove corrosion.",
        "Apply approved marine anti-corrosion coating system (refer Section 13).",
        "Repaint complete wave breaker structure.",
    ]))
    E.append(h2("9.2  Navigation Poles"))
    E.append(blist([
        "Remove corrosion and repaint navigation poles using the steel structure coating system (refer Section 13).",
        "Replace heavily corroded foundation anchor bolts (refer Appendix C).",
    ]))
    E.append(h2("9.3  Manhole Repair"))
    E.append(blist([
        "Remove temporary plywood cover.",
        "Supply and install permanent marine-grade manhole cover.",
        "Reinstate surrounding area.",
    ]))
    E.append(Paragraph("Exposed steel surfaces shall receive the steel structure epoxy coating system "
                       "described in Section 13.", body))

    # 10. QUAY WALL
    E.append(h1("10", "Quay Wall Repair and Coating"))
    E.append(h2("10.1  Surface Repainting"))
    E.append(blist([
        "Prepare top surface of quay wall surrounding the marina basin.",
        "Remove deteriorated coatings.",
        "Apply approved marine coating system and repaint quay wall top surfaces.",
    ]))
    E.append(h2("10.2  Concrete Repair and Epoxy Coating"))
    E.append(micro("Surface Preparation"))
    E.append(blist([
        "Removal of laitance, loose particles, dirt, oil, grease and contaminants.",
        "Mechanical grinding or abrasive preparation to achieve a sound and clean substrate.",
        "Concrete repairs to be completed and cured prior to coating application.",
    ]))
    E.append(spec_table(["Coat", "Product", "Function"],
                        [["Primer / Sealer", "Hempel's Sealer 05990",
                          "Penetrates and seals concrete substrate, promoting adhesion of subsequent coats"],
                         ["Intermediate Coat", "Hempafloor Mastic FC458",
                          "High-build epoxy providing excellent adhesion, chemical resistance and durability"],
                         ["Finish Coat", "Hempafloor Durable 300",
                          "Wear-resistant epoxy finish providing long-term protection and abrasion resistance"]],
                        [0.20, 0.28, 0.52]))

    # 11. SLIPWAY PONTOON
    E.append(h1("11", "Slipway Pontoon Structural Repairs"))
    E.append(Paragraph("The Contractor will:", lead))
    E.append(blist([
        "Inspect and repair stainless steel H-beam attached to quay wall.",
        "Reinstall and secure foundation anchor bolts (refer Appendix C).",
        "Repair damaged / spalled concrete around foundation areas.",
        "Reinstate concrete using approved marine repair materials (refer Section 12).",
        "Ensure structural integrity and long-term durability of the installation.",
    ]))

    # 12. CONCRETE REPAIR WORKS
    E.append(h1("12", "Concrete Repair Works"))
    E.append(Paragraph("Repair and rehabilitation of deteriorated concrete elements including treatment of "
                       "cracks, honeycombs, surface defects, voids and localized damaged areas, together with "
                       "grouting, surface preparation, reprofiling and finishing to restore structural "
                       "integrity and serviceability.", body))
    E.append(micro("Proposed Repair Materials"))
    E.append(spec_table(["Application", "Proposed Material"],
                        [["Crack Filling / Epoxy Repair", "CEMTEC R44 Epoxy Gel"],
                         ["Bonding Agent / Epoxy Adhesive", "EPO PRIME LP"],
                         ["Vertical &amp; Overhead Concrete Repair", "REPCON SHB"],
                         ["Fair-Faced Repair Mortar / General Repair", "REPCON FRM"],
                         ["Non-Shrink Grouting", "CEMTEC NS GROUT"],
                         ["Fine Surface Finishing", "REPCON FC FINE"],
                         ["Surface Levelling / Epoxy Floor Filling", "Hempafloor Fill 200"]],
                        [0.52, 0.48]))

    # 13. COATING & MATERIAL SYSTEMS
    E.append(h1("13", "Protective Coating and Material Systems"))
    E.append(Paragraph("The protective coating and repair material systems proposed across the works are "
                       "consolidated below. Products are indicative; equivalent approved products may be "
                       "submitted. The full manufacturer data sheets are provided in Appendix A.", body))
    E.append(micro("Coating System Summary"))
    E.append(spec_table(["Asset / Application", "System", "Section"],
                        [["Quay wall concrete", "Hempel's Sealer 05990 → Hempafloor Mastic FC458 → Hempafloor Durable 300", "10.2"],
                         ["Floating pontoon", "Hempafloor Durable 300 (epoxy top coat)", "6.1"],
                         ["Steel piles", "Hempadur Multi-Strength GF 35870 → Hempathane HS 55610", "6.2"],
                         ["Steel structures / wave breaker", "Hempaprime Multi 500 (×2) → Hempathane HS 55610", "4.6, 9"],
                         ["Gangway timber", "Woodshield Stain Exterior Gloss / Hempel PU Gloss Varnish", "4.1"],
                         ["Concrete repair materials", "REPCON / CEMTEC / EPO PRIME range", "12"]],
                        [0.30, 0.55, 0.15]))
    E.append(micro("Product Data Sheet Schedule (Appendix A)"))
    E.append(spec_table(["Ref.", "Product Data Sheet", "Related Section"],
                        [["A1", "Hempel's Sealer 05990", "10 – Quay Wall"],
                         ["A2", "Hempafloor Mastic FC458", "10 – Quay Wall"],
                         ["A3", "Hempafloor Durable 300", "6, 10 – Pontoon / Quay Wall"],
                         ["A4", "Hempadur Multi-Strength GF 35870", "6 – Steel Pile"],
                         ["A5", "Hempathane HS 55610", "4, 6, 9 – Steel"],
                         ["A6", "Hempaprime Multi 500", "4, 9 – Steel Structure"],
                         ["A7", "Woodshield Stain Exterior Gloss", "4 – Gangway Timber"],
                         ["A8", "REPCON FC FINE", "12 – Concrete Repair"],
                         ["A9", "CEMTEC NS GROUT", "12 – Concrete Repair"],
                         ["A10", "REPCON FRM", "12 – Concrete Repair"],
                         ["A11", "REPCON SHB", "12 – Concrete Repair"],
                         ["A12", "EPO PRIME LP", "12 – Concrete Repair"],
                         ["A13", "CEMTEC R44 Epoxy Gel", "12 – Concrete Repair"]],
                        [0.10, 0.55, 0.35]))
    E.append(Paragraph("Note: Hempafloor Fill 200 and Hempel's Polyurethane Gloss Varnish are referenced "
                       "above; their manufacturer data sheets will be provided separately upon issue.",
                       S("note", fontName="Helvetica-Oblique", fontSize=8.5, leading=12, textColor=SOFTGREY)))

    # 14. DELIVERABLES
    E.append(h1("14", "Deliverables"))
    E.append(Paragraph("The Contractor will provide:", lead))
    E.append(blist([
        "Detailed work programme.",
        "Material submittals and technical data sheets.",
        "Quality assurance and inspection records.",
        "Coating specifications and warranties.",
        "Electrical test certificates.",
        "Completion report including photographs.",
        "As-built documentation where modifications have occurred.",
    ]))

    # 15. APPENDICES (register; original pages embedded after this body)
    E.append(KeepTogether([
        h1("15", "Appendices"),
        Paragraph("The following supporting documents are reproduced in full on the pages that follow "
                  "this section, in the order listed:", lead),
        spec_table(["Appendix", "Contents"],
                   [["A", "Product Data Sheets — Hempel, Jotun, REPCON and CEMTEC products (A1–A13)"],
                    ["B", "Gangway Drawings — GA / repair details (roller, ramp, shore pivot) and roller guide / wear plate"],
                    ["C", "Hilti Anchor System Technical Data — HIT-HY 200-R V3 Injection Mortar"],
                    ["D", "Cable Tray Data — SS316 cable tray, covers and slotted C-channel support"],
                    ["E", "Cable Technical Data Sheets — 1×16 mm² and 1×70 mm² NCU-PVC 450/750 V (IEC 60227)"]],
                   [0.16, 0.84]),
    ]))
    return E


# =====================================================================
#  BODY BUILD  (cover + TOC + sections)
# =====================================================================
def build_body(toc_rows, out_path):
    doc = BaseDocTemplate(out_path, pagesize=A4, leftMargin=LM, rightMargin=RM,
                          topMargin=TM, bottomMargin=BM,
                          title="Technical Submittal - Marina Infrastructure Maintenance & Rehabilitation Works",
                          author="Skylance Technical Services")
    frame = Frame(LM, FRAME_BOTTOM, PAGE_W - LM - RM, FRAME_TOP - FRAME_BOTTOM, id="f")
    COVER_TOP = PAGE_H - 1.5 * cm
    cframe = Frame(LM, FRAME_BOTTOM, PAGE_W - LM - RM, COVER_TOP - FRAME_BOTTOM, id="cf")
    doc.addPageTemplates([
        PageTemplate(id="Cover", frames=[cframe], onPage=on_cover),
        PageTemplate(id="Body", frames=[frame], onPage=on_body),
    ])
    E = []
    # ----- COVER -----
    E.append(Spacer(1, 18))
    E.append(cover_logo_lockup())
    E.append(HRFlowable(width="100%", thickness=1.4, color=CYAN, spaceBefore=14, spaceAfter=0))
    E.append(Spacer(1, 110))
    E.append(HRFlowable(width="86%", thickness=0.8, color=HAIR, hAlign="CENTER", spaceAfter=18))
    E.append(Paragraph("TECHNICAL SUBMITTAL", S("ttl", fontName="Helvetica-Bold", fontSize=29, leading=34,
                                                textColor=BLACK, alignment=TA_CENTER, spaceAfter=10)))
    E.append(Paragraph("Marina Infrastructure Maintenance and Rehabilitation Works",
                       S("sub", fontName="Helvetica-Bold", fontSize=13, leading=17, textColor=CYAN,
                         alignment=TA_CENTER, spaceAfter=5)))
    E.append(Paragraph("Jeddah Yacht Club &amp; Marina (JYC)",
                       S("sub2", fontName="Helvetica", fontSize=11.5, leading=15, textColor=GREY,
                         alignment=TA_CENTER, spaceAfter=18)))
    E.append(HRFlowable(width="86%", thickness=1.4, color=CYAN, hAlign="CENTER", spaceAfter=22))
    ctl = [("Project", "Jeddah Yacht Club & Marina (JYC)"),
           ("Client / Consultant", "SELA"),
           ("Document No.", DOC_NO),
           ("Revision", REV),
           ("Prepared By", "SKYLANCE TECHNICAL SERVICES"),
           ("Date", "29 Jun 2026"),
           ("Classification", "Technical Submittal – For Review & Approval")]
    tbl = Table([[Paragraph(f'<b><font color="white">{k}</font></b>', S("k", fontName="Helvetica-Bold", fontSize=9.5, leading=12)),
                  Paragraph(v, S("v", fontName="Helvetica", fontSize=9.5, leading=12, textColor=INK))] for k, v in ctl],
                colWidths=[(PAGE_W - LM - RM) * 0.30, (PAGE_W - LM - RM) * 0.70], hAlign="CENTER")
    cs = [("BACKGROUND", (0, 0), (0, -1), BLACK), ("GRID", (0, 0), (-1, -1), 0.6, HAIR),
          ("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("TOPPADDING", (0, 0), (-1, -1), 7),
          ("BOTTOMPADDING", (0, 0), (-1, -1), 7), ("LEFTPADDING", (0, 0), (-1, -1), 9)]
    for i in range(len(ctl)):
        if i % 2 == 1:
            cs.append(("BACKGROUND", (1, i), (1, i), SOFT))
    tbl.setStyle(TableStyle(cs))
    E.append(tbl)
    E.append(Spacer(1, 16))
    E.append(Paragraph("This document is prepared by Skylance Technical Services for submission to the "
                       "Client / Consultant for review and approval. It consolidates three source submittals "
                       "(Marina Infrastructure scope, Marine Structure Rehabilitation, and Gangway &amp; Cable "
                       "Tray Repairs) into a single technical submittal.",
                       S("disc", fontName="Helvetica-Oblique", fontSize=8.5, leading=12, textColor=SOFTGREY,
                         alignment=TA_CENTER)))
    E.append(NextPageTemplate("Body"))
    E.append(PageBreak())

    # ----- TABLE OF CONTENTS -----
    hdr = Table([[Paragraph('<b><font color="white">TABLE OF CONTENTS</font></b>',
                            S("toch", fontName="Helvetica-Bold", fontSize=11, leading=14))]],
                colWidths=[PAGE_W - LM - RM])
    hdr.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), BLACK), ("LINEBEFORE", (0, 0), (0, -1), 3.2, CYAN),
                             ("TOPPADDING", (0, 0), (-1, -1), 6), ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
                             ("LEFTPADDING", (0, 0), (-1, -1), 10)]))
    E.append(hdr)
    E.append(Spacer(1, 4))
    w = PAGE_W - LM - RM
    rows = []
    for num, title, pg in toc_rows:
        rows.append([Paragraph(f'<b>{num}</b>', S("tn", fontName="Helvetica-Bold", fontSize=9.5, textColor=INK)),
                     Paragraph(title, S("tt", fontName="Helvetica", fontSize=9.5, textColor=GREY)),
                     Paragraph(str(pg), S("tp", fontName="Helvetica", fontSize=9.5, textColor=GREY))])
    toc = Table(rows, colWidths=[w * 0.07, w * 0.83, w * 0.10])
    ts = [("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("LINEBELOW", (0, 0), (-1, -1), 0.5, HAIR),
          ("ALIGN", (2, 0), (2, -1), "RIGHT"), ("LINEBEFORE", (1, 0), (1, -1), 2.2, CYAN),
          ("TOPPADDING", (0, 0), (-1, -1), 7), ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
          ("LEFTPADDING", (1, 0), (1, -1), 10)]
    for i in range(len(rows)):
        if i % 2 == 1:
            ts.append(("BACKGROUND", (0, i), (-1, i), ROWGREY))
    toc.setStyle(TableStyle(ts))
    E.append(toc)
    E.append(Paragraph("Appendices A, B and C follow Section 15; the embedded data sheets and drawings "
                       "retain their original pagination.",
                       S("tocn", fontName="Helvetica-Oblique", fontSize=8, leading=11, textColor=SOFTGREY,
                         alignment=TA_LEFT, spaceBefore=6)))
    E.append(PageBreak())

    # ----- SECTIONS -----
    for f in content_flowables():
        E.append(f)

    doc.build(E)


# =====================================================================
#  APPENDIX DIVIDER PAGE
# =====================================================================
def build_divider(letter, title, items, abs_page, out_path):
    c = __import__("reportlab.pdfgen.canvas", fromlist=["canvas"]).Canvas(out_path, pagesize=A4)
    # watermark + footer + header band
    draw_brand_canvas(c, abs_page)
    # big appendix label
    c.setFillColor(CYAN)
    c.setFont("Helvetica-Bold", 13)
    c.drawString(LM, PAGE_H - HDR_H - 70, f"APPENDIX {letter}")
    c.setFillColor(BLACK)
    c.setFont("Helvetica-Bold", 22)
    c.drawString(LM, PAGE_H - HDR_H - 100, title)
    c.setStrokeColor(CYAN); c.setLineWidth(1.4)
    c.line(LM, PAGE_H - HDR_H - 114, PAGE_W - RM, PAGE_H - HDR_H - 114)
    y = PAGE_H - HDR_H - 150
    c.setFont("Helvetica", 10.5); c.setFillColor(INK)
    for it in items:
        c.setFillColor(CYAN); c.drawString(LM, y, "•")
        c.setFillColor(INK); c.drawString(LM + 14, y, it)
        y -= 18
    c.showPage()
    c.save()


def draw_brand_canvas(c, abs_page):
    try:
        wm = ImageReader(WM_IMG); iw, ih = wm.getSize()
        tw = PAGE_W * 0.62; th = tw * ih / iw
        c.drawImage(wm, (PAGE_W - tw) / 2, (PAGE_H - th) / 2 - 10, width=tw, height=th, mask="auto")
    except Exception:
        pass
    c.drawImage(HEADER_IMG, 0, PAGE_H - HDR_H, width=HDR_W, height=HDR_H, mask="auto")
    ry = PAGE_H - HDR_H - 12
    c.setStrokeColor(CYAN); c.setLineWidth(1.4); c.line(LM, ry, PAGE_W - RM, ry)
    c.drawImage(FOOTER_IMG, 0, 0, width=FTR_W, height=FTR_H, mask="auto")
    c.setFont("Helvetica", 7.5); c.setFillColor(SOFTGREY)
    c.drawCentredString(PAGE_W / 2, FTR_H + 6, f"Page {abs_page}  |  {FOOTER_LINE}")


# =====================================================================
#  MAIN — two-pass for TOC page numbers, then assemble with appendices
# =====================================================================
def find_section_pages(pdf_path):
    d = fitz.open(pdf_path)
    pages = {}
    for num, title in SECTIONS:
        target = f"{num}.  {title.upper()}"
        for pi in range(d.page_count):
            txt = d[pi].get_text()
            if target in txt:
                pages[num] = pi + 1
                break
    d.close()
    return pages


# pass 1: placeholder page numbers
toc_pass1 = [(num, title, "") for num, title in SECTIONS]
build_body(toc_pass1, BODY)
sec_pages = find_section_pages(BODY)
n_main = fitz.open(BODY).page_count

# appendix page geometry in final document
A_div = n_main + 1
B_div = A_div + 1 + 40          # +divA +40 marine sheets
C_div = B_div + 1 + 5           # +divB +5 gangway drawings (incl. roller guide / wear plate)
D_div = C_div + 1 + 27          # +divC +27 Hilti pages (trimmed to core datasheet)
E_div = D_div + 1 + 5           # +divD +5 cable tray data pages (gangway pp 55-59)

# pass 2: real page numbers
toc_pass2 = []
for num, title in SECTIONS:
    toc_pass2.append((num, title, sec_pages.get(num, "")))
build_body(toc_pass2, BODY)

# dividers
DA = os.path.join(HERE, "_divA.pdf"); DB = os.path.join(HERE, "_divB.pdf")
DC = os.path.join(HERE, "_divC.pdf"); DD = os.path.join(HERE, "_divD.pdf")
DE = os.path.join(HERE, "_divE.pdf")
build_divider("A", "Product Data Sheets",
              ["Hempel, Jotun, REPCON and CEMTEC product data sheets (Refs A1–A13).",
               "Referenced throughout Sections 4, 6, 9, 10 and 12 of this submittal."], A_div, DA)
build_divider("B", "Gangway Drawings",
              ["General Arrangement and repair details for the gangway assemblies.",
               "Pontoon roller, ramp, shore pivot and shore ramp drawings.",
               "Roller guide and wear plate details."], B_div, DB)
build_divider("C", "Hilti Anchor System Technical Data",
              ["Hilti HIT-HY 200-R V3 injection mortar anchor system.",
               "Supports the gangway fixing anchor bolt works (Section 4.7)."], C_div, DC)
build_divider("D", "Cable Tray Data",
              ["SS316 cable tray (P-Type) and covers — 300 mm and 400 mm sizes.",
               "SS316 slotted C-channel support.",
               "Supports the cable tray reinstatement works (Section 4.9)."], D_div, DD)
build_divider("E", "Cable Technical Data Sheets",
              ["1×16 mm² and 1×70 mm² NCU-PVC 450/750 V power cables (IEC 60227).",
               "Supports the cabling reinstatement works (Section 4.9)."], E_div, DE)

# assemble final
final = fitz.open()
final.insert_pdf(fitz.open(BODY))
final.insert_pdf(fitz.open(DA))
final.insert_pdf(fitz.open(SRC_MARINE), from_page=10, to_page=49)   # pp 11-50 (13 data sheets)
final.insert_pdf(fitz.open(DB))
final.insert_pdf(fitz.open(SRC_GANGWAY), from_page=3, to_page=6)    # pp 4-7 drawings
final.insert_pdf(fitz.open(SRC_GANGWAY), from_page=59, to_page=59)  # p 60 roller guide / wear plate drawing
final.insert_pdf(fitz.open(DC))
final.insert_pdf(fitz.open(SRC_GANGWAY), from_page=7, to_page=33)   # pp 8-34 Hilti (core datasheet only)
final.insert_pdf(fitz.open(DD))                                     # App D divider: Cable Tray Data
final.insert_pdf(fitz.open(SRC_GANGWAY), from_page=54, to_page=58)  # pp 55-59 cable tray data
final.insert_pdf(fitz.open(DE))                                     # App E divider: Cable Technical Data Sheets
final.insert_pdf(fitz.open(SRC_CABLE16))                            # 1x16 mm² cable TD
final.insert_pdf(fitz.open(SRC_CABLE70))                            # 1x70 mm² cable TD
final.save(OUT)
print("MERGED PDF:", OUT)
print("body pages:", n_main, "| total pages:", final.page_count)
print("Appendix A:", A_div, "B:", B_div, "C:", C_div, "D:", D_div, "E:", E_div)

# cleanup temp files
for f in (BODY, DA, DB, DC, DD, DE):
    try:
        os.remove(f)
    except OSError:
        pass
