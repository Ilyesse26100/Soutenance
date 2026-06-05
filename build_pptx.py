# -*- coding: utf-8 -*-
"""
Génération du PowerPoint de soutenance — Projet Aphélie
BUT Réseaux & Télécommunications, parcours Cybersécurité — Ilyesse KEBAILI
"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

# ---------------------------------------------------------------- Palette
NAVY      = RGBColor(0x0F, 0x1B, 0x2D)   # fond sombre / titres
BLUE      = RGBColor(0x1B, 0x3A, 0x5C)   # bleu profond
ACCENT    = RGBColor(0x00, 0xB4, 0xC8)   # cyan tech (accent principal)
ACCENT2   = RGBColor(0x2E, 0xC4, 0x8F)   # vert (ok / dispo)
RED       = RGBColor(0xE5, 0x3E, 0x3E)   # rouge alarme
AMBER     = RGBColor(0xF2, 0xA6, 0x3B)   # ambre
WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT     = RGBColor(0xEC, 0xF2, 0xF7)   # fond clair
GREY      = RGBColor(0x6B, 0x7A, 0x8C)   # texte secondaire
DARKTXT   = RGBColor(0x1B, 0x26, 0x33)   # texte principal sombre
CARD      = RGBColor(0xF4, 0xF7, 0xFA)   # carte claire

FONT = "Arial"
FONT_H = "Arial"

prs = Presentation()
prs.slide_width  = Inches(13.333)
prs.slide_height = Inches(7.5)
SW, SH = prs.slide_width, prs.slide_height
BLANK = prs.slide_layouts[6]

# ---------------------------------------------------------------- Helpers
def slide():
    return prs.slides.add_slide(BLANK)

def rect(s, x, y, w, h, fill=None, line=None, line_w=None, shape=MSO_SHAPE.RECTANGLE):
    sp = s.shapes.add_shape(shape, x, y, w, h)
    sp.shadow.inherit = False
    if fill is None:
        sp.fill.background()
    else:
        sp.fill.solid(); sp.fill.fore_color.rgb = fill
    if line is None:
        sp.line.fill.background()
    else:
        sp.line.color.rgb = line
        sp.line.width = line_w or Pt(1)
    return sp

def txt(s, x, y, w, h, runs, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
        space_after=6, line_spacing=1.0, wrap=True):
    """runs: list of paragraphs; each paragraph is list of (text, size, color, bold, italic)."""
    tb = s.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = wrap
    tf.vertical_anchor = anchor
    tf.margin_left = 0; tf.margin_right = 0
    tf.margin_top = 0; tf.margin_bottom = 0
    for i, para in enumerate(runs):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.space_after = Pt(space_after)
        p.space_before = Pt(0)
        p.line_spacing = line_spacing
        for (t, sz, col, bold, ital) in para:
            r = p.add_run(); r.text = t
            r.font.size = Pt(sz); r.font.color.rgb = col
            r.font.bold = bold; r.font.italic = ital
            r.font.name = FONT
    return tb

def P(*runs):  # paragraph builder
    return list(runs)
def R(t, sz, col, bold=False, ital=False):
    return (t, sz, col, bold, ital)

def bullet(s, x, y, w, items, size=15, color=DARKTXT, gap=10, marker_col=ACCENT,
           lh=1.05, mk="▸"):
    tb = s.shapes.add_textbox(x, y, w, Inches(0.5))
    tf = tb.text_frame; tf.word_wrap = True
    tf.margin_left=0; tf.margin_right=0; tf.margin_top=0; tf.margin_bottom=0
    for i, it in enumerate(items):
        p = tf.paragraphs[0] if i==0 else tf.add_paragraph()
        p.space_after = Pt(gap); p.space_before = Pt(0); p.line_spacing = lh
        if isinstance(it, tuple):
            head, rest = it
            r = p.add_run(); r.text = mk+"  "; r.font.size=Pt(size); r.font.color.rgb=marker_col; r.font.bold=True; r.font.name=FONT
            r = p.add_run(); r.text = head; r.font.size=Pt(size); r.font.color.rgb=color; r.font.bold=True; r.font.name=FONT
            if rest:
                r = p.add_run(); r.text = rest; r.font.size=Pt(size); r.font.color.rgb=color; r.font.bold=False; r.font.name=FONT
        else:
            r = p.add_run(); r.text = mk+"  "; r.font.size=Pt(size); r.font.color.rgb=marker_col; r.font.bold=True; r.font.name=FONT
            r = p.add_run(); r.text = it; r.font.size=Pt(size); r.font.color.rgb=color; r.font.name=FONT
    return tb

def gradient_navy(sp, c1=NAVY, c2=BLUE, angle=45):
    """Apply a 2-stop linear gradient to a shape's fill."""
    spPr = sp.fill._xPr
    for tag in ('a:noFill','a:solidFill','a:gradFill','a:blipFill','a:pattFill','a:grpFill'):
        e = spPr.find(qn(tag))
        if e is not None: spPr.remove(e)
    grad = spPr.makeelement(qn('a:gradFill'), {})
    lst = grad.makeelement(qn('a:gsLst'), {})
    for pos, col in ((0, c1), (100000, c2)):
        gs = grad.makeelement(qn('a:gs'), {'pos': str(pos)})
        clr = grad.makeelement(qn('a:srgbClr'), {'val': '%02X%02X%02X' % (col[0], col[1], col[2])})
        gs.append(clr); lst.append(gs)
    grad.append(lst)
    lin = grad.makeelement(qn('a:lin'), {'ang': str(int(angle*60000)), 'scaled':'1'})
    grad.append(lin)
    # insert before a:ln if present
    ln = spPr.find(qn('a:ln'))
    if ln is not None: ln.addprevious(grad)
    else: spPr.append(grad)

def chapter_header(s, num, title, kicker=None, badge=None):
    """Standard content-slide header band."""
    band = rect(s, 0, 0, SW, Inches(1.18), fill=NAVY)
    gradient_navy(band, NAVY, BLUE, angle=0)
    # accent number chip
    rect(s, Inches(0.5), Inches(0.27), Inches(0.62), Inches(0.62), fill=ACCENT)
    txt(s, Inches(0.5), Inches(0.27), Inches(0.62), Inches(0.62),
        [P(R(num, 26, NAVY, True))], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    sub = [P(R(title, 23, WHITE, True))]
    if kicker:
        sub = [P(R(kicker.upper(), 10.5, ACCENT, True)), P(R(title, 22, WHITE, True))]
    txt(s, Inches(1.32), Inches(0.20), Inches(9.6), Inches(0.85), sub, anchor=MSO_ANCHOR.MIDDLE)
    if badge:
        bw = Inches(2.05)
        rect(s, SW-bw-Inches(0.5), Inches(0.34), bw, Inches(0.5), fill=ACCENT2,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        txt(s, SW-bw-Inches(0.5), Inches(0.34), bw, Inches(0.5),
            [P(R(badge, 12, NAVY, True))], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    # thin accent rule
    rect(s, 0, Inches(1.18), SW, Pt(3), fill=ACCENT)

def footer(s, page):
    txt(s, Inches(0.5), Inches(7.06), Inches(6), Inches(0.34),
        [P(R("Aphélie — Supervision centralisée des faisceaux hertziens", 8.5, GREY, False))],
        anchor=MSO_ANCHOR.MIDDLE)
    txt(s, Inches(10.4), Inches(7.06), Inches(2.43), Inches(0.34),
        [P(R("Ilyesse KEBAILI   |   " + str(page), 8.5, GREY, False))],
        align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)

def card(s, x, y, w, h, fill=CARD, line=None):
    r = rect(s, x, y, w, h, fill=fill, line=line, line_w=Pt(1) if line else None,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    try:
        r.adjustments[0] = 0.06
    except Exception:
        pass
    return r

def kpi(s, x, y, w, value, label, vcol=ACCENT, h=Inches(1.5)):
    card(s, x, y, w, h, fill=WHITE, line=RGBColor(0xDD,0xE6,0xEE))
    rect(s, x, y, Inches(0.10), h, fill=vcol)
    txt(s, x+Inches(0.28), y+Inches(0.16), w-Inches(0.4), Inches(0.7),
        [P(R(value, 30, vcol, True))], anchor=MSO_ANCHOR.MIDDLE)
    txt(s, x+Inches(0.28), y+Inches(0.86), w-Inches(0.4), h-Inches(0.95),
        [P(R(label, 11.5, DARKTXT, False))], anchor=MSO_ANCHOR.TOP, line_spacing=1.0)

# ================================================================ SLIDE 1 — TITLE
s = slide()
bg = rect(s, 0, 0, SW, SH, fill=NAVY)
gradient_navy(bg, NAVY, BLUE, angle=60)
# decorative hexagon-ish accent lines (network motif) — diagonal accent band
rect(s, 0, 0, Inches(0.22), SH, fill=ACCENT)
# small dots/links motif top-right
for (dx, dy) in [(11.7,0.9),(12.4,1.5),(11.2,1.7),(12.7,0.7),(10.9,1.2)]:
    rect(s, Inches(dx), Inches(dy), Inches(0.13), Inches(0.13), fill=ACCENT, shape=MSO_SHAPE.OVAL)
rect(s, Inches(0.85), Inches(0.75), Inches(3.4), Pt(2.5), fill=ACCENT)
txt(s, Inches(0.85), Inches(0.5), Inches(11), Inches(0.4),
    [P(R("BUT RÉSEAUX & TÉLÉCOMMUNICATIONS  —  PARCOURS CYBERSÉCURITÉ", 12, ACCENT, True))])
txt(s, Inches(0.85), Inches(0.92), Inches(11), Inches(0.4),
    [P(R("Mémoire de réalisation professionnelle  ·  Session 2025-2026", 12.5, RGBColor(0xB9,0xC7,0xD6), False))])

txt(s, Inches(0.85), Inches(2.15), Inches(11.6), Inches(1.2),
    [P(R("APHÉLIE", 66, WHITE, True))])
txt(s, Inches(0.85), Inches(3.25), Inches(11.4), Inches(1.3),
    [P(R("Conception d’un système de supervision centralisé", 27, ACCENT, True)),
     P(R("pour des faisceaux hertziens", 27, ACCENT, True))], line_spacing=1.05)

# bottom info bar
rect(s, 0, Inches(5.95), SW, Inches(1.55), fill=RGBColor(0x0A,0x14,0x22))
rect(s, 0, Inches(5.95), SW, Pt(2.5), fill=ACCENT)
txt(s, Inches(0.85), Inches(6.18), Inches(5), Inches(1.2),
    [P(R("PRÉSENTÉ PAR", 10, ACCENT, True)),
     P(R("Ilyesse KEBAILI", 18, WHITE, True)),
     P(R("Technicien Télécom — Alternant", 11.5, RGBColor(0xB9,0xC7,0xD6)))], space_after=2)
txt(s, Inches(6.2), Inches(6.18), Inches(3.4), Inches(1.2),
    [P(R("MAÎTRE D’APPRENTISSAGE", 10, ACCENT, True)),
     P(R("Nordine YAAQOBI", 16, WHITE, True)),
     P(R("Chef de Projet — Axians RMP", 11.5, RGBColor(0xB9,0xC7,0xD6)))], space_after=2)
txt(s, Inches(9.9), Inches(6.18), Inches(2.9), Inches(1.2),
    [P(R("ENTREPRISE & ÉCOLE", 10, ACCENT, True)),
     P(R("Axians Réseaux Mobiles Privés", 13, WHITE, True)),
     P(R("VINCI Energies · IUT de Valence", 11.5, RGBColor(0xB9,0xC7,0xD6)))], space_after=2)

# ================================================================ SLIDE 2 — SOMMAIRE
s = slide()
rect(s, 0, 0, SW, SH, fill=WHITE)
rect(s, 0, 0, Inches(4.5), SH, fill=NAVY)
gradient_navy(rect(s, 0, 0, Inches(4.5), SH, fill=NAVY), NAVY, BLUE, angle=90)
rect(s, Inches(4.5), 0, Pt(3), SH, fill=ACCENT)
txt(s, Inches(0.55), Inches(0.9), Inches(3.5), Inches(0.5),
    [P(R("PLAN DE LA", 13, ACCENT, True))])
txt(s, Inches(0.55), Inches(1.3), Inches(3.6), Inches(1.2),
    [P(R("Soutenance", 40, WHITE, True))])
txt(s, Inches(0.55), Inches(2.6), Inches(3.5), Inches(3),
    [P(R("Du contexte métier à la mise en production d’une plateforme de supervision sécurisée, puis à son évaluation financière et environnementale.", 13, RGBColor(0xC3,0xD0,0xDE)))],
    line_spacing=1.25)
plan = [
    ("Introduction", "Contexte, enjeux & problématique", False),
    ("1 — Environnement professionnel", "Company & work context — in English (5 min)", True),
    ("2 — Initialisation & cadrage", "Besoins, choix techniques, planning, EBIOS RM", False),
    ("3 — Réalisation technique", "Architecture, Zabbix/Grafana, sécurisation, IA", False),
    ("4 — Impacts du projet", "Rentabilité (ROI), environnement, opérationnel", False),
    ("5 — Bilan & perspectives", "Résultats, montée en compétences, avenir", False),
]
y = Inches(0.9)
for i,(t,d,en) in enumerate(plan):
    rect(s, Inches(5.0), y, Inches(0.5), Inches(0.5), fill=LIGHT, shape=MSO_SHAPE.OVAL)
    txt(s, Inches(5.0), y, Inches(0.5), Inches(0.5),
        [P(R(str(i) if i>0 else "•", 18, ACCENT, True))], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    txt(s, Inches(5.75), y-Inches(0.04), Inches(6.0), Inches(0.6),
        [P(R(t, 17, NAVY, True))])
    if en:
        rect(s, Inches(11.55), y+Inches(0.02), Inches(1.2), Inches(0.42), fill=ACCENT2, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        txt(s, Inches(11.55), y+Inches(0.02), Inches(1.2), Inches(0.42),
            [P(R("EN · 5 min", 9.5, NAVY, True))], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    txt(s, Inches(5.75), y+Inches(0.4), Inches(7.0), Inches(0.4),
        [P(R(d, 12, GREY, False))])
    y += Inches(0.95)
# annexes note
card(s, Inches(5.0), y+Inches(0.05), Inches(7.75), Inches(0.7), fill=NAVY)
rect(s, Inches(5.0), y+Inches(0.05), Inches(0.1), Inches(0.7), fill=ACCENT)
txt(s, Inches(5.3), y+Inches(0.05), Inches(7.4), Inches(0.7),
    [P(R("+ ANNEXES  ", 12, ACCENT, True),
       R("Questions anticipées · hypothèses ROI & CO2 · justification des choix (slides de réserve)", 11, RGBColor(0xC3,0xD0,0xDE), False))],
    anchor=MSO_ANCHOR.MIDDLE)

# ================================================================ SLIDE 3 — INTRODUCTION / PROBLÉMATIQUE
s = slide()
rect(s, 0, 0, SW, SH, fill=WHITE)
chapter_header(s, "•", "Introduction — Contexte & problématique", kicker="Pourquoi Aphélie ?")
txt(s, Inches(0.55), Inches(1.45), Inches(6.0), Inches(0.4),
    [P(R("LE CONSTAT", 12, ACCENT, True))])
bullet(s, Inches(0.55), Inches(1.85), Inches(6.0), [
    ("Faisceaux hertziens ", ": liaisons radio point-à-point qui raccordent des sites isolés ou de secours (hôpitaux, sites industriels)."),
    ("Disponibilité critique ", ": une coupure = rupture de la chaîne de soins ou arrêt de production."),
    ("Supervision obsolète ", ": l’outil historique n’était plus fiable ni sécurisé — les pannes étaient signalées par les clients eux-mêmes."),
    ("3 ans sans outil fonctionnel ", "après deux tentatives échouées (2012 & 2017)."),
], size=14, gap=11)
# problématique card
card(s, Inches(6.95), Inches(1.7), Inches(5.85), Inches(3.05), fill=NAVY)
rect(s, Inches(6.95), Inches(1.7), Inches(0.12), Inches(3.05), fill=ACCENT)
txt(s, Inches(7.3), Inches(1.95), Inches(5.3), Inches(0.5),
    [P(R("PROBLÉMATIQUE", 13, ACCENT, True))])
txt(s, Inches(7.3), Inches(2.45), Inches(5.25), Inches(2.2),
    [P(R("Comment remplacer une solution de supervision obsolète par une plateforme ", 16.5, WHITE, False),
       R("centralisée", 16.5, ACCENT, True),
       R(", capable de fournir des informations ", 16.5, WHITE, False),
       R("fiables en temps réel", 16.5, ACCENT, True),
       R(" sur les liaisons hertziennes et ", 16.5, WHITE, False),
       R("accessible à distance", 16.5, ACCENT, True),
       R(" ?", 16.5, WHITE, False))], line_spacing=1.15)
# strip objectives
for i,(ic,t) in enumerate([("Centraliser","Une plateforme unique pour tout le parc"),
                            ("Fiabiliser","Données temps réel validées"),
                            ("Sécuriser","Architecture durcie (cybersécurité)"),
                            ("Pérenniser","Solution documentée & réutilisable")]):
    x = Inches(0.55 + i*3.12)
    card(s, x, Inches(5.25), Inches(2.95), Inches(1.45), fill=CARD, line=RGBColor(0xDD,0xE6,0xEE))
    rect(s, x, Inches(5.25), Inches(2.95), Pt(3.5), fill=ACCENT)
    txt(s, x+Inches(0.22), Inches(5.45), Inches(2.6), Inches(0.4),
        [P(R(ic, 16, NAVY, True))])
    txt(s, x+Inches(0.22), Inches(5.9), Inches(2.6), Inches(0.7),
        [P(R(t, 11.5, DARKTXT, False))], line_spacing=1.0)
footer(s, 3)

# ================================================================ SLIDE 4 — [EN] THE COMPANY
s = slide()
rect(s, 0, 0, SW, SH, fill=WHITE)
chapter_header(s, "1", "The professional environment", kicker="Chapter 1 — Company", badge="ENGLISH")
# left: group cascade
txt(s, Inches(0.55), Inches(1.45), Inches(6.0), Inches(0.4),
    [P(R("A GLOBAL GROUP, A HUMAN-SIZED AGENCY", 12, ACCENT, True))])
casc = [("VINCI","280,000 employees · €74.6 bn revenue (2025)", BLUE),
        ("VINCI Energies","Energy infrastructure & IT", BLUE),
        ("Axians","ICT brand — 15,000 employees", ACCENT),
        ("Axians RMP AURA","Alixans agency · ~38 employees", NAVY)]
y=Inches(1.9)
for i,(t,d,c) in enumerate(casc):
    w = Inches(5.7 - i*0.55)
    card(s, Inches(0.55), y, w, Inches(0.74), fill=c)
    txt(s, Inches(0.78), y+Inches(0.07), w-Inches(0.4), Inches(0.62),
        [P(R(t, 14.5, WHITE, True), R("   "+d, 10.5, RGBColor(0xD9,0xE4,0xEF), False))],
        anchor=MSO_ANCHOR.MIDDLE)
    y += Inches(0.86)
txt(s, Inches(0.55), Inches(5.5), Inches(5.9), Inches(1.4),
    [P(R("Specialties: ", 12.5, NAVY, True), R("TETRA · private 4G/5G · microwave radio links.", 12.5, DARKTXT, False)),
     P(R("Customers: ", 12.5, NAVY, True), R("mostly hospitals + continuous-production industrial sites → a very high availability requirement.", 12.5, DARKTXT, False))],
    line_spacing=1.1, space_after=6)
# right: reference projects + criticality
card(s, Inches(6.65), Inches(1.7), Inches(6.15), Inches(2.55), fill=CARD, line=RGBColor(0xDD,0xE6,0xEE))
rect(s, Inches(6.65), Inches(1.7), Inches(0.12), Inches(2.55), fill=ACCENT)
txt(s, Inches(7.0), Inches(1.88), Inches(5.6), Inches(0.4),
    [P(R("TWO REFERENCE PROJECTS", 12.5, NAVY, True))])
bullet(s, Inches(7.0), Inches(2.38), Inches(5.55), [
    ("EDF Renewables (LTE on a wind farm): ", "a private 4G network — redundant cores, outdoor base stations, onboard terminals for maintenance boats."),
    ("Sytral Mobilités (TETRA): ", "modernised radio network covering 1,000+ buses, 100 tramways and 550 field agents."),
], size=12.5, gap=11)
card(s, Inches(6.65), Inches(4.45), Inches(6.15), Inches(2.25), fill=NAVY)
rect(s, Inches(6.65), Inches(4.45), Inches(6.15), Pt(4), fill=ACCENT)
txt(s, Inches(7.0), Inches(4.62), Inches(5.6), Inches(0.4),
    [P(R("WHY APHÉLIE STARTED HERE", 12.5, ACCENT, True))])
bullet(s, Inches(7.0), Inches(5.1), Inches(5.55), [
    "A geographically scattered customer base (sites with no fibre).",
    "Critical clients: a link outage = care chain disruption or production stop.",
    "No working supervision tool for ~3 years after two failed attempts.",
], size=12, gap=8, color=WHITE, marker_col=ACCENT)
footer(s, 4)

# ================================================================ SLIDE 5 — [EN] MY ROLE & CONTEXT
s = slide()
rect(s, 0, 0, SW, SH, fill=WHITE)
chapter_header(s, "1", "My position and missions", kicker="Chapter 1 — My role", badge="ENGLISH")
# left: my role
card(s, Inches(0.55), Inches(1.6), Inches(6.0), Inches(2.6), fill=CARD, line=RGBColor(0xDD,0xE6,0xEE))
rect(s, Inches(0.55), Inches(1.6), Inches(0.12), Inches(2.6), fill=ACCENT)
txt(s, Inches(0.9), Inches(1.8), Inches(5.4), Inches(0.4),
    [P(R("MY POSITION — TELECOM TECHNICIAN (APPRENTICE)", 12, NAVY, True))])
bullet(s, Inches(0.9), Inches(2.3), Inches(5.45), [
    ("Core mission: ", "the full design and deployment of the Aphélie supervision project."),
    ("Strong autonomy: ", "architecture choices, testing in an isolated lab, documentation."),
    ("Key constraint: ", "never disrupt the customers’ live production networks."),
], size=12.5, gap=11)
# left bottom: bridge
card(s, Inches(0.55), Inches(4.4), Inches(6.0), Inches(2.3), fill=NAVY)
rect(s, Inches(0.55), Inches(4.4), Inches(6.0), Pt(4), fill=ACCENT2)
txt(s, Inches(0.9), Inches(4.57), Inches(5.4), Inches(0.4),
    [P(R("FROM ENGINEERING TO SUPERVISION", 12, ACCENT2, True))])
txt(s, Inches(0.9), Inches(5.05), Inches(5.4), Inches(1.5),
    [P(R("Mastering ", 12.5, WHITE, False),
       R("link budgets", 12.5, ACCENT2, True),
       R(" is what defines the ", 12.5, WHITE, False),
       R("real alarm thresholds", 12.5, ACCENT2, True),
       R(" to monitor. Supervision then checks whether the theory holds against the field reality.", 12.5, WHITE, False))],
    line_spacing=1.2)
# right: transverse radio engineering
card(s, Inches(6.7), Inches(1.6), Inches(6.1), Inches(5.1), fill=CARD, line=RGBColor(0xDD,0xE6,0xEE))
rect(s, Inches(6.7), Inches(1.6), Inches(6.1), Pt(4), fill=ACCENT)
txt(s, Inches(7.0), Inches(1.8), Inches(5.5), Inches(0.4),
    [P(R("TRANSVERSE MISSIONS — RADIO ENGINEERING", 12, ACCENT, True))])
bullet(s, Inches(7.0), Inches(2.3), Inches(5.5), [
    ("Analysing CCTPs", " (technical specifications) from major accounts such as RTE and EDF."),
    ("Link studies with HTZ software", " — modelling terrain, the Fresnel zone, clearance and rainfall (ITU models)."),
    ("A concrete case", " — a wind-turbine proximity study validated by calculation (a 16.3 m safety margin)."),
    ("Deliverables", " — General Study, Link Engineering File (DIL), Site Installation File (DIS)."),
], size=12.5, gap=12)
txt(s, Inches(7.0), Inches(6.05), Inches(5.5), Inches(0.5),
    [P(R("→ These studies build the field knowledge of exactly what had to be supervised.", 11, GREY, True))], line_spacing=1.0)
footer(s, 5)

# ================================================================ SLIDE 5 — CHAP 2 DIAGNOSTIC + BESOINS
s = slide()
rect(s, 0, 0, SW, SH, fill=WHITE)
chapter_header(s, "2", "Initialisation & cadrage — Diagnostic & besoins", kicker="Chapitre 2")
# timeline of failures
txt(s, Inches(0.55), Inches(1.4), Inches(6), Inches(0.4),
    [P(R("UN LOURD PASSIF TECHNIQUE", 12, ACCENT, True))])
tl = [("2012","Aphélie v1 (ESISAR)","Interface sur-mesure devenue une « boîte noire » : 1 module sur 2 HS, fausses alarmes.", AMBER),
      ("2017","WhatsUp Gold","Tentative restée au stade de l’étude des MIB — jamais exploitable.", AMBER),
      ("2025","Projet Aphélie","Refonte complète confiée à l’alternant : enfin une solution opérationnelle.", ACCENT2)]
y=Inches(1.85)
for (yr,t,d,c) in tl:
    rect(s, Inches(0.55), y, Inches(1.05), Inches(1.0), fill=c)
    txt(s, Inches(0.55), y, Inches(1.05), Inches(1.0),
        [P(R(yr, 18, WHITE, True))], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    card(s, Inches(1.72), y, Inches(4.55), Inches(1.0), fill=CARD, line=RGBColor(0xDD,0xE6,0xEE))
    txt(s, Inches(1.92), y+Inches(0.1), Inches(4.2), Inches(0.4), [P(R(t, 13.5, NAVY, True))])
    txt(s, Inches(1.92), y+Inches(0.46), Inches(4.2), Inches(0.5), [P(R(d, 10.5, DARKTXT, False))], line_spacing=1.0)
    y += Inches(1.13)
# right: needs
card(s, Inches(6.65), Inches(1.6), Inches(6.15), Inches(5.1), fill=NAVY)
rect(s, Inches(6.65), Inches(1.6), Inches(6.15), Pt(4), fill=ACCENT)
txt(s, Inches(6.95), Inches(1.8), Inches(5.6), Inches(0.4),
    [P(R("LE BESOIN À COUVRIR (PÉRIMÈTRE INITIAL)", 12.5, ACCENT, True))])
bullet(s, Inches(6.95), Inches(2.35), Inches(5.55), [
    ("Parc : ", "30 à 50 faisceaux Ericsson MINI-LINK (IDU)."),
    ("Temps réel : ", "polling 1 min sur les niveaux Rx (réception) et Tx (émission)."),
    ("Alertes : ", "traps SNMP classées sur 5 niveaux de criticité, traduites en clair."),
    ("Cartographie : ", "vue dynamique, icônes colorées + segments de liaison."),
    ("Notification proactive : ", "e-mail automatique aux équipes d’astreinte sur seuil critique."),
    ("Haute disponibilité : ", "la supervision ne doit pas devenir elle-même un point faible."),
], size=13, gap=10, color=WHITE, marker_col=ACCENT)
footer(s, 6)

# ================================================================ SLIDE 6 — CHAP 2 CHOIX TECHNO
s = slide()
rect(s, 0, 0, SW, SH, fill=WHITE)
chapter_header(s, "2", "Étude de marché & choix technologiques", kicker="Chapitre 2")
txt(s, Inches(0.55), Inches(1.4), Inches(12), Inches(0.4),
    [P(R("Deux décisions structurantes, instruites par filtres successifs (souveraineté, modèle économique, maturité).", 13, GREY, False))])
# Transport choice
card(s, Inches(0.55), Inches(1.95), Inches(6.0), Inches(4.6), fill=CARD, line=RGBColor(0xDD,0xE6,0xEE))
rect(s, Inches(0.55), Inches(1.95), Inches(6.0), Inches(0.62), fill=BLUE)
txt(s, Inches(0.8), Inches(1.98), Inches(5.6), Inches(0.56),
    [P(R("① VECTEUR DE TRANSPORT", 13.5, WHITE, True))], anchor=MSO_ANCHOR.MIDDLE)
txt(s, Inches(0.8), Inches(2.75), Inches(5.5), Inches(0.4),
    [P(R("Cisco  vs  Sierra Wireless  vs  ", 12, GREY, False), R("Teltonika ✓", 13, ACCENT2, True))])
bullet(s, Inches(0.8), Inches(3.3), Inches(5.5), [
    ("Teltonika RUT956 + RMS", " : management 100 % Cloud, léger et intuitif."),
    ("Sécurité", " : VPN Hub déployé en quelques clics, accès même derrière NAT opérateur."),
    ("Coût", " : modèle à crédits optimisé (≠ licences Cisco onéreuses)."),
    ("Atouts", " : Dual-SIM (bascule 4G), retours d’expérience favorables dans le groupe VINCI."),
], size=12.5, gap=9)
# Supervision choice
card(s, Inches(6.8), Inches(1.95), Inches(6.0), Inches(4.6), fill=CARD, line=RGBColor(0xDD,0xE6,0xEE))
rect(s, Inches(6.8), Inches(1.95), Inches(6.0), Inches(0.62), fill=ACCENT)
txt(s, Inches(7.05), Inches(1.98), Inches(5.6), Inches(0.56),
    [P(R("② MOTEUR DE SUPERVISION", 13.5, WHITE, True))], anchor=MSO_ANCHOR.MIDDLE)
txt(s, Inches(7.05), Inches(2.75), Inches(5.5), Inches(0.4),
    [P(R("WhatsUp Gold · PRTG · Centreon  vs  ", 12, GREY, False), R("Zabbix ✓", 13, ACCENT2, True))])
bullet(s, Inches(7.05), Inches(3.3), Inches(5.5), [
    ("100 % open source", " : suppression des coûts de licence, scalabilité totale."),
    ("Moteur très souple", " : création libre d’items, déclencheurs et prétraitements."),
    ("Interopérabilité native", " : polling SNMP optimisé sur tunnels VPN."),
    ("Couplage parfait avec Grafana", " pour la restitution visuelle."),
    ("Urbanisation SI", " : déjà utilisé par la Direction Technique."),
], size=12.5, gap=8)
footer(s, 7)

# ================================================================ SLIDE 7 — CHAP 2 PLANNING & JALONS
s = slide()
rect(s, 0, 0, SW, SH, fill=WHITE)
chapter_header(s, "2", "Planning, jalons & analyse des écarts", kicker="Chapitre 2 — Pilotage")
txt(s, Inches(0.55), Inches(1.4), Inches(12.2), Inches(0.4),
    [P(R("5 jalons séquentiels (J1→J5) : chaque livrable conditionne le suivant.", 13, GREY, False))])
jalons = [("J1","Étude & cadrage MIB", ACCENT),
          ("J2","Maquettage Zabbix", ACCENT),
          ("J3","Déploiement client", ACCENT),
          ("J4","Grafana & cartographie", ACCENT),
          ("J5","Serveur & production", ACCENT)]
x=Inches(0.55); w=Inches(2.35)
for (j,t,c) in jalons:
    card(s, x, Inches(1.95), w, Inches(1.25), fill=NAVY)
    rect(s, x, Inches(1.95), w, Pt(4), fill=c)
    txt(s, x+Inches(0.18), Inches(2.1), w-Inches(0.3), Inches(0.4), [P(R(j, 20, ACCENT, True))])
    txt(s, x+Inches(0.18), Inches(2.55), w-Inches(0.3), Inches(0.6), [P(R(t, 11.5, WHITE, True))], line_spacing=0.95)
    x += Inches(2.46)
# arrows row implied by sequence
# écarts analysis
txt(s, Inches(0.55), Inches(3.5), Inches(6), Inches(0.4),
    [P(R("UN RETARD DE 4 SEMAINES… ENTIÈREMENT EXPLIQUÉ", 12, ACCENT, True))])
bullet(s, Inches(0.55), Inches(3.95), Inches(6.1), [
    ("J1 (+7 sem.) : ", "MIB Ericsson MINI-LINK non documentées, aucune expertise interne — investigation imprévisible."),
    ("J4 (+ interruptions) : ", "32 jours ouvrés détournés vers des missions DIL/DIS contractuelles (priorité client)."),
    ("Les 5 jalons livrés à 100 %", " — périmètre conforme au cadrage."),
], size=12.5, gap=10)
# right KPIs
card(s, Inches(7.0), Inches(3.7), Inches(5.8), Inches(2.95), fill=CARD, line=RGBColor(0xDD,0xE6,0xEE))
txt(s, Inches(7.25), Inches(3.85), Inches(5.3), Inches(0.4),
    [P(R("CHIFFRES CLÉS DU PILOTAGE", 12, NAVY, True))])
mini = [("142,8 j","effort projet (1 001 h)"),("81 %","du temps en entreprise"),
        ("100 %","des jalons livrés"),("0","jalon abandonné")]
for i,(v,l) in enumerate(mini):
    xx = Inches(7.25 + (i%2)*2.85); yy = Inches(4.35 + (i//2)*1.05)
    txt(s, xx, yy, Inches(2.7), Inches(0.5), [P(R(v, 24, ACCENT, True))])
    txt(s, xx, yy+Inches(0.46), Inches(2.7), Inches(0.4), [P(R(l, 11, DARKTXT, False))])
txt(s, Inches(0.55), Inches(6.35), Inches(6.2), Inches(0.6),
    [P(R("→ En retirant ces deux causes externes, la fin recalculée tombe ", 11.5, GREY, False),
       R("exactement à la date prévue", 11.5, ACCENT2, True),
       R(".", 11.5, GREY, False))])
footer(s, 8)

# ================================================================ SLIDE 8 — CHAP 2 EBIOS / CYBER
s = slide()
rect(s, 0, 0, SW, SH, fill=WHITE)
chapter_header(s, "2", "Cartographie des risques cyber — EBIOS RM", kicker="Chapitre 2 — Cybersécurité")
txt(s, Inches(0.55), Inches(1.4), Inches(12.2), Inches(0.4),
    [P(R("Méthode officielle ANSSI : modéliser le pire scénario d’attaque pour valider la défense.", 13, GREY, False))])
# scenario steps
txt(s, Inches(0.55), Inches(1.95), Inches(6), Inches(0.4),
    [P(R("SCÉNARIO REDOUTÉ — ATTAQUE PAR REBOND (G4 critique)", 12, RED, True))])
steps = [("1 · Infiltrer","Vol des accès d’un technicien au portail RMS (phishing, dark web).", RED),
         ("2 · Trouver","Via le VPN Hub, l’attaquant atteint un routeur RUT956 cible.", RED),
         ("3 · Exploiter","Rebond vers le LAN du client → propagation d’un rançongiciel.", RED)]
y=Inches(2.4)
for (t,d,c) in steps:
    rect(s, Inches(0.55), y, Inches(0.12), Inches(0.95), fill=c)
    card(s, Inches(0.75), y, Inches(5.5), Inches(0.95), fill=CARD, line=RGBColor(0xEE,0xD8,0xD8))
    txt(s, Inches(0.95), y+Inches(0.1), Inches(5.1), Inches(0.4), [P(R(t, 13, NAVY, True))])
    txt(s, Inches(0.95), y+Inches(0.46), Inches(5.1), Inches(0.45), [P(R(d, 11, DARKTXT, False))], line_spacing=0.95)
    y += Inches(1.08)
# countermeasures
card(s, Inches(6.65), Inches(1.95), Inches(6.15), Inches(2.35), fill=NAVY)
txt(s, Inches(6.95), Inches(2.12), Inches(5.6), Inches(0.4),
    [P(R("DEUX MESURES DE TRAITEMENT", 12.5, ACCENT2, True))])
bullet(s, Inches(6.95), Inches(2.6), Inches(5.6), [
    ("MFA strict sur RMS", " → bloque la compromission initiale, même mot de passe volé."),
    ("Durcissement RUT956 (NoNat / blocage des ports)", " → bloque toute latéralisation vers le LAN client."),
], size=13, gap=12, color=WHITE, marker_col=ACCENT2)
# risk transition
card(s, Inches(6.65), Inches(4.5), Inches(6.15), Inches(2.2), fill=CARD, line=RGBColor(0xDD,0xE6,0xEE))
txt(s, Inches(6.95), Inches(4.65), Inches(5.6), Inches(0.4),
    [P(R("RÉSULTAT — VRAISEMBLANCE", 12, NAVY, True))])
rect(s, Inches(6.95), Inches(5.15), Inches(2.4), Inches(1.0), fill=RED, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
txt(s, Inches(6.95), Inches(5.15), Inches(2.4), Inches(1.0),
    [P(R("V3", 30, WHITE, True)), P(R("Très vraisemblable", 10, WHITE, False))], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
txt(s, Inches(9.45), Inches(5.4), Inches(0.8), Inches(0.5), [P(R("→", 30, ACCENT, True))], align=PP_ALIGN.CENTER)
rect(s, Inches(10.25), Inches(5.15), Inches(2.4), Inches(1.0), fill=ACCENT2, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
txt(s, Inches(10.25), Inches(5.15), Inches(2.4), Inches(1.0),
    [P(R("V1", 30, WHITE, True)), P(R("Peu vraisemblable", 10, WHITE, False))], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
footer(s, 9)

# ================================================================ SLIDE 9 — CHAP 3 ARCHITECTURE
s = slide()
rect(s, 0, 0, SW, SH, fill=WHITE)
chapter_header(s, "3", "Réalisation — Architecture matérielle", kicker="Chapitre 3")
txt(s, Inches(0.55), Inches(1.4), Inches(12.2), Inches(0.4),
    [P(R("« Placer l’intelligence au plus près des équipements » — chaîne de bout en bout, du site client au serveur Axians.", 13, GREY, False))])
# flow diagram: 3 zones
zones = [
    ("SITE CLIENT", BLUE, [("Faisceau hertzien","Ericsson MINI-LINK / NEC iPASOLINK"),
                            ("Routeur RUT956","Routeur + pare-feu + VPN dans 1 boîtier")]),
    ("TRANSPORT SÉCURISÉ", ACCENT, [("Tunnel VPN chiffré","Sur le réseau IP du client (pas de SIM dédiée)"),
                                     ("Cloud RMS Teltonika","Administration distante + MFA")]),
    ("INFRASTRUCTURE AXIANS", NAVY, [("Serveur HPE DL360","Debian 13 · 32 Go RAM · RAID 1"),
                                      ("Zabbix + Grafana","Collecte SNMP + visualisation")]),
]
x=Inches(0.55)
for zi,(zt,zc,items) in enumerate(zones):
    w=Inches(3.95)
    card(s, x, Inches(2.0), w, Inches(4.0), fill=WHITE, line=RGBColor(0xDD,0xE6,0xEE))
    rect(s, x, Inches(2.0), w, Inches(0.6), fill=zc)
    txt(s, x, Inches(2.0), w, Inches(0.6), [P(R(zt, 13, WHITE, True))], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    yy=Inches(2.85)
    for (t,d) in items:
        card(s, x+Inches(0.2), yy, w-Inches(0.4), Inches(1.35), fill=CARD)
        rect(s, x+Inches(0.2), yy, Inches(0.08), Inches(1.35), fill=zc)
        txt(s, x+Inches(0.4), yy+Inches(0.15), w-Inches(0.65), Inches(0.5), [P(R(t, 13, NAVY, True))], line_spacing=0.95)
        txt(s, x+Inches(0.4), yy+Inches(0.68), w-Inches(0.65), Inches(0.6), [P(R(d, 11, DARKTXT, False))], line_spacing=0.95)
        yy += Inches(1.5)
    if zi<2:
        txt(s, x+w-Inches(0.05), Inches(3.7), Inches(0.55), Inches(0.6), [P(R("→", 26, ACCENT, True))], align=PP_ALIGN.CENTER)
    x += Inches(4.25)
txt(s, Inches(0.55), Inches(6.2), Inches(12), Inches(0.5),
    [P(R("Déploiement « Plug & Play » : ", 12.5, NAVY, True),
       R("préconfiguration en atelier, mise en service à distance par téléphone, vagues de 10 routeurs — sans déplacement Axians.", 12.5, DARKTXT, False))])
footer(s, 10)

# ================================================================ SLIDE 10 — CHAP 3 ZABBIX & MIB
s = slide()
rect(s, 0, 0, SW, SH, fill=WHITE)
chapter_header(s, "3", "Intégration Zabbix & traduction des dictionnaires MIB", kicker="Chapitre 3 — Le cœur technique")
txt(s, Inches(0.55), Inches(1.4), Inches(6), Inches(0.4),
    [P(R("LE DÉFI : DÉCODER L’ERICSSON MINI-LINK", 12, ACCENT, True))])
bullet(s, Inches(0.55), Inches(1.85), Inches(6.1), [
    ("Centaines d’OID à trier", " : capture des trames SNMP directement sur les équipements en production."),
    ("Valeurs piégeuses", " : Rx en dixièmes de dBm (-450 = -45,0 dBm) → prétraitement par multiplication dans le Template."),
    ("Traps propriétaires", " : chaîne brute décodée champ par champ via expressions régulières (regex apprises en contexte)."),
    ("Horodatage hexadécimal", " + codes de cause traduits par table de correspondance."),
], size=12.5, gap=10)
# right: outcomes
card(s, Inches(7.0), Inches(1.7), Inches(5.8), Inches(2.45), fill=NAVY)
rect(s, Inches(7.0), Inches(1.7), Inches(5.8), Pt(4), fill=ACCENT)
txt(s, Inches(7.3), Inches(1.88), Inches(5.3), Inches(0.4),
    [P(R("CE QUI A ÉTÉ PRODUIT", 12.5, ACCENT, True))])
bullet(s, Inches(7.3), Inches(2.35), Inches(5.3), [
    "Référentiel OID validé — un actif technique inédit pour l’entreprise.",
    "Templates Zabbix génériques, réutilisables sur tout le parc MINI-LINK.",
    "4 niveaux de sévérité + trigger Rx < −75 dBm (seuil de dégradation).",
], size=12.5, gap=9, color=WHITE, marker_col=ACCENT)
# bottom: example value transform
card(s, Inches(7.0), Inches(4.35), Inches(5.8), Inches(2.25), fill=CARD, line=RGBColor(0xDD,0xE6,0xEE))
txt(s, Inches(7.3), Inches(4.5), Inches(5.3), Inches(0.4),
    [P(R("EXEMPLE DE CHAÎNE DE TRAITEMENT", 12, NAVY, True))])
chain = [("-450","valeur brute SNMP", GREY),
         ("× 0,1","prétraitement", ACCENT),
         ("-45,0 dBm","stocké & affiché", ACCENT2)]
xx=Inches(7.3)
for i,(v,l,c) in enumerate(chain):
    rect(s, xx, Inches(5.05), Inches(1.55), Inches(1.1), fill=WHITE, line=c, line_w=Pt(1.5), shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    txt(s, xx, Inches(5.18), Inches(1.55), Inches(0.5), [P(R(v, 17, c, True))], align=PP_ALIGN.CENTER)
    txt(s, xx, Inches(5.68), Inches(1.55), Inches(0.4), [P(R(l, 9.5, DARKTXT, False))], align=PP_ALIGN.CENTER)
    if i<2: txt(s, xx+Inches(1.5), Inches(5.2), Inches(0.45), Inches(0.6), [P(R("→", 22, ACCENT, True))], align=PP_ALIGN.CENTER)
    xx += Inches(1.95)
footer(s, 11)

# ================================================================ SLIDE 11 — CHAP 3 SECU RESEAU
s = slide()
rect(s, 0, 0, SW, SH, fill=WHITE)
chapter_header(s, "3", "Durcissement des flux & identification des faisceaux", kicker="Chapitre 3 — Sécurisation")
# left philosophy
txt(s, Inches(0.55), Inches(1.45), Inches(6), Inches(0.4),
    [P(R("POSTURE FERMÉE PAR DÉFAUT", 12, ACCENT, True))])
bullet(s, Inches(0.55), Inches(1.9), Inches(6.1), [
    ("« Tout ce qui n’est pas autorisé est rejeté »", " — pare-feu en REJECT par défaut."),
    ("Interfaces d’admin (SSH/HTTP/CLI) désactivées côté WAN", " : admin uniquement via tunnel RMS chiffré + MFA."),
    ("Zonage strict", " : LAN / WAN / OpenVPN / RMS, aucun transit non prévu."),
    ("Seul flux entrant autorisé", " : le tunnel de supervision."),
], size=12.5, gap=10)
# right NoNat
card(s, Inches(7.0), Inches(1.6), Inches(5.8), Inches(5.1), fill=NAVY)
rect(s, Inches(7.0), Inches(1.6), Inches(5.8), Pt(4), fill=ACCENT2)
txt(s, Inches(7.3), Inches(1.78), Inches(5.3), Inches(0.4),
    [P(R("LES RÈGLES NoNat — IDENTIFIER CHAQUE FAISCEAU", 12, ACCENT2, True))])
txt(s, Inches(7.3), Inches(2.3), Inches(5.3), Inches(1.0),
    [P(R("Sans NoNat, le masquerade remplace l’IP de chaque IDU par celle du routeur : tous les faisceaux d’un site apparaissent identiques dans Zabbix.", 12, WHITE, False))], line_spacing=1.1)
# before/after
card(s, Inches(7.3), Inches(3.5), Inches(5.2), Inches(1.05), fill=RGBColor(0x2A,0x1C,0x1C))
txt(s, Inches(7.5), Inches(3.6), Inches(4.9), Inches(0.4), [P(R("✗ Sans NoNat", 12, RED, True))])
txt(s, Inches(7.5), Inches(3.98), Inches(4.9), Inches(0.5), [P(R("3 faisceaux → 1 seule IP source → supervision individuelle impossible.", 10.5, RGBColor(0xE8,0xC8,0xC8), False))], line_spacing=0.95)
card(s, Inches(7.3), Inches(4.7), Inches(5.2), Inches(1.05), fill=RGBColor(0x12,0x2A,0x22))
txt(s, Inches(7.5), Inches(4.8), Inches(4.9), Inches(0.4), [P(R("✓ Avec NoNat (ciblé SNMP)", 12, ACCENT2, True))])
txt(s, Inches(7.5), Inches(5.18), Inches(4.9), Inches(0.5), [P(R("IP source préservée → chaque faisceau identifié précisément en alarme.", 10.5, RGBColor(0xC8,0xE8,0xD8), False))], line_spacing=0.95)
txt(s, Inches(7.3), Inches(5.95), Inches(5.2), Inches(0.6),
    [P(R("Désactivation limitée au seul trafic de supervision → cloisonnement préservé.", 10.5, RGBColor(0xB9,0xC7,0xD6), True))], line_spacing=1.0)
txt(s, Inches(0.55), Inches(5.6), Inches(6.1), Inches(1.1),
    [P(R("Traduction opérationnelle directe des mesures EBIOS RM", 12.5, NAVY, True)),
     P(R("→ la surface d’attaque exposée chez le client est réduite au strict minimum.", 12, DARKTXT, False))], line_spacing=1.1, space_after=4)
footer(s, 12)

# ================================================================ SLIDE 12 — CHAP 3 RESTITUTION VISUELLE
s = slide()
rect(s, 0, 0, SW, SH, fill=WHITE)
chapter_header(s, "3", "Restitution visuelle & cartographie dynamique", kicker="Chapitre 3 — L’interface")
txt(s, Inches(0.55), Inches(1.4), Inches(12.2), Inches(0.4),
    [P(R("Deux niveaux de lecture pour transformer un flux de données brut en information immédiate.", 13, GREY, False))])
# two big panels
card(s, Inches(0.55), Inches(2.0), Inches(6.0), Inches(4.6), fill=CARD, line=RGBColor(0xDD,0xE6,0xEE))
rect(s, Inches(0.55), Inches(2.0), Inches(6.0), Inches(0.6), fill=BLUE)
txt(s, Inches(0.8), Inches(2.0), Inches(5.6), Inches(0.6), [P(R("NIVEAU OPÉRATIONNEL — ZABBIX", 13, WHITE, True))], anchor=MSO_ANCHOR.MIDDLE)
# hex mock
hexcols=[ACCENT2,ACCENT2,RED,ACCENT2,AMBER,ACCENT2]
for i,c in enumerate(hexcols):
    hx=Inches(0.95+ (i%3)*1.85); hy=Inches(2.85+(i//3)*1.0)
    sp=rect(s, hx, hy, Inches(1.5), Inches(0.85), fill=c, shape=MSO_SHAPE.HEXAGON)
    txt(s, hx, hy, Inches(1.5), Inches(0.85), [P(R("-4%d dBm"%(2+i), 11, WHITE, True))], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
bullet(s, Inches(0.85), Inches(4.95), Inches(5.5), [
    "Panneaux hexagonaux Rx en temps réel (rouge dès qu’une alarme est active).",
    "Tables Rx/Tx par hôte + panneaux de disponibilité (SLO) par client.",
    "E-mail automatique d’alerte ET de résolution, acquittement nominatif.",
], size=11.5, gap=7)
# Grafana
card(s, Inches(6.8), Inches(2.0), Inches(6.0), Inches(4.6), fill=CARD, line=RGBColor(0xDD,0xE6,0xEE))
rect(s, Inches(6.8), Inches(2.0), Inches(6.0), Inches(0.6), fill=ACCENT)
txt(s, Inches(7.05), Inches(2.0), Inches(5.6), Inches(0.6), [P(R("NIVEAU CARTOGRAPHIQUE — GRAFANA", 13, WHITE, True))], anchor=MSO_ANCHOR.MIDDLE)
# map mock
rect(s, Inches(7.1), Inches(2.85), Inches(5.4), Inches(1.95), fill=RGBColor(0xE3,0xEC,0xF2), shape=MSO_SHAPE.ROUNDED_RECTANGLE)
pts=[(7.6,3.2,ACCENT2),(9.3,3.0,ACCENT2),(11.4,3.4,RED),(8.2,4.2,ACCENT2),(10.6,4.3,ACCENT2),(9.6,3.8,AMBER)]
import itertools
for (a,b,c) in pts:
    rect(s, Inches(a), Inches(b), Inches(0.18), Inches(0.18), fill=c, shape=MSO_SHAPE.OVAL)
# links
def link(x1,y1,x2,y2,c):
    ln=s.shapes.add_connector(2, Inches(x1+0.09), Inches(y1+0.09), Inches(x2+0.09), Inches(y2+0.09))
    ln.line.color.rgb=c; ln.line.width=Pt(2)
link(7.6,3.2,9.3,3.0,ACCENT2); link(9.3,3.0,9.6,3.8,AMBER); link(9.6,3.8,11.4,3.4,RED)
link(8.2,4.2,9.6,3.8,ACCENT2); link(10.6,4.3,11.4,3.4,RED)
bullet(s, Inches(7.05), Inches(4.95), Inches(5.5), [
    ("Script grafana_topology.py", " développé sur-mesure : lit MySQL Zabbix → GeoJSON (plugin Geomap)."),
    "Liaisons tracées site-à-site ; trait rouge = faisceau en alarme.",
    "Pop-up : Rx/Tx, alarmes actives, lien direct vers l’équipement.",
], size=11.5, gap=7)
footer(s, 13)

# ================================================================ SLIDE 13 — CHAP 3 IA
s = slide()
rect(s, 0, 0, SW, SH, fill=WHITE)
chapter_header(s, "3", "L’IA locale comme accélérateur technique", kicker="Chapitre 3 — Au-delà du périmètre")
txt(s, Inches(0.55), Inches(1.4), Inches(12.2), Inches(0.45),
    [P(R("Deux prototypes exploratoires, exécutés ", 13, GREY, False),
       R("100 % en local (Ollama)", 13, ACCENT, True),
       R(" — confidentialité des données, aucun cloud tiers.", 13, GREY, False))])
# proto 1
card(s, Inches(0.55), Inches(2.05), Inches(6.0), Inches(3.7), fill=CARD, line=RGBColor(0xDD,0xE6,0xEE))
rect(s, Inches(0.55), Inches(2.05), Inches(0.12), Inches(3.7), fill=ACCENT)
txt(s, Inches(0.9), Inches(2.25), Inches(5.4), Inches(0.5), [P(R("① Corrélation d’alarmes", 16, NAVY, True))])
txt(s, Inches(0.9), Inches(2.75), Inches(5.4), Inches(0.4), [P(R("Ollama + DBSCAN (clustering non supervisé)", 12, ACCENT, True))])
bullet(s, Inches(0.9), Inches(3.25), Inches(5.4), [
    "Rapproche traps SNMP, niveaux Rx, logs Zabbix dans une même fenêtre temporelle.",
    "Objectif : faire émerger une cause racine à partir de signaux dispersés.",
    "Aide aussi à l’analyse des MIB et à la construction des regex.",
], size=12, gap=9)
# proto 2
card(s, Inches(6.8), Inches(2.05), Inches(6.0), Inches(3.7), fill=NAVY)
rect(s, Inches(6.8), Inches(2.05), Inches(0.12), Inches(3.7), fill=ACCENT2)
txt(s, Inches(7.15), Inches(2.25), Inches(5.4), Inches(0.5), [P(R("② Assistant conversationnel", 16, WHITE, True))])
txt(s, Inches(7.15), Inches(2.75), Inches(5.4), Inches(0.4), [P(R("Modèle Mistral hébergé localement", 12, ACCENT2, True))])
bullet(s, Inches(7.15), Inches(3.25), Inches(5.4), [
    "Consultation de la supervision en langage naturel (utile en astreinte).",
    "Intégré au portail centralisé (SSO Microsoft Entra ID + MFA).",
    "Choix Mistral : français, taille adaptée au serveur HPE, modèle ouvert.",
], size=12, gap=9, color=WHITE, marker_col=ACCENT2)
# disclaimer
card(s, Inches(0.55), Inches(5.95), Inches(12.25), Inches(0.75), fill=RGBColor(0xFD,0xF3,0xE0))
rect(s, Inches(0.55), Inches(5.95), Inches(0.1), Inches(0.75), fill=AMBER)
txt(s, Inches(0.85), Inches(5.95), Inches(11.8), Inches(0.75),
    [P(R("Posture honnête : ", 12.5, AMBER, True),
       R("ce sont des pistes d’expérimentation validant un intérêt — pas encore des fonctions industrialisées en production.", 12.5, DARKTXT, False))],
    anchor=MSO_ANCHOR.MIDDLE)
footer(s, 14)

# ================================================================ SLIDE 14 — CHAP 4 ROI
s = slide()
rect(s, 0, 0, SW, SH, fill=WHITE)
chapter_header(s, "4", "Impacts financiers — Rentabilité (ROI / TCO)", kicker="Chapitre 4")
# KPIs row
kpi(s, Inches(0.55), Inches(1.55), Inches(2.95), "310 %", "ROI cumulé à 5 ans", vcol=ACCENT)
kpi(s, Inches(3.65), Inches(1.55), Inches(2.95), "14,6 mois", "Point mort après mise en prod.", vcol=ACCENT2)
kpi(s, Inches(6.75), Inches(1.55), Inches(2.95), "176 284 €", "Gain net absolu sur 5 ans", vcol=ACCENT)
kpi(s, Inches(9.85), Inches(1.55), Inches(2.95), "56 646 €", "Investissement total (TCO)", vcol=NAVY)
# left: TCO breakdown
txt(s, Inches(0.55), Inches(3.35), Inches(6), Inches(0.4),
    [P(R("STRUCTURE DU COÛT (TCO)", 12, ACCENT, True))])
rows=[("CAPEX humain (savoir-faire interne)","46 440 €", ACCENT, 0.82),
      ("CAPEX matériel (serveur, routeurs…)","10 206 €", BLUE, 0.18),
      ("OPEX récurrent","2 694 €/an", GREY, 0.05)]
yy=Inches(3.8)
for (t,v,c,frac) in rows[:2]:
    txt(s, Inches(0.55), yy, Inches(4.4), Inches(0.35), [P(R(t, 11.5, DARKTXT, False))])
    txt(s, Inches(5.0), yy, Inches(1.4), Inches(0.35), [P(R(v, 12, NAVY, True))], align=PP_ALIGN.RIGHT)
    rect(s, Inches(0.55), yy+Inches(0.32), Inches(5.85), Inches(0.18), fill=RGBColor(0xE3,0xEA,0xF1))
    rect(s, Inches(0.55), yy+Inches(0.32), Emu(int(Inches(5.85)*frac)), Inches(0.18), fill=c)
    yy += Inches(0.72)
txt(s, Inches(0.55), yy+Inches(0.05), Inches(6), Inches(0.6),
    [P(R("82 % du TCO = coût humain", 12.5, ACCENT, True),
       R(" → un actif immatériel (templates, doc, savoir-faire) et non un coût perdu.", 12, DARKTXT, False))], line_spacing=1.05)
# right: model robustness
card(s, Inches(6.9), Inches(3.35), Inches(5.9), Inches(3.3), fill=NAVY)
txt(s, Inches(7.2), Inches(3.5), Inches(5.3), Inches(0.4),
    [P(R("UN MODÈLE ÉCONOMIQUE ROBUSTE", 12.5, ACCENT, True))])
bullet(s, Inches(7.2), Inches(3.98), Inches(5.4), [
    ("Revenus à deux têtes", " : abonnements (16 clients × 2 000 €/an) + économies GTI (déplacements évités)."),
    ("Scénario conservateur", " : même à 1 500 €/client, ROI 169 % à 5 ans."),
    ("Effet de levier", " : OPEX fixe → un serveur absorbe 50+ clients sans surcoût."),
    ("Open source", " : 0 € de licence (vs 5 000–20 000 €/an en propriétaire)."),
], size=12, gap=8, color=WHITE, marker_col=ACCENT)
footer(s, 15)

# ================================================================ SLIDE 15 — CHAP 4 ENV + OPERATIONNEL
s = slide()
rect(s, 0, 0, SW, SH, fill=WHITE)
chapter_header(s, "4", "Impacts environnemental & opérationnel", kicker="Chapitre 4")
# left env
txt(s, Inches(0.55), Inches(1.45), Inches(6), Inches(0.4),
    [P(R("LE « PARADOXE DU SERVEUR ALLUMÉ » LEVÉ PAR LE CALCUL", 12, ACCENT2, True))])
# comparison bars
card(s, Inches(0.55), Inches(1.95), Inches(6.0), Inches(2.55), fill=CARD, line=RGBColor(0xDD,0xE6,0xEE))
txt(s, Inches(0.8), Inches(2.1), Inches(5.5), Inches(0.4), [P(R("Bilan CO₂ annuel (scénario de base 100 km)", 12, NAVY, True))])
# évité bar
txt(s, Inches(0.8), Inches(2.6), Inches(3), Inches(0.3), [P(R("Déplacements évités", 11, DARKTXT))])
rect(s, Inches(0.8), Inches(2.9), Inches(5.4), Inches(0.4), fill=ACCENT2)
txt(s, Inches(0.95), Inches(2.9), Inches(5.1), Inches(0.4), [P(R("576 kg CO₂", 12, WHITE, True))], anchor=MSO_ANCHOR.MIDDLE)
txt(s, Inches(0.8), Inches(3.45), Inches(3), Inches(0.3), [P(R("Serveur HPE", 11, DARKTXT))])
rect(s, Inches(0.8), Inches(3.75), Inches(0.92), Inches(0.4), fill=GREY)
txt(s, Inches(1.8), Inches(3.75), Inches(2), Inches(0.4), [P(R("98 kg CO₂", 11, GREY, True))], anchor=MSO_ANCHOR.MIDDLE)
txt(s, Inches(0.8), Inches(4.18), Inches(5.5), Inches(0.3),
    [P(R("→ Le serveur émet 5,9× moins que les trajets qu’il remplace.", 11, NAVY, True))])
# env KPIs
kpi(s, Inches(0.55), Inches(4.75), Inches(2.9), "478 kg", "CO₂ net économisés / an", vcol=ACCENT2, h=Inches(1.45))
kpi(s, Inches(3.65), Inches(4.75), Inches(2.9), "≈ 2,4 t", "CO₂ évités cumulés sur 5 ans", vcol=ACCENT2, h=Inches(1.45))
# right operational
card(s, Inches(6.9), Inches(1.95), Inches(5.9), Inches(4.75), fill=NAVY)
rect(s, Inches(6.9), Inches(1.95), Inches(5.9), Pt(4), fill=ACCENT)
txt(s, Inches(7.2), Inches(2.15), Inches(5.3), Inches(0.4),
    [P(R("UN 3ᵉ BÉNÉFICE : LA VALIDATION TERRAIN", 12.5, ACCENT, True))])
bullet(s, Inches(7.2), Inches(2.65), Inches(5.4), [
    ("Confronter Rx mesuré vs Rx théorique (DIL)", " en temps réel, sans déplacement."),
    ("Détection immédiate des non-conformités", " : défaut d’alignement, obstacle, perte câble…"),
    ("De plusieurs jours à quelques minutes", " pour identifier une anomalie d’installation."),
    ("Passage du réactif au proactif", " : Zabbix détecte les signes précurseurs avant la panne."),
    ("Gage de crédibilité", " vis-à-vis des clients : preuve chiffrée de conformité."),
], size=12.5, gap=11, color=WHITE, marker_col=ACCENT)
footer(s, 16)

# ================================================================ SLIDE 16 — CHAP 5 BILAN
s = slide()
rect(s, 0, 0, SW, SH, fill=WHITE)
chapter_header(s, "5", "Bilan & perspectives", kicker="Chapitre 5")
txt(s, Inches(0.55), Inches(1.4), Inches(12.2), Inches(0.4),
    [P(R("La problématique trouve une réponse opérationnelle complète : les 5 jalons sont atteints.", 13.5, NAVY, True))])
# answer to problematique mapping
ans=[("Centralisée","Plateforme Zabbix-Grafana unique sur serveur HPE."),
     ("Fiable","MIB Ericsson investiguées + validation contradictoire des signaux."),
     ("Temps réel","Polling 1 min + traitement des traps spontanés."),
     ("À distance","Tunnels VPN + consultation web des tableaux de bord.")]
x=Inches(0.55)
for (t,d) in ans:
    card(s, x, Inches(1.95), Inches(2.97), Inches(1.5), fill=CARD, line=RGBColor(0xDD,0xE6,0xEE))
    rect(s, x, Inches(1.95), Inches(2.97), Pt(3.5), fill=ACCENT2)
    txt(s, x+Inches(0.2), Inches(2.1), Inches(2.6), Inches(0.4), [P(R("✓ "+t, 14, NAVY, True))])
    txt(s, x+Inches(0.2), Inches(2.55), Inches(2.6), Inches(0.85), [P(R(d, 11, DARKTXT, False))], line_spacing=1.0)
    x += Inches(3.1)
# perspectives + competences
card(s, Inches(0.55), Inches(3.65), Inches(6.0), Inches(3.05), fill=NAVY)
txt(s, Inches(0.85), Inches(3.82), Inches(5.4), Inches(0.4), [P(R("PERSPECTIVES", 12.5, ACCENT, True))])
bullet(s, Inches(0.85), Inches(4.3), Inches(5.4), [
    ("Élargir le parc supervisé", " : axe n°1 — chaque nouveau client ≈ flux net intégral (OPEX fixe)."),
    ("Industrialiser l’IA locale", " : corrélation d’alarmes & assistant conversationnel."),
    ("Pérennité", " : documentation complète, templates réutilisables, reprise par le service RMP."),
], size=12, gap=10, color=WHITE, marker_col=ACCENT)
card(s, Inches(6.8), Inches(3.65), Inches(6.0), Inches(3.05), fill=CARD, line=RGBColor(0xDD,0xE6,0xEE))
txt(s, Inches(7.1), Inches(3.82), Inches(5.4), Inches(0.4), [P(R("MA MONTÉE EN COMPÉTENCES", 12.5, ACCENT, True))])
bullet(s, Inches(7.1), Inches(4.3), Inches(5.4), [
    ("Technique", " : SNMP, MIB, liaisons hertziennes, Zabbix/Grafana de A à Z, VPN/NoNat."),
    ("Méthode", " : WBS, RACI, EBIOS RM, SWOT, suivi des écarts."),
    ("Avenir", " : poursuite en école d’ingénieur en alternance, chez Axians RMP."),
], size=12, gap=10)
footer(s, 17)

# ================================================================ SLIDE 17 — CONCLUSION / MERCI
s = slide()
bg = rect(s, 0, 0, SW, SH, fill=NAVY)
gradient_navy(bg, NAVY, BLUE, angle=60)
rect(s, 0, 0, Inches(0.22), SH, fill=ACCENT)
for (dx, dy) in [(11.7,5.6),(12.4,6.2),(11.2,6.4),(12.7,5.4)]:
    rect(s, Inches(dx), Inches(dy), Inches(0.13), Inches(0.13), fill=ACCENT, shape=MSO_SHAPE.OVAL)
txt(s, Inches(0.9), Inches(1.4), Inches(11), Inches(0.5),
    [P(R("EN CONCLUSION", 14, ACCENT, True))])
txt(s, Inches(0.9), Inches(1.95), Inches(11.6), Inches(2.0),
    [P(R("D’une supervision obsolète à une plateforme", 30, WHITE, True)),
     P(R("centralisée, fiable, sécurisée et rentable.", 30, WHITE, True))], line_spacing=1.1)
txt(s, Inches(0.9), Inches(3.7), Inches(11.4), Inches(1.0),
    [P(R("Un projet mené de bout en bout — du cadrage à la mise en production — qui constitue désormais un actif technique durable pour Axians RMP.", 15, RGBColor(0xC3,0xD0,0xDE), False))], line_spacing=1.2)
# mini recap chips
chips=[("310 %","ROI à 5 ans"),("16","clients supervisés"),("478 kg","CO₂ évités/an"),("5/5","jalons livrés")]
x=Inches(0.9)
for (v,l) in chips:
    card(s, x, Inches(4.75), Inches(2.7), Inches(1.0), fill=RGBColor(0x12,0x22,0x36))
    rect(s, x, Inches(4.75), Inches(2.7), Pt(3), fill=ACCENT)
    txt(s, x+Inches(0.2), Inches(4.9), Inches(2.4), Inches(0.5), [P(R(v, 22, ACCENT, True))])
    txt(s, x+Inches(0.2), Inches(5.38), Inches(2.4), Inches(0.35), [P(R(l, 11, WHITE, False))])
    x += Inches(2.92)
rect(s, 0, Inches(6.25), SW, Inches(1.25), fill=RGBColor(0x0A,0x14,0x22))
rect(s, 0, Inches(6.25), SW, Pt(2.5), fill=ACCENT)
txt(s, Inches(0.9), Inches(6.5), Inches(8), Inches(0.8),
    [P(R("Merci de votre attention.", 22, WHITE, True), R("   Je suis à votre disposition pour vos questions.", 14, RGBColor(0xB9,0xC7,0xD6), False))],
    anchor=MSO_ANCHOR.MIDDLE)
txt(s, Inches(9.5), Inches(6.5), Inches(3.3), Inches(0.8),
    [P(R("Ilyesse KEBAILI", 14, ACCENT, True)), P(R("BUT R&T · Cybersécurité", 11, RGBColor(0xB9,0xC7,0xD6)))],
    align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE, space_after=2)

# ================================================================ ANNEX HELPERS
def set_cell(cell, runs, fill, size=11, anchor=MSO_ANCHOR.MIDDLE):
    cell.fill.solid(); cell.fill.fore_color.rgb = fill
    cell.margin_left = Inches(0.08); cell.margin_right = Inches(0.08)
    cell.margin_top = Inches(0.03); cell.margin_bottom = Inches(0.03)
    cell.vertical_anchor = anchor
    tf = cell.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]; p.line_spacing = 0.98
    for (t, col, bold) in runs:
        r = p.add_run(); r.text = t; r.font.size = Pt(size)
        r.font.color.rgb = col; r.font.bold = bold; r.font.name = FONT

def ntable(s, x, y, col_w, header, rows, hfill=BLUE, row_h=Inches(0.78),
           head_h=Inches(0.42), size=11):
    nrows = len(rows)+1; ncols = len(header)
    total_w = sum(col_w, Emu(0))
    gt = s.shapes.add_table(nrows, ncols, x, y, total_w, head_h+row_h*len(rows)).table
    gt.first_row = False; gt.horz_banding = False
    for i,cw in enumerate(col_w): gt.columns[i].width = cw
    gt.rows[0].height = head_h
    for j,h in enumerate(header):
        set_cell(gt.cell(0,j), [(h, WHITE, True)], hfill, size=size)
    for i,row in enumerate(rows, start=1):
        gt.rows[i].height = row_h
        for j,cell_runs in enumerate(row):
            bg = WHITE if i%2==1 else CARD
            set_cell(gt.cell(i,j), cell_runs, bg, size=size)
    return gt

def annex_header(s, title, kicker, badge="ANNEXE", badge_col=ACCENT):
    band = rect(s, 0, 0, SW, Inches(1.18), fill=NAVY)
    gradient_navy(band, NAVY, BLUE, angle=0)
    rect(s, Inches(0.5), Inches(0.27), Inches(0.62), Inches(0.62), fill=badge_col)
    txt(s, Inches(0.5), Inches(0.27), Inches(0.62), Inches(0.62),
        [P(R("A", 26, NAVY, True))], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    txt(s, Inches(1.32), Inches(0.20), Inches(9.6), Inches(0.85),
        [P(R(kicker.upper(), 10.5, ACCENT, True)), P(R(title, 21, WHITE, True))],
        anchor=MSO_ANCHOR.MIDDLE)
    bw = Inches(1.7)
    rect(s, SW-bw-Inches(0.5), Inches(0.34), bw, Inches(0.5), fill=badge_col, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    txt(s, SW-bw-Inches(0.5), Inches(0.34), bw, Inches(0.5),
        [P(R(badge, 12, NAVY, True))], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    rect(s, 0, Inches(1.18), SW, Pt(3), fill=ACCENT)

def qa_card(s, x, y, w, h, q, a, qcol=ACCENT):
    card(s, x, y, w, h, fill=CARD, line=RGBColor(0xDD,0xE6,0xEE))
    rect(s, x, y, Inches(0.1), h, fill=qcol)
    txt(s, x+Inches(0.28), y+Inches(0.12), w-Inches(0.45), Inches(0.55),
        [P(R(q, 12.5, NAVY, True))], line_spacing=0.98)
    txt(s, x+Inches(0.28), y+Inches(0.66), w-Inches(0.45), h-Inches(0.75),
        [P(R("→ ", 12, qcol, True), R(a, 11.5, DARKTXT, False))], line_spacing=1.0)

# ================================================================ SLIDE 18 — ANNEX DIVIDER
s = slide()
bg = rect(s, 0, 0, SW, SH, fill=NAVY); gradient_navy(bg, NAVY, BLUE, angle=60)
rect(s, 0, 0, Inches(0.22), SH, fill=ACCENT)
for (dx, dy) in [(11.7,0.9),(12.4,1.5),(11.2,1.7),(12.7,0.7)]:
    rect(s, Inches(dx), Inches(dy), Inches(0.13), Inches(0.13), fill=ACCENT, shape=MSO_SHAPE.OVAL)
rect(s, Inches(0.9), Inches(2.5), Inches(3.4), Pt(2.5), fill=ACCENT)
txt(s, Inches(0.9), Inches(2.7), Inches(11), Inches(1.4),
    [P(R("ANNEXES", 58, WHITE, True))])
txt(s, Inches(0.9), Inches(3.95), Inches(11.4), Inches(0.6),
    [P(R("Questions anticipées · Hypothèses détaillées · Justification des choix", 19, ACCENT, True))])
card(s, Inches(0.9), Inches(5.0), Inches(11.5), Inches(1.0), fill=RGBColor(0x12,0x22,0x36))
rect(s, Inches(0.9), Inches(5.0), Inches(0.1), Inches(1.0), fill=ACCENT)
txt(s, Inches(1.2), Inches(5.0), Inches(11), Inches(1.0),
    [P(R("Diapositives de réserve — mobilisables pendant la phase de questions/réponses pour appuyer chaque chiffre et chaque choix par une preuve.", 13.5, RGBColor(0xC3,0xD0,0xDE), False))],
    anchor=MSO_ANCHOR.MIDDLE)

# ================================================================ SLIDE 19 — LE BESOIN EN DÉTAIL
s = slide()
rect(s, 0, 0, SW, SH, fill=WHITE)
annex_header(s, "Le besoin — pourquoi, pour qui, comment, où", "Annexe 1 — Cadrage du besoin")
quad = [
    ("POURQUOI", ACCENT, [
        "Supervision obsolète : 1 module sur 2 HS, fausses alarmes.",
        "~3 ans sans outil → pannes signalées par les clients.",
        "Parc critique : hôpitaux, industrie continue, OIV.",
    ]),
    ("POUR QUI", BLUE, [
        "Techniciens / NOC : diagnostic immédiat, moins de déplacements.",
        "Chef de projet & responsable d’affaires : pilotage.",
        "Clients & agence : continuité de service + rentabilité.",
    ]),
    ("COMMENT", ACCENT2, [
        "RUT956 → tunnel VPN / RMS → Zabbix → Grafana → serveur HPE.",
        "SNMP polling 1 min + traps ; templates génériques.",
        "Sécurité EBIOS RM : MFA + durcissement / NoNat.",
    ]),
    ("OÙ / QUOI", NAVY, [
        "Sites dispersés en France (zones sans fibre).",
        "Périmètre : 16 clients, 30-50 faisceaux Ericsson MINI-LINK (IDU).",
        "Hors périmètre : IA locale, portail Entra ID (pistes).",
    ]),
]
for k,(t,c,items) in enumerate(quad):
    x = Inches(0.55 + (k%2)*6.25); y = Inches(1.5 + (k//2)*2.65)
    card(s, x, y, Inches(6.0), Inches(2.45), fill=CARD, line=RGBColor(0xDD,0xE6,0xEE))
    rect(s, x, y, Inches(6.0), Inches(0.5), fill=c)
    txt(s, x+Inches(0.2), y, Inches(5.6), Inches(0.5), [P(R(t, 14, WHITE, True))], anchor=MSO_ANCHOR.MIDDLE)
    bullet(s, x+Inches(0.25), y+Inches(0.68), Inches(5.5), items, size=12, gap=9, marker_col=c)
footer(s, 20)

# ================================================================ SLIDE 20 — HYPOTHÈSES ROI
s = slide()
rect(s, 0, 0, SW, SH, fill=WHITE)
annex_header(s, "Hypothèses du ROI — chaque chiffre justifié", "Annexe 2 — Modèle financier")
txt(s, Inches(0.55), Inches(1.35), Inches(12.2), Inches(0.4),
    [P(R("TCO 56 646 €  ·  flux net 46 586 €/an  ·  point mort 14,6 mois  ·  ROI 310 % à 5 ans  ·  gain net 176 284 €", 12.5, NAVY, True))])
ntable(s, Inches(0.55), Inches(1.85),
    [Inches(3.0), Inches(4.6), Inches(4.65)],
    ["Chiffre", "Hypothèse retenue", "Pourquoi c’est solide / prudent"],
    [
     [[("40 €/h (alternant)", NAVY, True)], [("Taux interne chargé Axians", DARKTXT, False)], [("= coût réel pour l’entreprise, pas un salaire", DARKTXT, False)]],
     [[("2 000 €/client/an", NAVY, True)], [("Minimum contractuel (plancher)", DARKTXT, False)], [("Marché concurrent : 2 400 à 4 800 €/an", DARKTXT, False)]],
     [[("Coût humain = 82 %", NAVY, True)], [("Valorisé en actif immatériel", DARKTXT, False)], [("Templates + doc réutilisables, pas un coût perdu", DARKTXT, False)]],
     [[("OPEX 2 694 €/an", NAVY, True)], [("Fixe jusqu’à 50+ clients", DARKTXT, False)], [("Chaque nouveau client = flux net quasi intégral", DARKTXT, False)]],
     [[("2 déplacements/mois évités", NAVY, True)], [("Borne basse pessimiste", DARKTXT, False)], [("Toute imprécision joue en faveur du projet", DARKTXT, False)]],
     [[("Scénario 0 déplacement", NAVY, True)], [("Revenus = abonnements seuls", DARKTXT, False)], [("ROI encore 159 % → rentable sur l’abonnement seul", RGBColor(0x1E,0x8E,0x66), True)]],
    ], hfill=ACCENT, row_h=Inches(0.66), size=11)
txt(s, Inches(0.55), Inches(6.45), Inches(12.2), Inches(0.5),
    [P(R("Fiabilité du chiffrage : ", 11, GREY, True),
       R("1 001 h calculées tâche par tâche (GanttProject) et recoupées aux 176 jours réellement disponibles — cohérent à 0,7 % près.", 11, GREY, False))])
footer(s, 21)

# ================================================================ SLIDE 21 — HYPOTHÈSES CO2
s = slide()
rect(s, 0, 0, SW, SH, fill=WHITE)
annex_header(s, "Hypothèses du bilan CO2 — sources officielles", "Annexe 3 — Bilan carbone", badge_col=ACCENT2)
txt(s, Inches(0.55), Inches(1.35), Inches(12.2), Inches(0.4),
    [P(R("478 kg CO2 nets économisés / an  ·  le serveur émet 5,9× moins que les trajets évités  ·  ≈ 2,4 t sur 5 ans", 12.5, NAVY, True))])
ntable(s, Inches(0.55), Inches(1.85),
    [Inches(3.4), Inches(3.4), Inches(5.45)],
    ["Paramètre", "Valeur", "Source / justification"],
    [
     [[("Émission VUL diesel", NAVY, True)], [("180 g CO2/km", DARKTXT, False)], [("ADEME — Base Carbone 2023 (référentiel officiel)", DARKTXT, False)]],
     [[("Mix électrique français", NAVY, True)], [("56 g CO2/kWh", DARKTXT, False)], [("RTE éco2mix 2023 — faible car >70 % nucléaire", DARKTXT, False)]],
     [[("Puissance serveur HPE", NAVY, True)], [("200 W", DARKTXT, False)], [("HPE QuickSpecs (valeur en charge, majorante)", DARKTXT, False)]],
     [[("Déplacements évités", NAVY, True)], [("32 / an", DARKTXT, False)], [("2 interventions/client/an × 16 clients (borne basse)", DARKTXT, False)]],
     [[("Distance A/R", NAVY, True)], [("100 km", DARKTXT, False)], [("Seule hyp. non sourcée → sensibilité 50-200 km", RGBColor(0xB0,0x6A,0x0A), True)]],
     [[("Bilan net", NAVY, True)], [("576 − 98 kg", DARKTXT, False)], [("= 478 kg/an ; positif même à 50 km (+190 kg, 2,9×)", RGBColor(0x1E,0x8E,0x66), True)]],
    ], hfill=ACCENT2, row_h=Inches(0.66), size=11)
txt(s, Inches(0.55), Inches(6.45), Inches(12.2), Inches(0.5),
    [P(R("Le « paradoxe du serveur allumé H24 » est levé par le calcul, pas par l’intuition — et dans aucun scénario testé le bilan ne devient négatif.", 11, GREY, True))])
footer(s, 22)

# ================================================================ SLIDE 22 — Q/R CHOIX TECHNIQUES
s = slide()
rect(s, 0, 0, SW, SH, fill=WHITE)
annex_header(s, "Questions anticipées — les choix techniques", "Annexe 4 — « Pourquoi ce choix ? »", badge="Q / R")
qac = [
    ("Pourquoi Zabbix, pas PRTG/Centreon ?", "Open source (0 licence), moteur d’items très souple, couplage parfait avec Grafana, déjà urbanisé dans le groupe."),
    ("Pourquoi Teltonika, pas Cisco ?", "RMS nativement Cloud, VPN Hub en quelques clics, accès derrière le NAT opérateur, coût optimisé, Dual-SIM."),
    ("Pourquoi écarter Huawei / Robustel ?", "Souveraineté et recommandations ANSSI : clients santé / énergie, certains OIV."),
    ("Pourquoi un serveur physique, pas le cloud ?", "DAT + devis : le coût récurrent d’une VM dépassait l’achat d’un serveur dédié sur la durée de vie."),
    ("Pourquoi un polling à la minute ?", "Voir l’évolution réelle du signal : baisse lente = désalignement, chute brutale = panne franche."),
    ("À quoi servent les règles NoNat ?", "Préserver l’IP source de chaque IDU dans le VPN → identifier individuellement chaque faisceau en alarme."),
]
for k,(q,a) in enumerate(qac):
    x = Inches(0.55 + (k%2)*6.25); y = Inches(1.5 + (k//2)*1.72)
    qa_card(s, x, y, Inches(6.0), Inches(1.55), q, a)
footer(s, 23)

# ================================================================ SLIDE 23 — Q/R MÉTHODE & PIÈGES
s = slide()
rect(s, 0, 0, SW, SH, fill=WHITE)
annex_header(s, "Questions anticipées — méthode, chiffres & vigilance", "Annexe 5 — « Comment ? » + points de vigilance", badge="Q / R")
qac2 = [
    ("Pourquoi ce projet maintenant, pas avant ?", "Départs successifs + intégration VINCI/Axians (2019) ont mobilisé les ressources ; la bande passante s’est libérée + un alternant dédié."),
    ("Comment justifier les 1 001 heures ?", "Calcul tâche par tâche (GanttProject), recoupé aux 176 jours disponibles (81 %) — cohérent à 0,7 %."),
    ("Comment expliquer le retard de 4 semaines ?", "J1 MIB Ericsson non anticipable (+7 sem.) + 32 j de missions DIL/DIS. Sans ces causes externes → date prévue."),
    ("Comment le risque cyber est-il maîtrisé ?", "EBIOS RM : MFA strict (compromission) + NoNat/durcissement (latéralisation) → vraisemblance V3 ramenée à V1."),
]
for k,(q,a) in enumerate(qac2):
    x = Inches(0.55 + (k%2)*6.25); y = Inches(1.5 + (k//2)*1.55)
    qa_card(s, x, y, Inches(6.0), Inches(1.4), q, a)
# vigilance card
card(s, Inches(0.55), Inches(4.75), Inches(11.7), Inches(1.95), fill=RGBColor(0xFD,0xF3,0xE0))
rect(s, Inches(0.55), Inches(4.75), Inches(0.1), Inches(1.95), fill=AMBER)
txt(s, Inches(0.85), Inches(4.9), Inches(11.2), Inches(0.4),
    [P(R("⚠ VIGILANCE — deux chiffres à harmoniser avant l’oral (à annoncer soi-même si on les aborde)", 12.5, AMBER, True))])
bullet(s, Inches(0.85), Inches(5.4), Inches(11.2), [
    ("Déplacements évités : 24/an (ROI, 2/mois) vs 32/an (CO2, 2/client). ", "Deux bornes basses différentes — la conclusion ne dépend pas de l’hypothèse."),
    ("Puissance serveur : 150 W (OPEX électrique) vs 200 W (CO2). ", "Moyenne d’exploitation vs pic majorant — chaque section prend la valeur prudente pour son calcul."),
], size=11.5, gap=8, marker_col=AMBER)
footer(s, 24)

prs.save("/home/user/Soutenance/Soutenance_Aphelie_Ilyesse_Kebaili.pptx")
print("OK — slides:", len(prs.slides._sldIdLst))
