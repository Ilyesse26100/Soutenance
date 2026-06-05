# -*- coding: utf-8 -*-
"""Fiche de préparation aux questions du jury — Soutenance Aphélie."""
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm, mm
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER
from reportlab.platypus import (BaseDocTemplate, PageTemplate, Frame, Paragraph,
                                Spacer, Table, TableStyle, KeepTogether)

NAVY   = colors.HexColor("#0F1B2D")
BLUE   = colors.HexColor("#1B3A5C")
ACCENT = colors.HexColor("#00B4C8")
GREEN  = colors.HexColor("#2EC48F")
RED    = colors.HexColor("#E53E3E")
AMBER  = colors.HexColor("#C9870F")
LIGHT  = colors.HexColor("#EEF3F7")
CARD   = colors.HexColor("#F4F7FA")
GREY   = colors.HexColor("#5B6A7B")
DARK   = colors.HexColor("#1B2633")
LINE   = colors.HexColor("#D7E0E8")

styles = getSampleStyleSheet()
def st(name, **kw):
    base = kw.pop("parent", styles["Normal"])
    kw.setdefault("fontName", "Helvetica")
    return ParagraphStyle(name, parent=base, **kw)

H1   = st("H1", fontName="Helvetica-Bold", fontSize=15, textColor=NAVY, spaceBefore=10, spaceAfter=6, leading=18)
H2   = st("H2", fontName="Helvetica-Bold", fontSize=11.5, textColor=ACCENT, spaceBefore=8, spaceAfter=3, leading=14)
BODY = st("BODY", fontSize=9.5, textColor=DARK, leading=13, spaceAfter=4)
SMALL= st("SMALL", fontSize=8.3, textColor=DARK, leading=10.5)
SMALLW=st("SMALLW", fontSize=8.3, textColor=colors.white, leading=10.5)
SMB  = st("SMB", fontSize=8.3, textColor=NAVY, leading=10.5, fontName="Helvetica-Bold")
QST  = st("QST", fontSize=9, textColor=NAVY, leading=11.5, fontName="Helvetica-Bold")
ANS  = st("ANS", fontSize=8.7, textColor=DARK, leading=11)
LEAD = st("LEAD", fontSize=10, textColor=DARK, leading=14, spaceAfter=6)
TAGW = st("TAGW", fontSize=9, textColor=colors.white, leading=12, fontName="Helvetica-Bold")

W, Hh = A4
MX = 1.6*cm
FW = W - 2*MX

def header_footer(canvas, doc):
    canvas.saveState()
    canvas.setFillColor(NAVY); canvas.rect(0, Hh-1.15*cm, W, 1.15*cm, fill=1, stroke=0)
    canvas.setFillColor(ACCENT); canvas.rect(0, Hh-1.20*cm, W, 0.06*cm, fill=1, stroke=0)
    canvas.setFillColor(colors.white); canvas.setFont("Helvetica-Bold", 10.5)
    canvas.drawString(MX, Hh-0.78*cm, "APHÉLIE — Fiche de préparation aux questions du jury")
    canvas.setFillColor(ACCENT); canvas.setFont("Helvetica-Bold", 8.5)
    canvas.drawRightString(W-MX, Hh-0.78*cm, "Ilyesse KEBAILI · BUT 3 R&T Cyber")
    canvas.setFillColor(GREY); canvas.setFont("Helvetica", 7.5)
    canvas.drawString(MX, 0.7*cm, "Anticiper le « pourquoi » et le « comment » — chaque chiffre justifié, chaque choix argumenté.")
    canvas.drawRightString(W-MX, 0.7*cm, "Page %d" % canvas.getPageNumber())
    canvas.restoreState()

doc = BaseDocTemplate("/home/user/Soutenance/Fiche_Preparation_Questions_Jury.pdf",
                      pagesize=A4, leftMargin=MX, rightMargin=MX,
                      topMargin=1.55*cm, bottomMargin=1.1*cm)
frame = Frame(MX, 1.1*cm, FW, Hh-1.55*cm-1.1*cm, id='n', leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0)
doc.addPageTemplates([PageTemplate(id='main', frames=[frame], onPage=header_footer)])

E = []
def band(text, color=NAVY):
    t = Table([[Paragraph('<font color="white"><b>%s</b></font>' % text, st("b", fontSize=11.5))]],
              colWidths=[FW])
    t.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,-1),color),
        ("LEFTPADDING",(0,0),(-1,-1),8),("TOPPADDING",(0,0),(-1,-1),5),
        ("BOTTOMPADDING",(0,0),(-1,-1),5),("LINEBELOW",(0,0),(-1,-1),2,ACCENT)]))
    return t

def info_table(rows, head, widths, headcolor=BLUE):
    data = [[Paragraph('<b>%s</b>'%h, TAGW) for h in head]]
    for r in rows:
        data.append([Paragraph(c, SMALL) if i>0 else Paragraph(c, SMB) for i,c in enumerate(r)])
    t = Table(data, colWidths=widths, repeatRows=1)
    ts = [("BACKGROUND",(0,0),(-1,0),headcolor),
          ("TEXTCOLOR",(0,0),(-1,0),colors.white),
          ("GRID",(0,0),(-1,-1),0.4,LINE),
          ("ROWBACKGROUNDS",(0,1),(-1,-1),[colors.white, CARD]),
          ("VALIGN",(0,0),(-1,-1),"TOP"),
          ("LEFTPADDING",(0,0),(-1,-1),5),("RIGHTPADDING",(0,0),(-1,-1),5),
          ("TOPPADDING",(0,0),(-1,-1),4),("BOTTOMPADDING",(0,0),(-1,-1),4)]
    t.setStyle(TableStyle(ts))
    return t

# ============================ INTRO
E.append(band("0 · COMMENT UTILISER CETTE FICHE", NAVY))
E.append(Spacer(1,4))
E.append(Paragraph("Objectif : <b>désamorcer les questions « pourquoi ? » et « comment ? » avant qu’elles ne soient posées</b>, "
    "en intégrant la justification de chaque choix et de chaque chiffre directement dans votre exposé. "
    "Une hypothèse annoncée et sourcée par vous-même devient une force ; la même hypothèse révélée par une question du jury devient une faiblesse.", LEAD))
E.append(Paragraph("<b>Règle d’or :</b> pour chaque chiffre, dites toujours trois choses — <b>la valeur</b>, <b>l’hypothèse retenue</b> "
    "et <b>pourquoi elle est conservatrice</b> (c.-à-d. qu’elle joue contre vous, donc la réalité est meilleure).", BODY))

# ============================ 1 — LE BESOIN
E.append(Spacer(1,6))
E.append(band("1 · LE BESOIN — POURQUOI, POUR QUI, COMMENT, OÙ"))
E.append(Spacer(1,4))
E.append(Paragraph("POURQUOI (le problème à résoudre)", H2))
E.append(Paragraph(
    "• L’ancienne supervision était <b>obsolète</b> : la v2012 (étudiants ESISAR) était une « boîte noire » non maintenable — "
    "<b>1 module sur 2 hors service</b>, fausses alarmes (rouge sur des sites sains), modules muets non détectés.<br/>"
    "• La tentative v2017 (WhatsUp Gold) est restée <b>inaboutie</b> au stade de l’étude des MIB.<br/>"
    "• Résultat : <b>~3 ans sans outil fonctionnel</b> → les pannes étaient signalées par les appels des clients (l’inverse de l’objectif).<br/>"
    "• <b>Criticité extrême</b> : parc majoritairement composé d’<b>hôpitaux</b> (rupture de la chaîne de soins) et de <b>sites industriels en production continue</b> "
    "(arrêt immédiat) ; certains clients sont des <b>OIV</b>. La disponibilité doit être de tous les instants.<br/>"
    "• Principe directeur : <i>« une mauvaise supervision est plus dangereuse qu’une absence de supervision »</i> (fausse confiance).", SMALL))
E.append(Spacer(1,4))
E.append(Paragraph("POUR QUI (les bénéficiaires)", H2))
E.append(info_table([
    ["Techniciens / NOC", "Utilisateurs finaux : voir l’état du parc d’un coup d’œil, diagnostic ciblé avec la bonne pièce, moins de déplacements."],
    ["Chef de projet & Resp. d’affaires", "Pilotage, preuve du service rendu, données pour les négociations commerciales."],
    ["Clients (hôpitaux, industrie, RTE/EDF)", "Continuité de service + preuve de conformité (Rx mesuré confronté au Rx théorique de la DIL)."],
    ["L’agence Axians RMP", "Rentabilité, levier de renégociation des contrats de maintenance, actif immatériel et avantage concurrentiel durable."],
], ["Partie prenante", "Ce que la supervision lui apporte"], [5.2*cm, FW-5.2*cm]))
E.append(Spacer(1,4))
E.append(Paragraph("COMMENT (la réponse technique, en une phrase par maillon)", H2))
E.append(Paragraph(
    "• <b>Sur site</b> : routeur <b>Teltonika RUT956</b> (routeur + pare-feu + VPN en un boîtier) branché sur le port de supervision de l’IDU hertzien.<br/>"
    "• <b>Transport</b> : <b>tunnel VPN chiffré</b> sur le réseau IP déjà payé par le client + Cloud <b>RMS</b> pour administrer à distance (contourne le NAT opérateur, MFA imposé).<br/>"
    "• <b>Collecte</b> : <b>Zabbix</b>, SNMP en polling 1 min (Rx/Tx) + traps spontanés ; templates génériques bâtis sur un <b>référentiel OID Ericsson reconstruit</b>.<br/>"
    "• <b>Restitution</b> : <b>Grafana</b> (hexagones temps réel, cartographie GeoJSON via le script maison <i>grafana_topology.py</i>), tables Rx/Tx, SLO par client, e-mails alerte + résolution.<br/>"
    "• <b>Sécurité</b> : analyse <b>EBIOS RM</b> → MFA strict + durcissement RUT956 (REJECT par défaut, NoNat ciblé, cloisonnement de zones).<br/>"
    "• <b>Hébergement</b> : serveur physique <b>HPE DL360 (Debian 13)</b>, retenu après comparaison de coût (DAT + devis) face à une VM cloud.", SMALL))
E.append(Spacer(1,4))
E.append(Paragraph("OÙ / QUOI (le périmètre — ce qui est dans le projet et ce qui n’y est pas)", H2))
E.append(Paragraph(
    "• <b>Où</b> : sites clients dispersés en France (zones non desservies par la fibre) ; serveur hébergé dans l’infrastructure Axians RMP.<br/>"
    "• <b>Dans le périmètre</b> : <b>16 clients</b> sous contrat de maintenance actif, <b>30 à 50 faisceaux</b> Ericsson MINI-LINK ; surveillance limitée aux <b>IDU</b>.<br/>"
    "• <b>Extensions hors périmètre initial</b> (assumées) : prototypes d’IA locale, portail centralisé Entra ID, assistant Mistral — présentés comme des <b>pistes</b>, pas des fonctions industrialisées.", SMALL))

# ============================ 2 — HYPOTHÈSES ROI
E.append(Spacer(1,6))
E.append(band("2 · LES HYPOTHÈSES DU ROI — CHAQUE CHIFFRE JUSTIFIÉ"))
E.append(Spacer(1,4))
E.append(Paragraph("Résultats à retenir : <b>TCO 56 646 €</b> · <b>flux net 46 586 €/an</b> · <b>point mort 14,6 mois</b> · "
    "<b>ROI 310 % à 5 ans</b> (gain net 176 284 €). Tout repose sur le <b>scénario le plus conservateur</b>.", BODY))
roi = [
 ["Taux alternant 40 €/h", "Coût horaire chargé interne Axians (et non un salaire). 1 001 h × 40 € = 40 040 €.", "C’est le taux de refacturation interne ; il valorise un coût réel pour l’entreprise, pas ma rémunération."],
 ["Effort 1 001 h (142,8 j)", "Calculé tâche par tâche dans GanttProject, puis vérifié : 176 j réellement disponibles en entreprise → 81 % du temps.", "Le chiffre est recoupé avec le calendrier réel d’alternance à 0,7 % près : il n’est pas estimé, il est vérifié."],
 ["CAPEX matériel 10 206 €", "Serveur HPE 5 694 € + 16× RUT956 = 3 392 € + licences RMS 480 € + API RMS 640 € + logiciels 0 € (open source).", "Le matériel ne pèse que 18 % : la valeur est dans le savoir-faire, pas dans l’achat."],
 ["CAPEX humain = 82 % du TCO", "Alternant 40 040 € + chef de projet (40 h×120 €) + resp. affaires (10 h×160 €).", "Caractéristique d’un projet R&D en régie : ce n’est pas un coût perdu mais un <b>actif immatériel</b> (templates, doc, savoir-faire) réutilisable."],
 ["OPEX 2 694 €/an", "Élec 264 € + MCO Zabbix/Grafana 1 056 € + évol. templates 696 € + amort. routeurs 678 €. Aucun coût de connectivité (VPN sur réseau client).", "L’OPEX est <b>fixe</b> jusqu’à 50+ clients → chaque nouveau client signé alimente presque intégralement le flux net (effet de levier)."],
 ["Tarif 2 000 €/client/an", "= minimum contractuel (plancher sous lequel Axians refuse de signer). 16 × 2 000 = 32 000 €/an.", "Choix volontairement conservateur : le marché concurrent facture 2 400–4 800 €/an. À 2 500 € le ROI passe à 450 %."],
 ["Économies GTI 17 280 €/an", "Déplacements physiques rendus inutiles par la télémaintenance, valorisés comme revenu.", "Plancher de revenus <b>indépendant du tarif</b> : même si les abonnements baissaient, les économies GTI subsisteraient."],
 ["ROI −18 % à 12 mois", "Investissement concentré sur 7 mois alors que les revenus s’accumulent dans le temps.", "Ce n’est pas un projet en difficulté : à 12 mois, <b>83 % du capital est déjà récupéré</b>."],
]
E.append(info_table(roi, ["Chiffre / hypothèse", "D’où vient-il", "Comment le défendre devant le jury"],
                    [3.5*cm, 6.2*cm, FW-3.5*cm-6.2*cm], headcolor=ACCENT))
E.append(Spacer(1,4))
E.append(Paragraph("Test de résistance à connaître : <b>année sans aucun déplacement évité</b> → revenus 32 000 € seuls, flux net 29 306 €, "
    "point mort 23 mois, <b>ROI encore 159 %</b>. Conclusion à marteler : <i>la rentabilité tient sur l’abonnement seul ; les économies GTI ne sont qu’un accélérateur.</i>", SMB))

# ============================ 3 — HYPOTHÈSES CO2
E.append(Spacer(1,6))
E.append(band("3 · LES HYPOTHÈSES DU BILAN CO2 — CHAQUE CHIFFRE SOURCÉ"))
E.append(Spacer(1,4))
E.append(Paragraph("Résultat : <b>478 kg CO2 nets économisés/an</b> (≈ 2,4 t sur 5 ans). Le serveur émet <b>5,9× moins</b> "
    "que les déplacements qu’il remplace. Trois des quatre paramètres sont <b>sourcés officiellement</b>.", BODY))
co2 = [
 ["Émission VUL diesel 180 g CO2/km", "Base Carbone <b>ADEME</b> v2023 — référentiel officiel français.", "Source publique et reconnue : non discutable."],
 ["Mix élec FR 56 g CO2/kWh", "<b>RTE éco2mix</b>, moyenne annuelle 2023.", "Faible car &gt;70 % de nucléaire en France (vs ~400 g en Allemagne) : avantage structurel du territoire."],
 ["Puissance serveur 200 W", "Documents techniques <b>HPE QuickSpecs</b> (charge de supervision).", "Valeur constructeur ; prise « en charge », donc majorante (donc prudente côté CO2 du serveur)."],
 ["32 déplacements évités/an", "2 interventions évitées par client et par an, sur 16 clients.", "Borne basse : un parc en production continue génère bien plus d’interventions terrain."],
 ["Distance A/R 100 km", "<b>Seule hypothèse non sourcée</b> (pas de relevé historique des km).", "Assumé ouvertement → <b>analyse de sensibilité 50→200 km</b> : même à 50 km le bilan reste positif (+190 kg, 2,9×)."],
 ["Calcul évité : 576 kg/an", "100 km × 180 g = 18 kg/trajet × 32 trajets.", "Calcul transparent, reproductible."],
 ["Calcul serveur : 98 kg/an", "200 W × 8 760 h = 1 752 kWh × 56 g.", "Le « paradoxe du serveur allumé H24 » est levé par le chiffre, pas par l’intuition."],
]
E.append(info_table(co2, ["Paramètre", "Source / calcul", "Comment le défendre"],
                    [4.3*cm, 5.4*cm, FW-4.3*cm-5.4*cm], headcolor=GREEN))

# ============================ 4 — PIÈGES
E.append(Spacer(1,6))
E.append(band("4 · LES DEUX PIÈGES CHIFFRÉS À CONNAÎTRE (à désamorcer soi-même)", AMBER))
E.append(Spacer(1,4))
trap = [
 ["Déplacements évités : 24 vs 32 ?",
  "Le ROI retient <b>2/mois au total (24/an)</b> ; le bilan CO2 retient <b>2/client/an (32/an)</b>.",
  "« Le chiffrage financier prend une borne <b>encore plus basse</b> pour ne jamais surévaluer les revenus ; le bilan carbone raisonne par client. Les deux sont des bornes basses pessimistes et, dans les deux cas, la conclusion ne dépend pas de l’hypothèse. » <b>(Idéalement : harmoniser sur 24 ou 32 avant l’oral.)</b>"],
 ["Puissance serveur : 150 W vs 200 W ?",
  "L’OPEX électrique utilise <b>150 W</b> ; le bilan CO2 utilise <b>200 W</b>.",
  "« 150 W = consommation moyenne d’exploitation ; 200 W = pic en charge de supervision, volontairement majorant côté carbone. » Chaque section prend la valeur la plus prudente pour SON calcul."],
]
E.append(info_table(trap, ["Le point", "L’écart", "La réponse à dégainer"],
                    [3.6*cm, 4.5*cm, FW-3.6*cm-4.5*cm], headcolor=AMBER))

# ============================ 5 — BANQUE Q/R
E.append(Spacer(1,6))
E.append(band("5 · BANQUE DE QUESTIONS « POURQUOI / COMMENT » — RÉPONSES PRÊTES"))
E.append(Spacer(1,4))
qa = [
 ("Pourquoi remplacer l’ancien outil plutôt que le réparer ?",
  "Code ESISAR non maintenable (boîte noire), 1 module/2 HS, plus aucun correctif ; v2017 inaboutie. Repartir proprement coûtait moins que rafistoler un système qui donnait de fausses alarmes."),
 ("Pourquoi ce projet maintenant et pas plus tôt ?",
  "Départs successifs (resp. info + stagiaire) ont cassé la dynamique, puis l’intégration VINCI/Axians (2019) a mobilisé les ressources pendant des années. Aphélie arrive quand la bande passante s’est libérée + un alternant dédié."),
 ("Pourquoi Zabbix plutôt que PRTG ou Centreon ?",
  "100 % open source (0 licence, scalabilité), moteur d’items/triggers très souple, couplage parfait avec Grafana, et déjà urbanisé dans le groupe VINCI Energies."),
 ("Pourquoi Teltonika plutôt que Cisco ?",
  "RMS nativement Cloud, VPN Hub déployé en quelques clics, accès même derrière le NAT opérateur, coût optimisé (vs licences Cisco), Dual-SIM pour la HA, et REX favorables dans le groupe."),
 ("Pourquoi avoir écarté Huawei / Robustel ?",
  "Souveraineté et recommandations ANSSI : nos clients relèvent de la santé et de l’énergie, certains sont OIV → incompatibilité géopolitique/réglementaire."),
 ("Pourquoi un serveur physique et pas le cloud ?",
  "Un DAT + un devis Axians Cloud : le coût récurrent de la VM, rapporté à la durée de vie, dépassait l’achat d’un serveur dédié. Décision prise sur la comparaison financière."),
 ("Pourquoi un polling à la minute ?",
  "Pour voir l’évolution réelle du signal : une baisse lente minute par minute = désalignement/obstacle ; une chute brutale = panne franche. Un simple test de joignabilité ne le montrerait pas."),
 ("À quoi servent concrètement les règles NoNat ?",
  "À préserver l’IP source de chaque IDU à travers le VPN. Sans elles, le masquerade rendrait tous les faisceaux d’un site identiques dans Zabbix → supervision individuelle impossible."),
 ("Comment avez-vous traité le risque cyber ?",
  "EBIOS RM : scénario d’attaque par rebond (G4). Deux mesures — MFA strict sur RMS (bloque la compromission) + durcissement RUT956/NoNat (bloque la latéralisation). Vraisemblance ramenée de V3 à V1."),
 ("Comment justifiez-vous les 1 001 heures ?",
  "Calcul tâche par tâche dans GanttProject, recoupé avec 176 jours réellement disponibles (81 % du temps) — cohérent à 0,7 % près."),
 ("Pourquoi 2 000 € par client et par an ?",
  "C’est le minimum contractuel (plancher), donc le pire cas signable. Le marché facture 2 400 à 4 800 €/an. Je présente le scénario le plus prudent ; tout est meilleur en réalité."),
 ("Comment expliquez-vous les 4 semaines de retard ?",
  "Décomposé honnêtement : +7 sem. sur J1 (MIB Ericsson non documentées, expertise inexistante en interne, non anticipable) + 32 j de missions DIL/DIS contractuelles. En retirant ces deux causes externes, la fin tombe à la date prévue."),
 ("Le serveur tourne H24 : est-ce vraiment écologique ?",
  "Oui : 98 kg CO2/an pour le serveur contre 576 kg évités = 5,9× moins. La conclusion tient même dans l’hypothèse la plus défavorable (50 km)."),
 ("Pourquoi le clustering DBSCAN / un modèle local ?",
  "DBSCAN = non supervisé, il regroupe des alarmes proches dans le temps sans catégories prédéfinies. Tout en local (Ollama/Mistral) pour ne jamais sortir de données réseau vers un cloud tiers — posture de confidentialité."),
]
data = []
for i,(q,a) in enumerate(qa):
    data.append([Paragraph("<b>Q%d.</b>"%(i+1), QST),
                 Paragraph(q, QST),
                 Paragraph(a, ANS)])
t = Table(data, colWidths=[1.0*cm, 5.2*cm, FW-1.0*cm-5.2*cm])
t.setStyle(TableStyle([
    ("VALIGN",(0,0),(-1,-1),"TOP"),
    ("ROWBACKGROUNDS",(0,0),(-1,-1),[colors.white, CARD]),
    ("LINEBELOW",(0,0),(-1,-1),0.4,LINE),
    ("TEXTCOLOR",(0,0),(0,-1),ACCENT),
    ("LEFTPADDING",(0,0),(-1,-1),5),("RIGHTPADDING",(0,0),(-1,-1),5),
    ("TOPPADDING",(0,0),(-1,-1),5),("BOTTOMPADDING",(0,0),(-1,-1),5),
]))
E.append(t)
E.append(Spacer(1,6))
E.append(Paragraph("Conseil de clôture : terminez chaque réponse chiffrée par la phrase « <i>… et c’est l’hypothèse la plus prudente</i> ». "
    "Le jury cherche votre capacité d’analyse et votre recul, pas seulement la technique : montrez que vous connaissez les limites de vos propres chiffres mieux que lui.", SMB))

doc.build(E)
print("PDF généré")
