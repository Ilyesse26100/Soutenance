# -*- coding: utf-8 -*-
"""Ajoute le script oral en notes du présentateur (pptx) + génère un PDF imprimable."""
from pptx import Presentation

PPTX = "/home/user/Soutenance/Soutenance_Aphelie_Ilyesse_Kebaili.pptx"

# (numéro, titre court, durée, langue, texte à dire)
SCRIPT = [
 (1, "Titre", "0:20", "FR",
  "Bonjour. Je m’appelle Ilyesse Kebaili, je suis alternant en BUT Réseaux et Télécommunications, "
  "parcours Cybersécurité, à l’IUT de Valence. J’effectue mon alternance chez Axians Réseaux Mobiles "
  "Privés, une entité de VINCI Energies. Je vais vous présenter mon projet de réalisation professionnelle : "
  "Aphélie, la conception d’un système de supervision centralisé pour des faisceaux hertziens."),
 (2, "Plan", "0:25", "FR",
  "Ma présentation se déroulera en cinq temps. Après une courte introduction, je vous présenterai en anglais "
  "l’entreprise et le contexte de mon travail. J’aborderai ensuite le cadrage du projet, sa réalisation "
  "technique, puis ses impacts financiers et environnementaux, avant de conclure sur le bilan et les "
  "perspectives. Des annexes me permettront de répondre précisément à vos questions."),
 (3, "Introduction & problématique", "1:30", "FR",
  "Commençons par le contexte. Les faisceaux hertziens sont des liaisons radio point à point qui raccordent "
  "des sites isolés, ou servent de secours, là où la fibre n’arrive pas. Nos clients sont surtout des hôpitaux "
  "et des sites industriels en production continue : pour eux, une coupure de liaison, c’est une rupture de la "
  "chaîne de soins, ou un arrêt de production. La disponibilité est donc critique.\n\n"
  "Or l’agence se trouvait dans une situation problématique : son outil de supervision était devenu obsolète. "
  "Pendant près de trois ans, les pannes n’étaient détectées que lorsque le client lui-même appelait — "
  "exactement l’inverse de ce qu’on attend d’une supervision.\n\n"
  "D’où ma problématique : comment remplacer une solution de supervision obsolète par une plateforme "
  "centralisée, capable de fournir des informations fiables en temps réel sur les liaisons hertziennes, et "
  "accessible à distance ? Pour y répondre, le projet s’est fixé quatre objectifs : centraliser, fiabiliser, "
  "sécuriser et pérenniser."),
 (4, "[EN] The company", "2:30", "EN",
  "I will now present the company and the context of my work in English.\n\n"
  "I did my apprenticeship at Axians Réseaux Mobiles Privés. Let me put it in context. The company belongs "
  "to VINCI, a global group with 280,000 employees and a revenue of 74.6 billion euros in 2025. Within VINCI, "
  "the energy and IT division is VINCI Energies, and Axians is its brand dedicated to information and "
  "communication technologies, with 15,000 employees worldwide. My agency, Axians RMP AURA, is based near "
  "Valence, and is a human-sized team of about 38 people.\n\n"
  "Our specialties are private radio networks: TETRA, private 4G and 5G, and microwave radio links. Our "
  "customers are mostly hospitals and industrial sites running in continuous production, which means a very "
  "high availability requirement.\n\n"
  "Two concrete examples: we deployed a private 4G network for EDF Renewables on a wind farm, and we "
  "modernised the TETRA radio network for Sytral, covering more than a thousand buses and a hundred tramways.\n\n"
  "This explains why the Aphélie project emerged in this agency: a scattered customer base with no fibre, "
  "critical clients where an outage is a serious matter, and no working supervision tool for nearly three "
  "years, after two failed attempts."),
 (5, "[EN] My role & context", "2:30", "EN",
  "Now, my position and my missions. I work as a telecom technician. My core mission was the full design and "
  "deployment of the Aphélie project — from the initial study to the production server. I was given a lot of "
  "autonomy: I made the architecture choices, I tested everything in an isolated lab, and I documented each "
  "step. My key constraint was simple but strict: never disrupt the customers’ live production networks.\n\n"
  "In parallel, I also worked on transverse missions in radio engineering. I analysed the technical "
  "specifications of major accounts such as RTE and EDF, and I carried out link studies with a software called "
  "HTZ — modelling the terrain, the Fresnel zone, the clearance and the rainfall. One concrete case was a "
  "feasibility study near a wind turbine, where I validated by calculation a safety margin of 16.3 metres "
  "between the blade and the radio beam.\n\n"
  "These studies were not a distraction from the project — quite the opposite. Mastering the link budget is "
  "exactly what defines the real alarm thresholds I had to supervise. They gave me the field knowledge of "
  "what the supervision actually had to monitor. Thank you — I will now continue in French."),
 (6, "Diagnostic & besoins", "1:15", "FR",
  "Je reprends en français. Avant de concevoir quoi que ce soit, il fallait comprendre pourquoi les versions "
  "précédentes avaient échoué. La première, en 2012, développée par des étudiants, était devenue une boîte "
  "noire impossible à maintenir : un module sur deux ne fonctionnait plus, et l’outil affichait de fausses "
  "alarmes. La seconde, en 2017, n’a jamais dépassé le stade de l’étude. D’où ma priorité : livrer enfin une "
  "solution stable et fiable.\n\n"
  "À partir de là, j’ai défini le besoin avec mon chef de projet : superviser 16 clients et une trentaine de "
  "faisceaux Ericsson, avec une relève toutes les minutes des niveaux de réception et d’émission, des alertes "
  "hiérarchisées et traduites en clair, une cartographie dynamique, et des notifications automatiques. Et "
  "surtout, une plateforme assez fiable pour ne pas devenir elle-même le point faible."),
 (7, "Choix technologiques", "1:00", "FR",
  "Pour répondre à ce besoin, j’ai mené une véritable étude de marché, par filtres successifs, sur deux "
  "briques. D’abord le transport : j’ai comparé Cisco, Sierra Wireless et Teltonika. J’ai retenu le routeur "
  "Teltonika RUT956, parce que sa plateforme de gestion est nativement dans le cloud, qu’il établit un VPN en "
  "quelques clics même derrière le NAT de l’opérateur, et que son coût est optimisé.\n\n"
  "Ensuite le moteur de supervision : face à PRTG et Centreon, j’ai choisi Zabbix. Il est open source — donc "
  "zéro licence —, son moteur est très souple, il se couple parfaitement à Grafana, et il était déjà utilisé "
  "dans le groupe. Deux choix qui servent aussi l’urbanisation du système d’information."),
 (8, "Planning & écarts", "1:00", "FR",
  "J’ai découpé le projet en cinq jalons séquentiels, chacun conditionnant le suivant. Le projet a pris quatre "
  "semaines de retard, mais je tiens à l’expliquer précisément, car ces retards ne viennent pas d’une mauvaise "
  "estimation. Le premier jalon a pris sept semaines de plus : les fichiers techniques des équipements "
  "Ericsson n’étaient pas documentés, et personne en interne n’avait cette expertise — c’était structurellement "
  "impossible à anticiper. Le reste vient de 32 jours de missions urgentes pour d’autres clients, prioritaires "
  "car contractuelles. En retirant ces deux causes externes, le projet tombe exactement à la date prévue. Et "
  "les cinq jalons ont été livrés à 100 %."),
 (9, "Cybersécurité — EBIOS RM", "1:00", "FR",
  "Ouvrir des flux depuis les sites clients vers notre serveur crée de nouvelles vulnérabilités. J’ai donc "
  "conduit une analyse de risque avec la méthode EBIOS Risk Manager, la méthode officielle de l’ANSSI. J’ai "
  "modélisé le pire scénario : une attaque par rebond, où un attaquant vole les accès d’un technicien, atteint "
  "un routeur, puis rebondit vers le réseau du client pour y propager un rançongiciel — un risque critique.\n\n"
  "Pour le neutraliser, deux mesures : l’authentification multifacteur sur la plateforme de gestion, qui "
  "bloque la compromission initiale, et un durcissement des routeurs, qui bloque tout rebond. Résultat : la "
  "vraisemblance passe de très élevée à peu vraisemblable, ce qui rend l’architecture compatible avec une mise "
  "en production sécurisée."),
 (10, "Architecture matérielle", "1:00", "FR",
  "Voici l’architecture. Le principe directeur : placer l’intelligence au plus près des équipements. Sur le "
  "site client, le routeur RUT956 regroupe routeur, pare-feu et VPN dans un seul boîtier, branché sur le port "
  "de supervision du faisceau. Les données remontent par un tunnel VPN chiffré, en empruntant la connexion "
  "Internet déjà payée par le client — donc sans abonnement supplémentaire. Côté Axians, Zabbix collecte les "
  "données, Grafana les visualise, le tout hébergé sur un serveur physique HPE. Le déploiement est Plug and "
  "Play : je préconfigure en atelier, puis la mise en service se fait à distance, par téléphone, sans "
  "déplacement."),
 (11, "Intégration Zabbix & MIB", "1:00", "FR",
  "Le cœur technique du projet, et la partie la plus difficile, a été d’intégrer les équipements dans Zabbix. "
  "Les faisceaux Ericsson exposent leurs données via des centaines d’identifiants bruts, non documentés. J’ai "
  "dû capturer les trames directement sur les équipements pour identifier les bons. Avec des pièges : par "
  "exemple, la puissance reçue n’est pas remontée en décibels, mais en dixièmes — une valeur de moins 450 vaut "
  "en réalité moins 45. Sans correction, tout serait faux d’un facteur dix. J’ai aussi dû apprendre les "
  "expressions régulières pour décoder les alarmes. Le résultat, c’est un référentiel d’identifiants validés : "
  "un actif technique qui n’existait pas dans l’entreprise, et qui resservira pour les prochains clients."),
 (12, "Durcissement & NoNat", "1:00", "FR",
  "Côté sécurité réseau, j’ai appliqué une posture fermée par défaut : tout ce qui n’est pas explicitement "
  "autorisé est rejeté, et l’administration passe uniquement par le tunnel chiffré.\n\n"
  "Un point que je veux mettre en avant, ce sont les règles NoNat. Normalement, un routeur masque l’adresse de "
  "tous les équipements derrière la sienne. Problème : dans Zabbix, tous les faisceaux d’un même site "
  "apparaîtraient identiques, impossible de les distinguer. Les règles NoNat désactivent ce masquage, "
  "uniquement pour le trafic de supervision : on préserve l’adresse de chaque faisceau, et on peut identifier "
  "précisément lequel est en alarme, tout en gardant le cloisonnement de sécurité."),
 (13, "Restitution & cartographie", "1:00", "FR",
  "Toutes ces données, il fallait les rendre lisibles. J’ai mis en place deux niveaux. Le niveau opérationnel, "
  "dans Zabbix, avec des panneaux hexagonaux qui affichent la puissance reçue de chaque équipement en temps "
  "réel et passent au rouge en cas d’alarme, plus un suivi de la disponibilité par client. Et le niveau "
  "cartographique, sous Grafana, où chaque liaison est tracée sur une carte de France : le trait devient rouge "
  "quand un faisceau est en défaut. Pour cela, j’ai développé un script qui interroge la base de Zabbix et "
  "génère la carte automatiquement. Concrètement, un technicien voit d’un coup d’œil quel site est en panne, "
  "et où il se trouve."),
 (14, "IA locale", "0:45", "FR",
  "Au-delà du cahier des charges, j’ai exploré deux prototypes d’intelligence artificielle, exécutés "
  "entièrement en local, pour ne jamais sortir de données vers un cloud tiers : un premier pour corréler des "
  "alarmes et faire émerger une cause commune ; un second, un assistant conversationnel, pour interroger la "
  "supervision en langage naturel. Je reste transparent : ce sont des pistes prometteuses, pas encore des "
  "fonctions de production."),
 (15, "Rentabilité (ROI)", "1:15", "FR",
  "Venons-en aux impacts. Sur le plan financier, le coût total du projet est de 56 600 euros, dont 82 % de "
  "temps humain — autrement dit, l’essentiel de la valeur est un savoir-faire interne, un actif réutilisable, "
  "et non du matériel acheté. En face, le projet génère un flux net de 46 600 euros par an. Résultat : le "
  "point mort est atteint en moins de quinze mois, et le retour sur investissement atteint 310 % à cinq ans. "
  "Et j’insiste : ces chiffres reposent sur le scénario le plus prudent, le tarif plancher. Même sans aucune "
  "économie de déplacement, le projet reste rentable, à 159 %."),
 (16, "Environnement & opérationnel", "1:00", "FR",
  "Sur le plan environnemental, le groupe VINCI impose des objectifs de réduction carbone. Aphélie remplace "
  "des déplacements par une surveillance à distance. On pourrait objecter qu’un serveur allumé en permanence "
  "consomme : j’ai donc vérifié par le calcul, avec des sources officielles, l’ADEME et RTE. Le serveur émet "
  "98 kilos de CO2 par an, contre 576 kilos évités par les déplacements supprimés : un bilan net positif de "
  "478 kilos par an, et cela reste vrai même dans l’hypothèse la plus défavorable.\n\n"
  "Il y a enfin un troisième bénéfice, opérationnel : je peux désormais comparer en temps réel la puissance "
  "mesurée d’une liaison à sa valeur théorique. Ce qui prenait plusieurs jours se vérifie en quelques minutes."),
 (17, "Bilan & perspectives", "1:00", "FR",
  "En conclusion, la problématique trouve une réponse complète : la plateforme est centralisée sur un serveur "
  "unique, fiable grâce à l’investigation des équipements, en temps réel par une relève à la minute, et "
  "accessible à distance par les tunnels VPN. Les cinq jalons sont atteints.\n\n"
  "Pour l’entreprise, l’intérêt dépasse l’outil : c’est un actif durable, documenté et réutilisable, qui place "
  "l’agence en capacité de proposer une offre de supervision différenciée. La principale perspective est "
  "d’élargir le parc : chaque nouveau client est désormais déployable en un à deux jours, contre plusieurs "
  "semaines au départ."),
 (18, "Conclusion & remerciements", "0:45", "FR",
  "Sur le plan personnel, ce projet m’a permis de monter en compétences sur la supervision, les réseaux et la "
  "cybersécurité opérationnelle, mais aussi sur la conduite de projet. Il a confirmé mon orientation : je "
  "poursuis en école d’ingénieur en alternance, en restant chez Axians pour contribuer à l’évolution de la "
  "plateforme.\n\nJe vous remercie de votre attention, et je suis à votre disposition pour vos questions."),
 (19, "Annexe — Séparateur", "réserve", "FR",
  "RÉSERVE — diapositive de transition vers les annexes pendant les questions."),
 (20, "Annexe — Le besoin", "réserve", "FR",
  "RÉSERVE — projeter si le jury demande de préciser le besoin, les bénéficiaires ou le périmètre."),
 (21, "Annexe — Hypothèses ROI", "réserve", "FR",
  "RÉSERVE — projeter pour toute question sur les chiffres financiers. Réflexe : valeur → hypothèse → "
  "« c’est le scénario le plus prudent »."),
 (22, "Annexe — Hypothèses CO2", "réserve", "FR",
  "RÉSERVE — projeter pour toute question sur le bilan carbone. Sources : ADEME, RTE, HPE ; seule la distance "
  "de 100 km est une hypothèse, couverte par l’analyse de sensibilité."),
 (23, "Annexe — Q/R choix techniques", "réserve", "FR",
  "RÉSERVE — réponses prêtes sur les choix : Zabbix, Teltonika, serveur physique, polling, NoNat."),
 (24, "Annexe — Q/R méthode & vigilance", "réserve", "FR",
  "RÉSERVE — réponses méthode + les deux chiffres à harmoniser (déplacements 24/32 ; serveur 150/200 W) : "
  "les annoncer soi-même si le sujet vient."),
]

# --- 1) Notes du présentateur ---
prs = Presentation(PPTX)
for sl, item in zip(prs.slides, SCRIPT):
    tf = sl.notes_slide.notes_text_frame
    tf.text = "[%s · %s] %s" % (item[2], item[3], item[4])
prs.save(PPTX)
print("Notes ajoutées :", len(SCRIPT))

# --- 2) PDF imprimable ---
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (BaseDocTemplate, PageTemplate, Frame, Paragraph,
                                Spacer, Table, TableStyle)

NAVY=colors.HexColor("#0F1B2D"); BLUE=colors.HexColor("#1B3A5C")
ACCENT=colors.HexColor("#00B4C8"); GREEN=colors.HexColor("#2EC48F")
GREY=colors.HexColor("#5B6A7B"); DARK=colors.HexColor("#1B2633")
CARD=colors.HexColor("#F4F7FA"); AMBER=colors.HexColor("#C9870F")

ss=getSampleStyleSheet()
TXT=ParagraphStyle("TXT",parent=ss["Normal"],fontName="Helvetica",fontSize=10.5,leading=14.5,textColor=DARK,spaceAfter=2)
TIT=ParagraphStyle("TIT",parent=ss["Normal"],fontName="Helvetica-Bold",fontSize=11.5,leading=14,textColor=NAVY)
MET=ParagraphStyle("MET",parent=ss["Normal"],fontName="Helvetica-Bold",fontSize=8.5,leading=10,textColor=colors.white)
SEC=ParagraphStyle("SEC",parent=ss["Normal"],fontName="Helvetica-Bold",fontSize=8.5,leading=11,textColor=GREY)

W,H=A4; MX=1.6*cm; FW=W-2*MX
def hf(c,d):
    c.saveState()
    c.setFillColor(NAVY); c.rect(0,H-1.15*cm,W,1.15*cm,fill=1,stroke=0)
    c.setFillColor(ACCENT); c.rect(0,H-1.20*cm,W,0.06*cm,fill=1,stroke=0)
    c.setFillColor(colors.white); c.setFont("Helvetica-Bold",10.5)
    c.drawString(MX,H-0.78*cm,"APHÉLIE — Script de soutenance (texte à dire)")
    c.setFillColor(ACCENT); c.setFont("Helvetica-Bold",8.5)
    c.drawRightString(W-MX,H-0.78*cm,"~20 min · dont 5 min EN · Ilyesse KEBAILI")
    c.setFillColor(GREY); c.setFont("Helvetica",7.5)
    c.drawString(MX,0.7*cm,"Repères de minutage indicatifs — adaptez votre débit. Les annexes ne sont pas dans le temps imparti.")
    c.drawRightString(W-MX,0.7*cm,"Page %d"%c.getPageNumber())
    c.restoreState()

doc=BaseDocTemplate("/home/user/Soutenance/Script_Soutenance_Aphelie.pdf",pagesize=A4,
                    leftMargin=MX,rightMargin=MX,topMargin=1.5*cm,bottomMargin=1.05*cm)
fr=Frame(MX,1.05*cm,FW,H-1.5*cm-1.05*cm,leftPadding=0,rightPadding=0,topPadding=0,bottomPadding=0)
doc.addPageTemplates([PageTemplate(id='m',frames=[fr],onPage=hf)])

E=[]
def header_row(text):
    t=Table([[Paragraph('<font color="white"><b>%s</b></font>'%text,ss["Normal"])]],colWidths=[FW])
    t.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,-1),NAVY),("LEFTPADDING",(0,0),(-1,-1),8),
        ("TOPPADDING",(0,0),(-1,-1),4),("BOTTOMPADDING",(0,0),(-1,-1),4),("LINEBELOW",(0,0),(-1,-1),2,ACCENT)]))
    return t

E.append(header_row("SCRIPT ORAL — CE QUE VOUS DITES, DIAPOSITIVE PAR DIAPOSITIVE"))
E.append(Spacer(1,4))
E.append(Paragraph("Conseil : ne lisez pas mot à mot — mémorisez les idées et gardez ce texte comme filet de sécurité. "
    "Respirez aux retours à la ligne. La transition vers l’anglais (diapo 4) et le retour au français (diapo 6) sont "
    "indiqués dans le texte.", ParagraphStyle("i",parent=TXT,textColor=GREY,fontName="Helvetica-Oblique")))
E.append(Spacer(1,6))

last_section=None
for (n,title,dur,lang,text) in SCRIPT:
    if n==19:
        E.append(Spacer(1,4)); E.append(header_row("ANNEXES — DIAPOSITIVES DE RÉSERVE (hors temps, pour les questions)")); E.append(Spacer(1,4))
    badge_col = GREEN if lang=="EN" else ACCENT
    meta = Table([[Paragraph("Diapo %d"%n,MET), Paragraph(dur if dur else "",MET), Paragraph(lang,MET)]],
                 colWidths=[2.0*cm,2.4*cm,FW-2.0*cm-2.4*cm])
    meta.setStyle(TableStyle([("BACKGROUND",(0,0),(0,0),NAVY),("BACKGROUND",(1,0),(1,0),BLUE),
        ("BACKGROUND",(2,0),(2,0),badge_col),("LEFTPADDING",(0,0),(-1,-1),6),("RIGHTPADDING",(0,0),(-1,-1),6),
        ("TOPPADDING",(0,0),(-1,-1),2),("BOTTOMPADDING",(0,0),(-1,-1),2)]))
    head = Table([[meta, Paragraph(title,TIT)]],colWidths=[6.7*cm,FW-6.7*cm])
    head.setStyle(TableStyle([("VALIGN",(0,0),(-1,-1),"MIDDLE"),("LEFTPADDING",(1,0),(1,0),8),
        ("LEFTPADDING",(0,0),(0,0),0),("TOPPADDING",(0,0),(-1,-1),0),("BOTTOMPADDING",(0,0),(-1,-1),3)]))
    block=[head]
    for para in text.split("\n\n"):
        block.append(Paragraph(para.replace("\n"," "), TXT if not para.startswith("RÉSERVE") else
                               ParagraphStyle("r",parent=TXT,textColor=AMBER,fontName="Helvetica-Oblique")))
    block.append(Spacer(1,7))
    from reportlab.platypus import KeepTogether
    E.append(KeepTogether(block))

doc.build(E)
print("PDF script généré")
