#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Génère l'analyse de contexte AERIA (clause 4.1 ISO 37001) en .docx."""
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

OUT = "C:/Users/Administrateur/AppData/Local/hermes/hermes-agent/workspace/ARCA/MRI_03.01_ANALYSE_CONTEXTE_ISO37001.docx"

doc = Document()
sec = doc.sections[0]
sec.orientation = WD_ORIENT.PORTRAIT
sec.page_width = Cm(21.0); sec.page_height = Cm(29.7)
sec.top_margin = Cm(1.5); sec.bottom_margin = Cm(1.5)
sec.left_margin = Cm(1.8); sec.right_margin = Cm(1.8)

normal = doc.styles['Normal']
normal.font.name = 'Calibri'; normal.font.size = Pt(10)

def shade(cell, hex_color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd'); shd.set(qn('w:val'),'clear'); shd.set(qn('w:color'),'auto'); shd.set(qn('w:fill'),hex_color)
    tcPr.append(shd)

def set_width(cell, cm):
    cell.width = Cm(cm)

def set_cell_margins(cell, top=40, bottom=40, left=80, right=80):
    tcPr = cell._tc.get_or_add_tcPr(); m = OxmlElement('w:tcMar')
    for tag,val in (('top',top),('bottom',bottom),('start',left),('end',right)):
        e = OxmlElement(f'w:{tag}'); e.set(qn('w:w'),str(val)); e.set(qn('w:type'),'dxa'); m.append(e)
    tcPr.append(m)

def vcenter(cell):
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

def repeat_header(row):
    trPr = row._tr.get_or_add_trPr(); th = OxmlElement('w:tblHeader'); th.set(qn('w:val'),'true'); trPr.append(th)

def add_para(text, size=10, bold=False, italic=False, color=None, after=4, before=0, align=None):
    p = doc.add_paragraph(); p.paragraph_format.space_after = Pt(after); p.paragraph_format.space_before = Pt(before)
    if align: p.alignment = align
    r = p.add_run(text); r.font.size = Pt(size); r.bold = bold; r.italic = italic
    if color: r.font.color.rgb = RGBColor.from_string(color)
    return p

def add_table(headers, rows, widths, font_size=8.5, header_color="1F3864", bold_first=False):
    t = doc.add_table(rows=1, cols=len(headers)); t.style = 'Table Grid'; t.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr = t.rows[0].cells
    for j,h in enumerate(headers):
        set_width(hdr[j], widths[j]); shade(hdr[j], header_color)
        p = hdr[j].paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(h); r.bold = True; r.font.size = Pt(font_size); r.font.color.rgb = RGBColor.from_string("FFFFFF")
        set_cell_margins(hdr[j], top=20, bottom=20); vcenter(hdr[j])
    repeat_header(t.rows[0])
    for i,row in enumerate(rows):
        cells = t.add_row().cells
        for j,val in enumerate(row):
            set_width(cells[j], widths[j]); set_cell_margins(cells[j])
            p = cells[j].paragraphs[0]
            r = p.add_run(str(val)); r.font.size = Pt(font_size)
            if bold_first and j==0: r.bold = True
            if i%2==1: shade(cells[j], "F2F7FB")
            vcenter(cells[j])
    return t

# ============================ TITRE ============================
add_para("MRI 03.01 — ANALYSE DE CONTEXTE DU SYSTÈME DE MANAGEMENT ANTI-CORRUPTION", size=15, bold=True, color="1F3864", after=2)
add_para("Référence : MRA_PDC_03.01 | Version 1.0 | 22 octobre 2026", size=10, color="555555", after=2)
add_para("Propriétaire : Fonction anticorruption | Auteur : Responsable QSE | Vérificateur : Chef département Conformité & Gestion des Risques | Approbateur : Directeur Général",
         size=8.5, color="7c8899", after=8)

# ============================ 1. DÉFINITION DU PÉRIMÈTRE ============================
add_para("1. Définition du périmètre", size=12, bold=True, color="1F3864", after=4)
add_table(
    ["Élément", "Définition"],
    [
        ("Organisation", "AERIA — Gestionnaire de l'Aéroport International Félix Houphouët-Boigny d'Abidjan"),
        ("Système de management anti-corruption (ISMC)", "Système dédié à la lutte contre la corruption, conforme à ISO 37001:2025, la loi Sapin II (art. 17) et les recommandations AFA"),
        ("Zone géographique", "Aérodrome international Félix Houphouët-Boigny (Abidjan, Côte d'Ivoire) — siège + plateformes + tiers présents sur site"),
        ("Périmètre documentaire", "Corpus documentaire 60 à 90 documents, chapitres par référentiel (ISO 37001:2025, ISO 9001, ISO 14001, ISO 45001, ISO 55001, ISO 50001, ISO 22301, ISO 37301 + OACI + ACA)"),
        ("Statut", "Brouillon — à valider avant mise en production"),
    ],
    [3.6, 17.9], font_size=8.5
)

# ============================ 2. CONTEXTE EXTERNE ============================
add_para("2. Contexte externe — enjeux déterminés", size=12, bold=True, color="1F3864", after=4)
add_para("2.1 Environnement juridique et réglementaire", size=11, bold=True, after=2)
add_table(
    ["Domaine", "Enjeu externe", "Impact sur l'ISMC", "Criticité"],
    [
        ("Droit anticorruption", "Loi n° 2016-1691 du 9 déc. 2016 (art. 17) ; Convention de la Convention de La Havane/ONU sur la corruption internationale", "Mandat de déclaration, diligence raisonnable, procédure d'alerte", "Critique — cadre légal"),
        ("Certification AFA", "Recommandations AFA — 3 piliers : instance dirigeante, cartographie des risques, gestion des risques", "Cartographie des risques et dispositif de signalement", "Critique — certificat"),
        ("Airport Carbon Accreditation", "Programme ACA niveaux 1 à 5 — intégrité des données carbone", "Traçabilité des facteurs d'émission, non-soumission", "High — vérification"),
        ("Réglementaire aéronautique", "OACI Annexe 14, 19, Doc 9859 ; Règlement (UE) 139/2014 ; Doc 9774", "SMS & sécurité, processus d'exploitation", "Moyen — conformité"),
        ("Édition des normes", "ISO 9001:2026, ISO 14001:2026 (vigueur) ; ISO 45001 attend 2027", "Échéances de transition, mise à jour des documents", "Moyen — planification"),
    ],
    [2.6, 5.8, 4.6, 2.5], font_size=8
)

add_para("2.2 Environnement économique et concurrentiel", size=11, bold=True, after=2)
add_table(
    ["Domaine", "Enjeu externe", "Impact sur l'ISMC", "Criticité"],
    [
        ("Activité aéroportuaire", "Trafic passagers, fret, développement des liaisons", "Dépendance aux flux, pression sur les contrats et les appels d'offres", "Critique — économique"),
        ("Concurrence", "Aéroports voisins (Lomé, Accra, Abidjan) et nouveaux gestionnaires", "Tentatives de prise de contrats, influence sur les appels d'offres", "Moyen — concurrentiel"),
        ("Tiers et intermédiaires", "Agents commerciaux, courtiers, sous-traitants, fournisseurs", "Risque de corruption dans les négociations et paiements", "Critique — risque de mise en danger"),
        ("Financement", "Subventions, emprunts, partenariats public-privé", "Pressions, intérêts personnels sur les décisions d'achat", "Moyen"),
    ],
    [2.6, 5.6, 5.0, 2.3], font_size=8
)

add_para("2.3 Environnement social et culturel", size=11, bold=True, after=2)
add_table(
    ["Domaine", "Enjeu externe", "Impact sur l'ISMC", "Criticité"],
    [
        ("Culture d'entreprise", "Pratiques traditionnelles de cadeaux, pots de fin d'année, dons", "Normalisation de pratiques à contrôle", "Critique — culture"),
        ("Effectif et hiérarchie", "Structure pyramidale, respect de l'autorité", "Rareté de remontée d'information, culture de la discrétion", "Critique — culture"),
        ("Données personnelles", "RGPD-style, loi Ivoirienne sur la protection des données", "Traitement des dossiers de signalement (anonymat requis)", "Moyen — données"),
        ("Clients et passagers", "Interactions quotidiennes, paiements, accès", "Risque de chantage, extorsion ou corruption", "Moyen"),
    ],
    [2.6, 5.4, 5.0, 2.5], font_size=8
)

add_para("2.4 Environnement technologique", size=11, bold=True, after=2)
add_table(
    ["Domaine", "Enjeu externe", "Impact sur l'ISMC", "Criticité"],
    [
        ("GED / GRC / CMMS", "Outils d'exploitation, gestion des accès, registres", "Intégration des preuves, traçabilité des décisions", "Moyen"),
        ("Outil carbone / LMS", "Calcul GHG, plateforme de formation", "Intégrité des données, suivi formation", "Moyen"),
        ("Sécurité de l'information", "Protection des données sensibles, alertes", "Accès non autorisés aux dossiers anticorruption", "Moyen"),
    ],
    [2.6, 5.4, 5.0, 2.5], font_size=8
)

# ============================ 3. CONTEXTE INTERNE ============================
add_para("3. Contexte interne — enjeux déterminés", size=12, bold=True, color="1F3864", after=4)
add_table(
    ["Domaine", "Enjeu interne", "Impact sur l'ISMC", "Criticité"],
    [
        ("Organisation et gouvernance", "Structure pyramidale, dirigeants, titulaires des postes", "Dynamique de décision, valorisation des initiatives", "Critique"),
        ("Culture d'entreprise", "Pratiques de cadeaux, respect de l'autorité, usage de l'information", "Normalisation de comportements à contrôler", "Critique"),
        ("Processus existants", "Procédures anti-corruption partielles, gestion des intermédiaires", "Couacs de traçabilité, double-sens documentaire", "Critique"),
        ("Ressources et compétences", "Équipes SMI/SMS/carbone, formation", "Manque de sensibilisation anticorruption", "Moyen"),
        ("Outils d'information", "GED, CRM, GRC, ECCAIRS", "Diffusion non maîtrisée d'informations sensibles", "Moyen"),
        ("Audit et revue", "Revue de direction, audit interne", "Détection tardive des écarts", "Moyen"),
    ],
    [2.6, 5.4, 5.0, 2.5], font_size=8
)

# ============================ 4. PARTIES INTÉRESSÉES ============================
add_para("4. Parties intéressées (stakeholders)", size=12, bold=True, color="1F3864", after=4)
add_table(
    ["Partie intéressée", "Attente / enjeu", "Impact sur l'ISMC"],
    [
        ("Direction AERIA", "Certification AFA, validation", "Décision, autorisation, engagement"),
        ("Titulaires des postes (SMI, SMS, carbone)", "Exigences opérationnelles", "Application locale des documents"),
        ("Auditeurs (certificateurs AFA, jury panafrican)", "Conformité, règles de preuve", "Délivrance, audit à blanc"),
        ("Vérificateurs carbone", "CHA et données tracées", "Données fiables, traçables"),
        ("État / Autorité de l'aviation civile", "Contrôle, attestation", "Référentiel juridique"),
        ("Clients, passagers, usagers", "Sécurité, transparence", "Confiance, image"),
        ("Équipage, personnel de sécurité, agents de sécurité", "Sécurité, accès", "Contrôle des accès, signalements"),
        ("Intermédiaires et prestataires", "Contrats, paiements", "Risque de corruption, diligence raisonnable"),
        ("Communautés locales", "Impact socio-éco (eau, sol, bruit)", "Responsabilité environnementale"),
    ],
    [4.5, 6.0, 6.0], font_size=8
)

# ============================ 5. ÉNJEUX DÉTERMINÉS ============================
add_para("5. Enjeux déterminés et intégraux de l'ISMC", size=12, bold=True, color="1F3864", after=4)
add_para("Les enjeux ci-dessous sont représentatifs et sont pérennisés par la matrice de traçabilité (MRA_PDC_04.02).",
         size=8.5, italic=True, color="7c8899", after=4)
add_table(
    ["#", "Enjeu", "Nature", "Entrée ISO 37001", "Impact sur l'ISMC", "Criticité"],
    [
        ("E1", "Loi Sapin II / Convention de La Havane", "Externe — légal", "4.1.1", "Mandat de déclaration, diligence, alerte", "Critique"),
        ("E2", "Certification AFA et programme ACA", "Externe — certificat", "4.1.1", "Cartographie, preuves carbone", "Critique"),
        ("E3", "Trafic aéroportuaire et dépendance aux flux", "Externe — économique", "4.1.1", "Contrats, appels d'offres, pression", "Critique"),
        ("E4", "Culture d'entreprise et normalisations", "Interne — culturel", "4.1.1", "Rareté de la remontée d'information", "Critique"),
        ("E5", "Tiers et intermédiaires", "Interne — relationnel", "4.1.1", "Diligence raisonnable, contrats", "Critique"),
        ("E6", "Rareté de remontée d'information", "Interne — organisation", "4.1.1", "Canal de signalement confidentiel", "Critique"),
        ("E7", "Édition des normes et transition", "Externe — technique", "4.1.1", "Mise à jour des documents", "Moyen"),
        ("E8", "Sécurité de l'information", "Interne — technique", "4.1.1", "Protection des données sensibles", "Moyen"),
        ("E9", "Gouvernance et dirigeants", "Interne — organisation", "4.1.1", "Exemple, engagement, validation", "Moyen"),
    ],
    [0.8, 3.8, 2.4, 1.6, 5.0, 1.6], font_size=8
)

# ============================ 6. RÉSULTAT CONSOLIDÉ ============================
add_para("6. Résultat consolidé de l'analyse de contexte", size=12, bold=True, color="1F3864", after=4)
add_para("L'ISMC d'AERIA s'articule autour de trois axes prioritaires :", size=9, after=2)
add_para("• Axe 1 — Cadre légal et certificaire (E1, E2) : anticiper les exigences légales et les règles de la certification pour rester conforme.", size=9, after=1)
add_para("• Axe 2 — Culture et comportement (E4, E5, E6) : transformer la culture d'entreprise et créer des canaux sûrs de remontée d'information.", size=9, after=1)
add_para("• Axe 3 — Prévention des risques (E3, E7, E8, E9) : maîtriser la dépendance aux flux, les transitions normatives et la sécurité de l'information.", size=9, after=2)

add_para("Les enjeux prioritaires pour la première année de mise en œuvre :", size=9, bold=True, after=2)
add_para("1. Cartographie des risques anticorruption (E1, E2, E5) — base de la matrice de traçabilité.", size=9, after=1)
add_para("2. Procédure de signalement et canal confidentiel (E6).", size=9, after=1)
add_para("3. Code de déontologie et règles de cadeaux (E4).", size=9, after=1)
add_para("4. Programme de sensibilisation et formation (E4, E6).", size=9, after=1)
add_para("5. Diligence raisonnable sur les tiers (E5).", size=9, after=1)
add_para("6. Plan de transition normes (E7) et sécurité de l'information (E8).", size=9, after=1)

add_para("— Fin de l'analyse de contexte —", size=9, italic=True, color="808080", align=WD_ALIGN_PARAGRAPH.CENTER, before=8)

doc.save(OUT)
print("Saved:", OUT)
