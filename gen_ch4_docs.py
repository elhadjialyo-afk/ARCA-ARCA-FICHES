#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Génère les documents ISO 37001 Chapitre 4 + outils associés (.docx)."""
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

def new_doc(title, ref):
    d = Document()
    sec = d.sections[0]
    sec.page_width = Cm(21.0); sec.page_height = Cm(29.7)
    sec.top_margin = Cm(1.5); sec.bottom_margin = Cm(1.5)
    sec.left_margin = Cm(1.8); sec.right_margin = Cm(1.8)
    n = d.styles['Normal']; n.font.name='Calibri'; n.font.size=Pt(10)
    return d

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

def title_line(doc, text):
    add_para(doc, text, size=14, bold=True, color="1F3864", after=2)
    add_para(doc, "Référence : MRA_PDC_03.0X | Version 1.0 | 22 octobre 2026 | ISO 37001:2025 — chapitre 4",
             size=9, color="555555", after=6)

def add_para(doc, text, size=10, bold=False, italic=False, color=None, after=4, before=0, align=None):
    p = doc.add_paragraph(); p.paragraph_format.space_after = Pt(after); p.paragraph_format.space_before = Pt(before)
    if align: p.alignment = align
    r = p.add_run(text); r.font.size = Pt(size); r.bold = bold; r.italic = italic
    if color: r.font.color.rgb = RGBColor.from_string(color)
    return p

def add_table(d, headers, rows, widths, font_size=8.5):
    t = d.add_table(rows=1, cols=len(headers)); t.style='Table Grid'; t.alignment=WD_TABLE_ALIGNMENT.CENTER
    hdr = t.rows[0].cells
    for j,h in enumerate(headers):
        set_width(hdr[j], widths[j]); shade(hdr[j], "1F3864")
        p = hdr[j].paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(h); r.bold=True; r.font.size=Pt(font_size); r.font.color.rgb=RGBColor.from_string("FFFFFF")
        set_cell_margins(hdr[j], top=20, bottom=20); vcenter(hdr[j])
    repeat_header(t.rows[0])
    for i,row in enumerate(rows):
        cells = t.add_row().cells
        for j,val in enumerate(row):
            set_width(cells[j], widths[j]); set_cell_margins(cells[j])
            p = cells[j].paragraphs[0]; r = p.add_run(str(val)); r.font.size=Pt(font_size)
            if i%2==1: shade(cells[j], "F2F7FB")
            vcenter(cells[j])
    return t

# ===== DOC 1 — 4.1 Context of the organization =====
d1 = new_doc("4.1 Context of the organization", "MRA_PDC_03.01")
title_line(d1, "CONTEXTE DE L'ORGANISME — ISO 37001:2025 §4.1")
add_para(d1, "Définition du périmètre", size=11, bold=True, after=2)
add_table(d1,
    ["Élément", "Définition"],
    [
        ["Organisation", "AERIA — Gestionnaire de l'Aéroport International Félix Houphouët-Boigny d'Abidjan"],
        ["Système de management anti-corruption (ISMC)", "Conforme ISO 37001:2025, loi Sapin II (art. 17), recommandations AFA"],
        ["Zone géographique", "Aérodrome international Félix Houphouët-Boigny (Abidjan, CI) — siège + plateformes + tiers"],
        ["Périmètre documentaire", "Corpus 60–90 documents, chapitres par référentiel (ISO 37001, 9001, 14001, 45001, 55001, 50001, 22301, 37301 + OACI + ACA)"],
        ["Statut", "Brouillon — à valider avant mise en production"],
    ],
    [3.2, 18.3], font_size=8.5)

add_para(d1, "Contexte externe (enjeux déterminés)", size=11, bold=True, after=2)
add_table(d1,
    ["Domaine", "Enjeu", "Impact sur l'ISMC", "Crit."],
    [
        ["Légal", "Loi Sapin II art.17, Convention La Havane/ONU anticorruption", "Mandat déclaration, diligence, alerte", "Critique"],
        ["Certificaire", "AFA, ACA niveaux 1–5", "Cartographie, traçabilité carbone", "Critique"],
        ["Économique", "Trafic, concurrence, flux", "Pression sur contrats/appels d'offres", "Critique"],
        ["Culturel", "Cadeaux, pot de fin d'année normalisés", "Rareté remontée d'information", "Critique"],
        ["Tiers/intermédiaires", "Courtiers, sous-traitants, fournisseurs", "Diligence, contrats", "Critique"],
        ["Technique", "GED, outil carbone, ECCAIRS", "Traçabilité, intégrité des données", "Moyen"],
    ],
    [2.4,5.4,5.2,1.5], font_size=8)

add_para(d1, "Contexte interne (enjeux déterminés)", size=11, bold=True, after=2)
add_table(d1,
    ["Domaine", "Enjeu", "Impact", "Crit."],
    [
        ["Organisation/gouvernance", "Structure pyramidale, dirigeants, titulaires", "Dynamique décision, valorisation initiatives", "Critique"],
        ["Culture", "Cadeaux, respect autorité", "Normalisation comportements", "Critique"],
        ["Processus", "Procédures partielles, gestion tiers", "Traçabilité, double-sens", "Critique"],
        ["Ressources", "Équipes, formation", "Manque sensibl. anticorruption", "Moyen"],
        ["Outils", "GED, CRM, GRC, ECCAIRS", "Diffusion info sensible", "Moyen"],
    ],
    [2.8,5.2,5.0,1.5], font_size=8)

add_para(d1, "Parties intéressées et enjeux (4.2)", size=11, bold=True, after=2)
add_table(d1,
    ["Partie", "Attente", "Impact sur ISMC"],
    [
        ["Direction AERIA", "Certification AFA, gouvernance, transparence", "Direction du système, engagement"],
        ["Titulaires SMI / SMS / carbone", "Exigences opérationnelles quotidiennes", "Application locale des documents"],
        ["Auditeurs (AFA, jury panafrican)", "Conformité, règles de preuve, audit à blanc", "Définir la preuve exigée"],
        ["Vérificateurs ACA", "CHA, données tracées, vérification sans réserve", "Traçabilité des données carbone"],
        ["État / AAC", "Contrôle, attestation, respect du droit", "Référentiel juridique"],
        ["Clients, passagers, usagers", "Sécurité, image, confidentialité", "Contrôle des accès, non-divulgation"],
        ["Intermédiaires, prestataires", "Contrats, paiements", "Risque corruption, diligence"],
        ["Communautés locales", "Impact socio-éco (eau, sol, bruit)", "Responsabilité environnementale"],
    ],
    [4.5,5.5,6.5], font_size=8)

add_para(d1, "— Fin de la fiche §4.1 —", size=9, italic=True, color="808080", align=WD_ALIGN_PARAGRAPH.CENTER, before=6)

# ===== DOC 2 — 4.2 Needs and expectations of interested parties =====
d2 = new_doc("4.2 Needs and expectations of interested parties", "MRA_PDC_03.02")
title_line(d2, "ATTENTES DES PARTIES INTÉRESSÉES — ISO 37001:2025 §4.2")
add_table(d2,
    ["Partie intéressée", "Attente / enjeu", "Impact sur l'ISMC", "Réponse / preuve"],
    [
        ["Direction AERIA", "Certification AFA, gouvernance, transparence", "Direction du système, engagement exécutif", "Comité de direction, charte, revue"],
        ["Titulaires SMI / SMS / carbone", "Exigences opérationnelles quotidiennes", "Application locale des documents", "Procédures internes, fiches processus"],
        ["Auditeurs (AFA, jury panafrican)", "Conformité, règles de preuve, audit à blanc", "Définir la preuve exigée", "Matrice traçabilité, PV d'ateliers"],
        ["Vérificateurs ACA", "CHA, données tracées, vérification sans réserve", "Traçabilité des données carbone", "Dossier ACA, facteurs d'émission documentés"],
        ["État / AAC", "Contrôle, attestation, respect du droit", "Référentiel juridique", "Dossier d'audit, attestations"],
        ["Clients, passagers, usagers", "Sécurité, image, confidentialité", "Contrôle des accès, non-divulgation", "Charte, procédure accès"],
        ["Intermédiaires, prestataires", "Contrats, paiements, délais", "Diligence raisonnable, traçabilité paiements", "Registre tiers, devoirs de vigilance"],
        ["Communautés locales", "Impact socio-éco (eau, sol, bruit)", "Responsabilité environnementale", "Registre aspects/impacts"],
    ],
    [3.4,4.6,4.0,4.5], font_size=8)

add_para(d2, "Critères de conservation des attentes (4.2.1)", size=11, bold=True, after=2)
add_para(d2, "• Chaque partie intéressée et son attente est identifiée et conservée dans la matrice de traçabilité, avec la réponse retenue, la date de mise à jour et le responsable.", size=9, after=2)
add_para(d2, "• Les attentes contraignantes (légales, contractuelles) sont distinguées des attentes souhaitées (conventionnelles).", size=9, after=2)
add_para(d2, "• La matrice est revue après chaque audit, après toute modification de processus et au moins annuellement.", size=9, after=2)
add_para(d2, "— Fin de la fiche §4.2 —", size=9, italic=True, color="808080", align=WD_ALIGN_PARAGRAPH.CENTER, before=6)

# ===== DOC 3 — 4.3 Scope of the ISMC =====
d3 = new_doc("4.3 Scope of the ISMC", "MRA_PDC_03.03")
title_line(d3, "CHAMP D'APPLICATION DU SYSTÈME DE MANAGEMENT ANTI-CORRUPTION — ISO 37001:2025 §4.3")
add_table(d3,
    ["Élément", "Définition"],
    [
        ["Objet du SM", "Lutte contre la corruption — offres, pots-de-vin, fausses factures, conflits d'intérêts, déclenchement d'alerte, enquête."],
        ["Organisation concernée", "AERIA — siège + plateformes + tiers présents sur site (caissiers, agents accès, sécurité)."],
        ["Référentiels couverts", "ISO 37001:2025, ISO 37301, ISO 31000, ISO 19011, AFA, Loi Sapin II (art.17), OCDE, OACI Annexes 14/19, Doc 9859, Doc 9774, ACA (niveaux 1–5)."],
        ["Documents couverts", "ISS, SMS, carbone, qualité, conformité, SST, continuité, actifs, énergie."],
        ["Exclusions", "Conseil juridical opposé en justice (rôle humain), instruction des sanctions, certification de chiffre d'émissions, achat de crédits (rôles humains)."],
    ],
    [2.6,18.9], font_size=8.5)
add_para(d3, "Délimitation et justification (4.3.1)", size=11, bold=True, after=2)
add_para(d3, "Le système couvre l'ensemble des activités de gestion de l'aérodrome, car la corruption peut survenir à tous les niveaux : administration centrale, exploitation aéroportuaire, sécurité, caisses, études, maintenance, vente de terrains et services aux usagers. Aucune activité n'est exclue car chaque activité interagit avec les marchés publics, la sécurité et la relation avec les parties prenantes.", size=9, after=2)
add_para(d3, "Délimitation interne", size=11, bold=True, after=2)
add_para(d3, "• Activité exclue : conseil juridique opposé en justice (rôle humain), instruction des sanctions, certification indépendante des chiffres d'émission, achat de crédits carbone.", size=9, after=2)
add_para(d3, "• Activité partagée : l'agent ARCA prépare les dossiers, l'humain décide, valide, signe et fait appliquer.", size=9, after=2)
add_para(d3, "— Fin de la fiche §4.3 —", size=9, italic=True, color="808080", align=WD_ALIGN_PARAGRAPH.CENTER, before=6)

# ===== DOC 4 — 4.4 Anti-corruption management system =====
d4 = new_doc("4.4 Anti-corruption management system", "MRA_PDC_03.04")
title_line(d4, "SYSTÈME DE MANAGEMENT ANTI-CORRUPTION — ISO 37001:2025 §4.4")
add_para(d4, "Le SMAC est intégré au SM global d'AERIA. Il suit le cycle PDCA et est maintenu et amélioré continuellement.", size=9, after=4)
add_para(d4, "Cycle PDCA du SMAC", size=11, bold=True, after=2)
add_table(d4,
    ["Phase", "Activité", "Livrable", "Responsable"],
    [
        ["PLAN — Cadrage", "Périmètre, enjeux, parties prenantes, référentiels", "Note de cadrage (MRA_PDC_03.01)", "Directeur Général"],
        ["PLAN — Stratégie", "Architecture, risques, dispositions de contrôle", "Architecture doc + carte de risques", "Architecture + Chef CR&GR"],
        ["PLAN — Ressources", "Personnel, outils, formation, budget", "Plan de resources, annuel", "Stratégie"],
        ["DO — Acter", "Rédaction, ateliers, mise en œuvre", "Corpus brouillon, PV ateliers", "Agents + services"],
        ["CHECK — Évaluer", "Audit à blanc, vérification carbone, revue direction", "Rapport audit + plan d'action", "Vérificateur + Direction"],
        ["ACT — Améliorer", "Actions correctives, révision, actualisation", "Mises à jour, notifications", "Responsable QSE"],
    ],
    [2.2,5.6,5.6,3.1], font_size=8)
add_para(d4, "Structure du système (4.4.1)", size=11, bold=True, after=2)
add_table(d4,
    ["Niveau", "Élément", "Responsable", "Preuve de validation"],
    [
        ["Direction", "Comité de direction SMAC, charte, revue de direction", "Directeur Général", "PV réunion, présentation board"],
        ["Pilotes", "Responsable QSE, chefs départements, propriétaires doc", "Chefs départements", "Note de service, PV"],
        ["Mise en œuvre", "Agents, services, process owners", "Chef de processus", "Enregistrements, habilitations"],
        ["Vérification", "Audit interne, surveillance", "Auditeur interne", "Rapport audit, plan actions"],
        ["Amélioration", "Actions correctives, révisions", "Responsable QSE", "Plan d'actions clos"],
    ],
    [2.4,5.4,4.4,4.4], font_size=8)
add_para(d4, "Gouvernance et rôles (4.4.2)", size=11, bold=True, after=2)
add_table(d4,
    ["Rôle", "Personne", "Mission"],
    [
        ["Auteur", "Agent ARCA (LLM pilôté)", "Rédiger en brouillon, générer propositions, traçabilité"],
        ["Propriétaire", "Fonction anticorruption", "Décision, validation, pilotage du corpus"],
        ["Vérificateur", "Chef département Conformité & Gestion des Risques", "Contrôle, validation, recette"],
        ["Approbateur", "Directeur Général", "Décision finale, signature, diffusion"],
    ],
    [2.2,5.2,9.2], font_size=8)
add_para(d4, "— Fin de la fiche §4.4 —", size=9, italic=True, color="808080", align=WD_ALIGN_PARAGRAPH.CENTER, before=6)

# ===== DOC 5 — Other tool norms (glossaire + référentiel) =====
d5 = new_doc("Other tool norms (Glossaire + référentiel)", "MRA_PDC_03.05")
title_line(d5, "OUTILS ET GLOBIRE DÉFINITIFS — OUTILS SUPPORTEURS ISO 37001")
add_para(d5, "Glossaire des termes clés (4.4.2 / rapport de référence)", size=11, bold=True, after=2)
add_table(d5,
    ["Terme", "Définition (ISO 37001:2025)"],
    [
        ["Anticorruption", "Comportement contraire à l'intégrité et aux règles d'une organisation ou d'un agent dans les relations avec tout autre agent ou organisme."],
        ["Agent", "Personne physique ou morale qui est en relation avec une autre organisation ou une autre personne."],
        ["Attesteur", "Personne qui déclare ou atteste une affaire, dans le cadre d'une procédure ou d'une enquête."],
        ["Convention de la La Havane", "Traité visant à prévenir, détecter et réprimer la corruption dans la publicité internationale."],
        ["Facteur d'émission", "Quantité de gaz à effet de serre émise, traitée comme un produit ou un service."],
        ["Soumission", "Offre ou proposition faite par un agent pour réaliser un travail ou un projet."],
        ["Intégrité", "Principe de transparence et de conformité aux règles."],
        ["Interstate", "Relation entre organisations ou personnes ayant des intérêts."],
        ["Risque de non-conformité", "Probabilité qu'une obligation ne soit pas respectée."],
        ["Traçabilité", "Capacité à retrouver la provenance et l'historique d'une donnée ou d'un document."],
    ],
    [3.4,18.1], font_size=8)
add_para(d5, "Référentiels d'outils (tool norms) mobilisés", size=11, bold=True, after=2)
add_table(d5,
    ["Norme / programme", "Rôle dans le SMAC"],
    [
        ["ISO 37001:2025", "Norme de base — ISMC"],
        ["ISO 37301", "Norme de conformité — attireté le management"],
        ["ISO 9001 / 14001 / 45001 / 55001 / 50001 / 22301", "Systèmes de management intégrés"],
        ["ISO 19011, ISO 31000", "Audit et gestion des risques"],
        ["AFA recommandations", "Certificat anticorruption"],
        ["ACA niveaux 1–5", "Airport Carbon Accreditation"],
        ["OACI Annexes 14/19, Docs 9859, 9774", "Cadre aéronautique"],
        ["Loi Sapin II (art. 17) / ONU", "Cadre juridique international"],
    ],
    [4.5,17.0], font_size=8)
add_para(d5, "— Fin de la fiche outils —", size=9, italic=True, color="808080", align=WD_ALIGN_PARAGRAPH.CENTER, before=6)

OUT_DIR = "C:/Users/Administrateur/AppData/Local/hermes/hermes-agent/workspace/ARCA"
d1.save(OUT_DIR + "/MRI_03.01_ANALYSE_CONTEXTE_ISO37001.docx")
d2.save(OUT_DIR + "/MRI_03.02_ATTENTES_PARTIES.docx")
d3.save(OUT_DIR + "/MRI_03.03_CHAMP_APPLICATION.docx")
d4.save(OUT_DIR + "/MRI_03.04_ISMC.docx")
d5.save(OUT_DIR + "/MRI_03.05_GLOSSAIRE_REFERENCE.docx")
print("All generated.")
