#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Génère la matrice de traçabilité ARCA en .docx."""
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

OUT = "C:/Users/Administrateur/AppData/Local/hermes/hermes-agent/workspace/ARCA/MATRICE_TRAÇABILITÉ_ARCA.docx"

doc = Document()

# ---------- Page paysage ----------
sec = doc.sections[0]
sec.orientation = WD_ORIENT.LANDSCAPE
sec.page_width = Cm(29.7)
sec.page_height = Cm(21.0)
sec.top_margin = Cm(1.5)
sec.bottom_margin = Cm(1.5)
sec.left_margin = Cm(1.5)
sec.right_margin = Cm(1.5)

normal = doc.styles['Normal']
normal.font.name = 'Calibri'
normal.font.size = Pt(10)

def shade(cell, hex_color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    tcPr.append(shd)

def set_width(cell, cm):
    cell.width = Cm(cm)

def set_cell_margins(cell, top=40, bottom=40, left=60, right=60):
    tcPr = cell._tc.get_or_add_tcPr()
    m = OxmlElement('w:tcMar')
    for tag, val in (('top', top), ('bottom', bottom), ('start', left), ('end', right)):
        e = OxmlElement(f'w:{tag}')
        e.set(qn('w:w'), str(val))
        e.set(qn('w:type'), 'dxa')
        m.append(e)
    tcPr.append(m)

def vcenter(cell):
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

def repeat_header(row):
    trPr = row._tr.get_or_add_trPr()
    th = OxmlElement('w:tblHeader')
    th.set(qn('w:val'), 'true')
    trPr.append(th)

def add_para(text, size=10, bold=False, italic=False, color=None, after=4, before=0, align=None):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(after)
    p.paragraph_format.space_before = Pt(before)
    if align:
        p.alignment = align
    r = p.add_run(text)
    r.font.size = Pt(size)
    r.bold = bold
    r.italic = italic
    if color:
        r.font.color.rgb = RGBColor.from_string(color)
    return p

# ---------- Titre ----------
add_para("MATRICE DE TRAÇABILITÉ", size=20, bold=True, color="1F3864", after=2)
add_para("Agent Anticorruption (ARCA) — ISO 37001:2025 / loi Sapin II (art. 17)",
         size=12, bold=False, color="2E5496", after=8)

# ---------- Bloc d'identification ----------
info = [
    ("Projet", "Agent de Référentiel et de Construction Anti-corruption (ARCA)"),
    ("Document", "Matrice de traçabilité exigence → document → preuve"),
    ("Numérotation", "MRA_PDC_XX (documents du projet anticorruption)"),
    ("Version / Date", "0.1 — 2026-10-20"),
    ("Statut", "Brouillon (à valider documentaire)"),
    ("Propriétaire", "Fonction anticorruption"),
    ("Auteur / Responsable", "Responsable QSE"),
    ("Vérificateur", "Chef département Conformité & Gestion des Risques"),
    ("Approbateur", "Directeur Général"),
    ("Références", "ISO 37001:2025 ; Loi n° 2016-1691 du 9 décembre 2016 (art. 17) ; AFA"),
]
t = doc.add_table(rows=len(info), cols=2)
t.style = 'Table Grid'
t.alignment = WD_TABLE_ALIGNMENT.CENTER
for i, (k, v) in enumerate(info):
    c0, c1 = t.rows[i].cells
    c0.width = Cm(4.5)
    c1.width = Cm(22.0)
    shade(c0, "D9E2F3")
    set_cell_margins(c0); set_cell_margins(c1)
    p0 = c0.paragraphs[0]; r0 = p0.add_run(k); r0.bold = True; r0.font.size = Pt(9.5)
    p1 = c1.paragraphs[0]; r1 = p1.add_run(v); r1.font.size = Pt(9.5)
    for c in (c0, c1):
        vcenter(c)

add_para("", after=2)

# ---------- Glossaire / numérotation ----------
add_para("Glossaire des documents du projet (numérotation MRA_PDC_XX)", size=11, bold=True, color="1F3864", after=4)
docs = [
    ("MRA_PDC_010_01", "Politique anticorruption", "Niveau A"),
    ("MRA_PDC_010_02", "Code de déontologie / règles de conduite", "Niveau A"),
    ("MRA_PDC_010_03", "Charte des cadeaux et invitations", "Niveau A"),
    ("MRA_PDC_020_01", "Cartographie et évaluation des risques anticorruption", "Niveau B"),
    ("MRA_PDC_020_02", "Règles de diligence raisonnable (tiers, contrats, paiements)", "Niveau B"),
    ("MRA_PDC_020_03", "Procédures d'attribution des marchés publics", "Niveau B"),
    ("MRA_PDC_020_04", "Règles de signalement / lanceurs d'alerte", "Niveau B"),
    ("MRA_PDC_020_05", "Programme de formation et de sensibilisation", "Niveau B"),
    ("MRA_PDC_020_06", "Comité de suivi anticorruption", "Niveau B"),
    ("MRA_PDC_030_01", "Registre des déclarations d'intérêts", "Niveau C"),
    ("MRA_PDC_030_02", "Attestations de déontologie des parties prenantes", "Niveau C"),
    ("MRA_PDC_030_03", "Journaux de formation et de sensibilisation", "Niveau C"),
    ("MRA_PDC_030_04", "Comptes rendus de due diligence (tiers)", "Niveau C"),
    ("MRA_PDC_030_05", "Comptes rendus d'audits / comité de suivi", "Niveau C"),
    ("MRA_PDC_040_01", "Matrice de traçabilité anticorruption (ce document)", "Niveau D"),
    ("MRA_PDC_040_02", "Grilles de cotation et de notation des risques", "Niveau D"),
    ("MRA_PDC_040_03", "Tableau de bord anticorruption", "Niveau D"),
]
t2 = doc.add_table(rows=len(docs), cols=3)
t2.style = 'Table Grid'
t2.alignment = WD_TABLE_ALIGNMENT.CENTER
widths2 = [Cm(3.6), Cm(12.5), Cm(3.6)]
for i, (id_, nom, niv) in enumerate(docs):
    c0, c1, c2 = t2.rows[i].cells
    c0.width = widths2[0]; c1.width = widths2[1]; c2.width = widths2[2]
    set_cell_margins(c0); set_cell_margins(c1); set_cell_margins(c2)
    r0 = c0.paragraphs[0].add_run(id_); r0.font.size = Pt(9)
    r1 = c1.paragraphs[0].add_run(nom); r1.font.size = Pt(9)
    r2 = c2.paragraphs[0].add_run(niv); r2.font.size = Pt(9)
    if i % 2 == 1:
        shade(c0, "F2F2F2"); shade(c1, "F2F2F2"); shade(c2, "F2F2F2")
    for c in (c0, c1, c2):
        vcenter(c)

add_para("", after=2)

# ---------- Matrice principale ----------
add_para("Matrice principale — Exigences → Documents → Preuves", size=11, bold=True, color="1F3864", after=4)

headers = ["ID exigence", "Source", "Exigence", "Document", "Preuve d'application",
           "Auteur", "Vérificateur", "Approbateur", "Date", "Statut"]

rows_data = [
    # --- ISO 37001:2025, clause 4 ---
    ("EQ-001", "ISO 37001:2025 — cl. 4.1", "La direction détermine le contexte interne et externe affectant l'ISMC anticorruption.",
     "MRA_PDC_020_01", "Évaluation des risques / cartographie des risques", "Resp. QSE", "Chef CR&GR", "DG", "2026-11", "Brouillon"),
    ("EQ-002", "ISO 37001:2025 — cl. 4.2", "Les parties intéressées anticorruption sont identifiées et analysées.",
     "MRA_PDC_020_01", "Registre des parties intéressées", "Resp. QSE", "Chef CR&GR", "DG", "2026-11", "Brouillon"),
    # --- clause 5 ---
    ("EQ-003", "ISO 37001:2025 — cl. 5.2", "Politique anticorruption établie, publiée et communiquée à tous les employés.",
     "MRA_PDC_010_01", "Politique anticorruption signée + diffusée", "Resp. QSE", "Chef CR&GR", "DG", "2026-11", "Brouillon"),
    ("EQ-004", "ISO 37001:2025 — cl. 5.3", "Rôles, responsabilités et autorisations de la direction pour lutter contre la corruption.",
     "MRA_PDC_010_02", "Code de déontologie + organigramme d'attribution", "Resp. QSE", "Chef CR&GR", "DG", "2026-11", "Brouillon"),
    # --- clause 6 ---
    ("EQ-005", "ISO 37001:2025 — cl. 6.1", "Risques anticorruption identifiés et évalués (période de révision définie).",
     "MRA_PDC_020_01", "Rapport d'évaluation des risques + grille de cotation", "Resp. QSE", "Chef CR&GR", "DG", "2026-11", "Brouillon"),
    ("EQ-006", "ISO 37001:2025 — cl. 6.1", "Actions pour gérer les risques anticorruption identifiés.",
     "MRA_PDC_020_02 / 03 / 04", "Procédures de diligence, marchés, signalement", "Resp. QSE", "Chef CR&GR", "DG", "2026-12", "Brouillon"),
    # --- clause 7 ---
    ("EQ-007", "ISO 37001:2025 — cl. 7.2", "Compétences et formation du personnel adaptées au risque anticorruption.",
     "MRA_PDC_020_05", "Programme de formation + journaux de formation", "Resp. QSE", "Chef CR&GR", "DG", "2026-12", "Brouillon"),
    ("EQ-008", "ISO 37001:2025 — cl. 7.3", "Conscience du milieu et communication de l'ISMC.",
     "MRA_PDC_010_02", "Code de déontologie communiquée + campagnes", "Resp. QSE", "Chef CR&GR", "DG", "2026-12", "Brouillon"),
    ("EQ-009", "ISO 37001:2025 — cl. 7.5", "Documents et enregistrements de l'ISMC établis, mis à jour et conservés.",
     "MRA_PDC_040_01", "Matrice de traçabilité maintenue à jour", "Resp. QSE", "Chef CR&GR", "DG", "2026-10", "Brouillon"),
    # --- clause 8 ---
    ("EQ-010", "ISO 37001:2025 — cl. 8.1", "Procédures d'attribution des marchés publics et de sous-traitance selon les règles anti-corruption.",
     "MRA_PDC_020_03", "Procédure d'attribution / bilan des marchés", "Resp. QSE", "Chef CR&GR", "DG", "2026-12", "Brouillon"),
    ("EQ-011", "ISO 37001:2025 — cl. 8.2", "Diligence raisonnable sur les fournisseurs, sous-traitants et agents commerciaux.",
     "MRA_PDC_020_02", "Registre des tiers + due diligences", "Resp. QSE", "Chef CR&GR", "DG", "2026-12", "Brouillon"),
    ("EQ-012", "ISO 37001:2025 — cl. 8.3", "Gestion des risques anticorruption dans les activités courantes (boucle d'actions).",
     "MRA_PDC_020_02", "Plan d'actions anticorruption avec propriétaires", "Resp. QSE", "Chef CR&GR", "DG", "2026-12", "Brouillon"),
    ("EQ-013", "ISO 37001:2025 — cl. 8.4", "Mécanisme de signalement des allégations de corruption (lanceurs d'alerte).",
     "MRA_PDC_020_04", "Registre de signalement / procédure lanceurs d'alerte", "Resp. QSE", "Chef CR&GR", "DG", "2026-12", "Brouillon"),
    ("EQ-014", "ISO 37001:2025 — cl. 8.5", "Procédures de négociation et d'activités commerciales intégrant le code de déontologie.",
     "MRA_PDC_010_02", "Charte de négociation + code de déontologie", "Resp. QSE", "Chef CR&GR", "DG", "2026-12", "Brouillon"),
    # --- clause 9 ---
    ("EQ-015", "ISO 37001:2025 — cl. 9.2", "Audit interne de l'ISMC anticorruption mené sur une période déterminée.",
     "MRA_PDC_030_05", "Compte rendu d'audit anticorruption", "Resp. QSE", "Chef CR&GR", "DG", "2027-01", "Brouillon"),
    ("EQ-016", "ISO 37001:2025 — cl. 9.3", "Évaluation de conformité / délivrance du système de management anticorruption.",
     "MRA_PDC_030_05", "Attestation de conformité / revue de direction", "Resp. QSE", "Chef CR&GR", "DG", "2027-01", "Brouillon"),
    # --- clause 10 ---
    ("EQ-017", "ISO 37001:2025 — cl. 10.2", "Actions correctives et amélioration continue de l'ISMC.",
     "MRA_PDC_040_03", "Tableau de bord anticorruption avec indicateurs", "Resp. QSE", "Chef CR&GR", "DG", "2027-02", "Brouillon"),
    # --- Loi Sapin II, art. 17 (8 mesures) ---
    ("SA-001", "Sapin II — art. 17 (1ère mesure)", "Évaluation des risques anticorruption and cartographie des risques d'attentat.",
     "MRA_PDC_020_01", "Rapport d'évaluation des risques", "Resp. QSE", "Chef CR&GR", "DG", "2026-11", "Brouillon"),
    ("SA-002", "Sapin II — art. 17 (2e mesure)", "Règles de diligence raisonnable pour les intermédiaires et tiers.",
     "MRA_PDC_020_02", "Mémoire de diligence raisonnable", "Resp. QSE", "Chef CR&GR", "DG", "2026-12", "Brouillon"),
    ("SA-003", "Sapin II — art. 17 (3e mesure)", "Procédures d'attribution des marchés publics respectant les exigences anticorruption.",
     "MRA_PDC_020_03", "Procédure d'attribution des marchés", "Resp. QSE", "Chef CR&GR", "DG", "2026-12", "Brouillon"),
    ("SA-004", "Sapin II — art. 17 (4e mesure)", "Déclaration des intérêts des parties prenantes.",
     "MRA_PDC_030_01", "Registre des déclarations d'intérêts", "Resp. QSE", "Chef CR&GR", "DG", "2026-12", "Brouillon"),
    ("SA-005", "Sapin II — art. 17 (5e mesure)", "Code de déontologie applicable à l'ensemble des parties prenantes.",
     "MRA_PDC_010_02", "Code de déontologie signé", "Resp. QSE", "Chef CR&GR", "DG", "2026-11", "Brouillon"),
    ("SA-006", "Sapin II — art. 17 (6e mesure)", "Programme de formation et de sensibilisation à l'anticorruption.",
     "MRA_PDC_020_05", "Programme de formation + attestations", "Resp. QSE", "Chef CR&GR", "DG", "2026-12", "Brouillon"),
    ("SA-007", "Sapin II — art. 17 (7e mesure)", "Comité de suivi de la lutte anticorruption.",
     "MRA_PDC_020_06", "Compte rendu de comité de suivi", "Resp. QSE", "Chef CR&GR", "DG", "2027-01", "Brouillon"),
    ("SA-008", "Sapin II — art. 17 (8e mesure)", "Programme de détection et de signalement / lanceurs d'alerte.",
     "MRA_PDC_020_04", "Procédure de signalement", "Resp. QSE", "Chef CR&GR", "DG", "2026-12", "Brouillon"),
]

n_cols = len(headers)
table = doc.add_table(rows=1, cols=n_cols)
table.style = 'Table Grid'
table.alignment = WD_TABLE_ALIGNMENT.CENTER

# header
hdr = table.rows[0].cells
for j, h in enumerate(headers):
    set_width(hdr[j], [1.6, 1.4, 6.8, 2.6, 4.4, 2.2, 2.4, 2.2, 1.3, 1.4][j])
    shade(hdr[j], "1F3864")
    p = hdr[j].paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(h); r.bold = True; r.font.size = Pt(8.5); r.font.color.rgb = RGBColor.from_string("FFFFFF")
    set_cell_margins(hdr[j], top=30, bottom=30)
    vcenter(hdr[j])
repeat_header(table.rows[0])

# data rows
for i, row in enumerate(rows_data):
    cells = table.add_row().cells
    for j, val in enumerate(row):
        set_width(cells[j], [1.6, 1.4, 6.8, 2.6, 4.4, 2.2, 2.4, 2.2, 1.3, 1.4][j])
        set_cell_margins(cells[j])
        p = cells[j].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT if j == 2 else WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(str(val)); r.font.size = Pt(8)
        if j == 0:
            r.bold = True
        if i % 2 == 1:
            shade(cells[j], "F2F7FB")
        vcenter(cells[j])

# ---------- Remarques / lecture ----------
add_para("", after=2)
add_para("Lecture et usage du document", size=11, bold=True, color="1F3864", after=4)
notes = [
    "Chaque exigence normatif (ISO 37001:2025 / art. 17 Sapin II) est rattachée à un document du système et à une preuve d'application.",
    "Les colonnes Auteur / Vérificateur / Approbateur définissent la chaîne de validation de chaque document.",
    "Statut 'Brouillon' = document en élaboration, à valider lors de la revue du corpus.",
    "Cette matrice est maintenue à jour dans le cadre de la boucle de production ARCA (cadrer → décomposer → concevoir → rédiger → prouver).",
]
for n in notes:
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.5)
    p.paragraph_format.space_after = Pt(3)
    r = p.add_run("• " + n); r.font.size = Pt(9.5)

add_para("", after=2)
add_para("— Fin de la matrice —", size=9, italic=True, color="808080", align=WD_ALIGN_PARAGRAPH.CENTER)

doc.save(OUT)
print("Saved:", OUT)
