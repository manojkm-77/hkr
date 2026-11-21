#!/usr/bin/env python3
"""
Comprehensive COPD Postgraduate Master Textbook Generator
Produces a publication-grade, ~100-page equivalent Microsoft Word (.docx) monograph
for MD & DNB Respiratory Medicine residents, incorporating the latest GOLD 2024/2025,
ERS, ATS, and Indian Chest Society guidelines with embedded medical diagrams.
"""

import os
import sys
import tempfile
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont
import docx
from docx import Document
from docx.shared import Inches, Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls

print("Initializing COPD Master Textbook Generation Pipeline...")

# Define output directory and file paths
BASE_DIR = Path(r"C:\Users\harshith k r\Desktop\textbooks")
SCRATCH_DIR = Path(r"C:\Users\harshith k r\.gemini\antigravity-cli\brain\b3489ea1-3baa-44d2-b30e-c7976216faed\scratch")
ASSETS_DIR = Path(tempfile.gettempdir())
OUTPUT_DOCX = BASE_DIR / "Chronic_Obstructive_Pulmonary_Disease_Comprehensive_Master_Textbook.docx"
SCRATCH_DOCX = SCRATCH_DIR / "Chronic_Obstructive_Pulmonary_Disease_Comprehensive_Master_Textbook.docx"

# Color Palette definition (Navy/Teal/Burgundy academic scheme)
COLOR_NAVY_HEX = "183B56"
COLOR_TEAL_HEX = "1D6A70"
COLOR_BURGUNDY_HEX = "82424D"
COLOR_CHARCOAL_HEX = "212529"
COLOR_MUTED_HEX = "52616B"
COLOR_LIGHT_BG_HEX = "F4F6F8"
COLOR_CALLOUT_BORDER_HEX = "1D6A70"

# ---------------------------------------------------------------------------
# PART 1: HIGH-RESOLUTION SCIENTIFIC DIAGRAM GENERATION (PIL)
# ---------------------------------------------------------------------------

def create_pathology_diagram(output_path):
    """Generates an illustrative comparison diagram of Centriacinar vs Panacinar Emphysema and Chronic Bronchitis."""
    img = Image.new("RGB", (1200, 600), "#FFFFFF")
    draw = ImageDraw.Draw(img)
    
    # Header banner
    draw.rectangle([0, 0, 1200, 70], fill=f"#{COLOR_NAVY_HEX}")
    draw.text((30, 20), "PATHOLOGY OF COPD: CHRONIC BRONCHITIS & EMPHYSEMA PHENOTYPES", fill="#FFFFFF")
    
    # Left Box: Chronic Bronchitis
    draw.rounded_rectangle([30, 90, 390, 560], radius=12, outline=f"#{COLOR_TEAL_HEX}", width=3, fill="#F8FAFC")
    draw.text((50, 110), "CHRONIC BRONCHITIS", fill=f"#{COLOR_BURGUNDY_HEX}")
    draw.text((50, 135), "(Small Airway Disease & Hypersecretion)", fill=f"#{COLOR_MUTED_HEX}")
    
    # Bronchial cross-section schematic
    draw.ellipse([110, 180, 310, 380], outline=f"#{COLOR_NAVY_HEX}", width=8, fill="#E2E8F0")
    draw.ellipse([140, 210, 280, 350], outline=f"#{COLOR_BURGUNDY_HEX}", width=12, fill="#FEF3C7") # Thickened wall
    draw.ellipse([180, 250, 240, 310], fill=f"#{COLOR_TEAL_HEX}") # Mucus plug in narrowed lumen
    
    cb_points = [
        "• Goblet cell & submucosal gland hyperplasia",
        "• Reid Index > 0.50 (gland / wall ratio)",
        "• Luminal mucus plugging with cellular debris",
        "• Peribronchiolar fibrosis & smooth muscle hypertrophy",
        "• Dense CD8+ T-cell & macrophage infiltrate",
        "• Predominantly affects airways < 2 mm diameter"
    ]
    y_text = 400
    for pt in cb_points:
        draw.text((45, y_text), pt, fill=f"#{COLOR_CHARCOAL_HEX}")
        y_text += 24

    # Center Box: Centriacinar Emphysema
    draw.rounded_rectangle([420, 90, 780, 560], radius=12, outline=f"#{COLOR_TEAL_HEX}", width=3, fill="#F8FAFC")
    draw.text((440, 110), "CENTRIACINAR EMPHYSEMA", fill=f"#{COLOR_NAVY_HEX}")
    draw.text((440, 135), "(Centrilobular / Tobacco-Smoking Induced)", fill=f"#{COLOR_MUTED_HEX}")
    
    # Schematic: Centrilobular destruction
    draw.ellipse([500, 180, 700, 380], outline="#CBD5E1", width=3, fill="#FFFFFF")
    # Central hole
    draw.ellipse([560, 240, 640, 320], outline=f"#{COLOR_BURGUNDY_HEX}", width=6, fill="#FEE2E2")
    draw.text((570, 275), "Central\nDefect", fill=f"#{COLOR_BURGUNDY_HEX}")
    
    ca_points = [
        "• Respiratory bronchioles selectively destroyed",
        "• Distal alveolar ducts & alveoli preserved",
        "• Marked upper-lobe & apical predominance",
        "• Driven by cigarette smoke & carbon deposits",
        "• Uneven ventilation-perfusion mismatch",
        "• Characteristic lesion of common smoker's COPD"
    ]
    y_text = 400
    for pt in ca_points:
        draw.text((435, y_text), pt, fill=f"#{COLOR_CHARCOAL_HEX}")
        y_text += 24

    # Right Box: Panacinar Emphysema
    draw.rounded_rectangle([810, 90, 1170, 560], radius=12, outline=f"#{COLOR_TEAL_HEX}", width=3, fill="#F8FAFC")
    draw.text((830, 110), "PANACINAR EMPHYSEMA", fill=f"#{COLOR_NAVY_HEX}")
    draw.text((830, 135), "(Panlobular / α1-Antitrypsin Deficiency)", fill=f"#{COLOR_MUTED_HEX}")
    
    # Schematic: Uniform destruction
    draw.ellipse([890, 180, 1090, 380], outline=f"#{COLOR_BURGUNDY_HEX}", width=6, fill="#FEE2E2")
    draw.text((945, 275), "Whole Acinus\nDestroyed", fill=f"#{COLOR_BURGUNDY_HEX}")
    
    pa_points = [
        "• Uniform enlargement from bronchiole to alveoli",
        "• Total loss of alveolar septa & surface area",
        "• Lower-lobe & anterior basilar predominance",
        "• Hallmark of α1-Antitrypsin Deficiency (Pi*ZZ)",
        "• Severe early reduction in DLCO",
        "• Loss of parenchymal radial traction on airways"
    ]
    y_text = 400
    for pt in pa_points:
        draw.text((825, y_text), pt, fill=f"#{COLOR_CHARCOAL_HEX}")
        y_text += 24

    img.save(output_path, "PNG", dpi=(300, 300))
    print(f"  [+] Generated: {output_path.name}")

def create_epp_diagram(output_path):
    """Generates a diagram explaining Equal Pressure Point (EPP) and dynamic airway compression."""
    img = Image.new("RGB", (1200, 550), "#FFFFFF")
    draw = ImageDraw.Draw(img)
    
    draw.rectangle([0, 0, 1200, 65], fill=f"#{COLOR_NAVY_HEX}")
    draw.text((30, 20), "PULMONARY MECHANICS: EQUAL PRESSURE POINT (EPP) & DYNAMIC AIRWAY COLLAPSE", fill="#FFFFFF")
    
    # Normal Lung Panel
    draw.rounded_rectangle([30, 85, 580, 520], radius=10, outline=f"#{COLOR_TEAL_HEX}", width=2, fill="#F8FAFC")
    draw.text((50, 105), "A. NORMAL FORCED EXPIRATION", fill=f"#{COLOR_NAVY_HEX}")
    draw.text((50, 130), "Alveolar recoil maintains positive tethering pressure.", fill=f"#{COLOR_MUTED_HEX}")
    
    # Normal airway schematic
    draw.rectangle([100, 200, 520, 270], fill="#E2E8F0", outline=f"#{COLOR_NAVY_HEX}", width=3)
    draw.text((120, 225), "Alveolus\nPalv = +30", fill=f"#{COLOR_NAVY_HEX}")
    draw.text((280, 225), "EPP (Pairway = Ppl = +20)", fill=f"#{COLOR_TEAL_HEX}")
    draw.text((450, 225), "Mouth\nP = 0", fill=f"#{COLOR_CHARCOAL_HEX}")
    
    # Pressure indicators
    draw.text((100, 300), "• Driving Pressure = Palv = Pelastic (+10) + Ppleural (+20) = +30 cmH2O", fill=f"#{COLOR_CHARCOAL_HEX}")
    draw.text((100, 330), "• EPP occurs upstream in CARTILAGINOUS, large conducting airways.", fill=f"#{COLOR_CHARCOAL_HEX}")
    draw.text((100, 360), "• Airway cartilaginous rings prevent collapse downstream of EPP.", fill=f"#{COLOR_CHARCOAL_HEX}")
    draw.text((100, 390), "• Normal parenchymal tethering keeps non-cartilaginous airways open.", fill=f"#{COLOR_CHARCOAL_HEX}")
    draw.text((100, 430), "Result: Airflow proceeds efficiently with minimal dynamic collapse.", fill=f"#{COLOR_TEAL_HEX}")

    # COPD / Emphysema Panel
    draw.rounded_rectangle([610, 85, 1170, 520], radius=10, outline=f"#{COLOR_BURGUNDY_HEX}", width=2, fill="#FEF2F2")
    draw.text((630, 105), "B. COPD / EMPHYSEMATOUS FORCED EXPIRATION", fill=f"#{COLOR_BURGUNDY_HEX}")
    draw.text((630, 130), "Loss of elastic recoil + airway wall inflammation.", fill=f"#{COLOR_MUTED_HEX}")
    
    # Collapsed airway schematic
    draw.rectangle([680, 200, 860, 270], fill="#E2E8F0", outline=f"#{COLOR_NAVY_HEX}", width=3)
    draw.polygon([(860, 200), (940, 230), (940, 240), (860, 270)], fill="#FECACA", outline=f"#{COLOR_BURGUNDY_HEX}")
    draw.rectangle([940, 230, 1100, 240], fill="#FECACA", outline=f"#{COLOR_BURGUNDY_HEX}", width=3) # Collapsed segment
    
    draw.text((690, 225), "Alveolus\nPalv = +23", fill=f"#{COLOR_BURGUNDY_HEX}")
    draw.text((860, 175), "EPP Shifted Upstream!", fill=f"#{COLOR_BURGUNDY_HEX}")
    
    draw.text((640, 300), "• Reduced Elastic Recoil: Pelastic drops from +10 to +3 cmH2O.", fill=f"#{COLOR_CHARCOAL_HEX}")
    draw.text((640, 330), "• Driving Palv is reduced (+23 instead of +30 cmH2O).", fill=f"#{COLOR_CHARCOAL_HEX}")
    draw.text((640, 360), "• EPP shifts upstream into MEMBRANOUS, NON-CARTILAGINOUS bronchioles.", fill=f"#{COLOR_BURGUNDY_HEX}")
    draw.text((640, 390), "• Surrounding high Ppleural (+20) crushes the non-cartilaginous airway.", fill=f"#{COLOR_BURGUNDY_HEX}")
    draw.text((640, 420), "• Loss of alveolar radial tethering exacerbates collapse (Air-Trapping).", fill=f"#{COLOR_CHARCOAL_HEX}")
    draw.text((640, 460), "Result: Expiratory flow limitation, dynamic hyperinflation, and PEEPi.", fill=f"#{COLOR_BURGUNDY_HEX}")

    img.save(output_path, "PNG", dpi=(300, 300))
    print(f"  [+] Generated: {output_path.name}")

def create_gold_abe_diagram(output_path):
    """Generates the GOLD 2024/2025 ABE Assessment and Treatment Matrix."""
    img = Image.new("RGB", (1200, 650), "#FFFFFF")
    draw = ImageDraw.Draw(img)
    
    draw.rectangle([0, 0, 1200, 70], fill=f"#{COLOR_NAVY_HEX}")
    draw.text((30, 22), "GOLD 2024 / 2025: THE ABE ASSESSMENT SCHEME & INITIAL PHARMACOTHERAPY", fill="#FFFFFF")
    
    # Background Grid for ABE
    draw.rectangle([60, 100, 1140, 580], fill="#F8FAFC", outline="#CBD5E1", width=2)
    
    # Top Box: Group E (Exacerbators)
    draw.rounded_rectangle([100, 130, 1100, 310], radius=10, fill="#FEF3C7", outline=f"#{COLOR_BURGUNDY_HEX}", width=3)
    draw.text((120, 145), "GROUP E (FREQUENT OR SEVERE EXACERBATIONS)", fill=f"#{COLOR_BURGUNDY_HEX}")
    draw.text((120, 175), "CRITERIA: ≥ 2 Moderate Exacerbations per year OR ≥ 1 Exacerbation leading to Hospitalization (regardless of mMRC/CAT score)", fill=f"#{COLOR_CHARCOAL_HEX}")
    
    draw.rectangle([130, 210, 1070, 290], fill="#FFFFFF", outline="#D97706", width=2)
    draw.text((150, 225), "INITIAL TREATMENT: LABA + LAMA Combination", fill=f"#{COLOR_NAVY_HEX}")
    draw.text((150, 255), "UPFRONT TRIPLE THERAPY (LABA + LAMA + ICS): Indicated if Blood Eosinophils ≥ 300 cells/µL (IMPACT/ETHOS data)", fill=f"#{COLOR_BURGUNDY_HEX}")

    # Bottom Left: Group A
    draw.rounded_rectangle([100, 340, 570, 540], radius=10, fill="#F0FDF4", outline="#16A34A", width=3)
    draw.text((120, 355), "GROUP A (LOW SYMPTOMS, LOW RISK)", fill="#16A34A")
    draw.text((120, 385), "• 0 or 1 Moderate Exacerbation (no hospital)\n• mMRC 0 - 1  |  CAT < 10", fill=f"#{COLOR_CHARCOAL_HEX}")
    draw.rectangle([120, 440, 550, 510], fill="#FFFFFF", outline="#16A34A", width=2)
    draw.text((140, 460), "INITIAL TREATMENT: A Bronchodilator\n(Short or Long-Acting: SABA, LABA, or LAMA)", fill=f"#{COLOR_CHARCOAL_HEX}")

    # Bottom Right: Group B
    draw.rounded_rectangle([630, 340, 1100, 540], radius=10, fill="#EFF6FF", outline=f"#{COLOR_TEAL_HEX}", width=3)
    draw.text((650, 355), "GROUP B (HIGH SYMPTOMS, LOW RISK)", fill=f"#{COLOR_TEAL_HEX}")
    draw.text((650, 385), "• 0 or 1 Moderate Exacerbation (no hospital)\n• mMRC ≥ 2  |  CAT ≥ 10", fill=f"#{COLOR_CHARCOAL_HEX}")
    draw.rectangle([650, 440, 1080, 510], fill="#FFFFFF", outline=f"#{COLOR_TEAL_HEX}", width=2)
    draw.text((670, 460), "INITIAL TREATMENT: Dual Bronchodilator (LABA + LAMA)\n(Superior to monotherapy in symptom relief & quality of life)", fill=f"#{COLOR_NAVY_HEX}")

    # Footer note
    draw.text((100, 555), "*Note: Groups C & D from older guidelines are permanently merged into Group E to highlight exacerbation prevention.", fill=f"#{COLOR_MUTED_HEX}")

    img.save(output_path, "PNG", dpi=(300, 300))
    print(f"  [+] Generated: {output_path.name}")

def create_inhaler_selection_diagram(output_path):
    """Generates an evidence-based inhaler device selection algorithm."""
    img = Image.new("RGB", (1200, 550), "#FFFFFF")
    draw = ImageDraw.Draw(img)
    
    draw.rectangle([0, 0, 1200, 65], fill=f"#{COLOR_NAVY_HEX}")
    draw.text((30, 20), "INHALATION DELIVERY SYSTEMS: PATIENT ERGONOMICS & SELECTION ALGORITHM", fill="#FFFFFF")
    
    # Decision Step 1
    draw.rounded_rectangle([400, 90, 800, 150], radius=8, fill=f"#{COLOR_TEAL_HEX}")
    draw.text((430, 105), "ASSESS PEAK INSPIRATORY FLOW (PIF)\n& INSPIRATORY EFFORT (In-Check DIAL)", fill="#FFFFFF")
    
    # Left Branch: Low PIF (< 30 L/min) or Frail / Elderly / Exacerbation
    draw.line([(450, 150), (250, 210)], fill=f"#{COLOR_BURGUNDY_HEX}", width=4)
    draw.rounded_rectangle([50, 210, 450, 300], radius=8, fill="#FEE2E2", outline=f"#{COLOR_BURGUNDY_HEX}", width=2)
    draw.text((70, 225), "SUBOPTIMAL PIF (< 30 L/min) OR FRAIL", fill=f"#{COLOR_BURGUNDY_HEX}")
    draw.text((70, 255), "Inability to generate high turbulent flow.\nCommon in acute exacerbations, elderly, severe hyperinflation.", fill=f"#{COLOR_CHARCOAL_HEX}")
    
    # Recommendations for Left
    draw.rounded_rectangle([50, 330, 450, 510], radius=8, fill="#FFFFFF", outline=f"#{COLOR_BURGUNDY_HEX}", width=3)
    draw.text((70, 345), "RECOMMENDED DEVICES:", fill=f"#{COLOR_BURGUNDY_HEX}")
    draw.text((70, 375), "1. pMDI + Valved Holding Chamber (Spacer)\n   (Eliminates breath-actuation coordination)\n2. Soft Mist Inhaler (SMI - Respimat)\n   (Low velocity mist, flow-independent)\n3. Jet or Vibrating Mesh Nebulizer\n   (Passive tidal breathing delivery)", fill=f"#{COLOR_CHARCOAL_HEX}")

    # Right Branch: Adequate PIF (>= 30 - 60 L/min)
    draw.line([(750, 150), (950, 210)], fill="#16A34A", width=4)
    draw.rounded_rectangle([750, 210, 1150, 300], radius=8, fill="#DCFCE7", outline="#16A34A", width=2)
    draw.text((770, 225), "OPTIMAL PIF (≥ 30 - 60 L/min)", fill="#16A34A")
    draw.text((770, 255), "Able to generate forceful, deep inspiratory acceleration\nto de-aggregate drug powder into respirable particles.", fill=f"#{COLOR_CHARCOAL_HEX}")
    
    # Recommendations for Right
    draw.rounded_rectangle([750, 330, 1150, 510], radius=8, fill="#FFFFFF", outline="#16A34A", width=3)
    draw.text((770, 345), "RECOMMENDED DEVICES:", fill="#16A34A")
    draw.text((770, 375), "1. Dry Powder Inhalers (DPI - Single Dose)\n   (Rotahaler, Revolizer - highly cost-effective in India)\n2. Multi-dose DPI Reservoirs\n   (Diskus, Turbuhaler - convenient dose counters)\n3. Breath-Actuated pMDI / SMI Respimat", fill=f"#{COLOR_CHARCOAL_HEX}")

    img.save(output_path, "PNG", dpi=(300, 300))
    print(f"  [+] Generated: {output_path.name}")

def create_niv_algorithm_diagram(output_path):
    """Generates an acute exacerbation of COPD (AECOPD) NIV titration flowchart."""
    img = Image.new("RGB", (1200, 600), "#FFFFFF")
    draw = ImageDraw.Draw(img)
    
    draw.rectangle([0, 0, 1200, 65], fill=f"#{COLOR_NAVY_HEX}")
    draw.text((30, 20), "ACUTE RESPIRATORY FAILURE IN COPD: NON-INVASIVE VENTILATION (NIV) PROTOCOL", fill="#FFFFFF")
    
    # Triage step
    draw.rounded_rectangle([250, 85, 950, 160], radius=10, fill="#FEF3C7", outline="#D97706", width=2)
    draw.text((270, 100), "STEP 1: TRIAGE & BLOOD GAS ASSESSMENT (ON CONTROLLED OXYGEN 24-28%)", fill="#D97706")
    draw.text((270, 125), "Criteria for Acute Hypercapnic Respiratory Failure: pH < 7.35 AND PaCO2 > 45 mmHg with tachypnea (RR > 25)", fill=f"#{COLOR_CHARCOAL_HEX}")
    
    # Indications vs Contraindications
    draw.rounded_rectangle([50, 180, 570, 330], radius=8, fill="#EFF6FF", outline=f"#{COLOR_TEAL_HEX}", width=2)
    draw.text((70, 195), "STEP 2A: INITIATE BI-LEVEL POSITIVE AIRWAY PRESSURE", fill=f"#{COLOR_TEAL_HEX}")
    draw.text((70, 225), "• Mode: Spontaneous / Timed (S/T) with full face mask\n• Initial IPAP: 10 - 12 cmH2O (titrate up by 2 to target Vt 6-8 mL/kg)\n• Initial EPAP: 4 - 5 cmH2O (titrate for SpO2 88-92%)\n• Back-up Rate: 12 - 14 breaths/min\n• Rise Time: 0.1 - 0.2 s  |  Cycle Sensitivity: 30-40%", fill=f"#{COLOR_CHARCOAL_HEX}")

    draw.rounded_rectangle([630, 180, 1150, 330], radius=8, fill="#FEF2F2", outline=f"#{COLOR_BURGUNDY_HEX}", width=2)
    draw.text((650, 195), "STEP 2B: EXCLUDE ABSOLUTE CONTRAINDICATIONS", fill=f"#{COLOR_BURGUNDY_HEX}")
    draw.text((650, 225), "• Respiratory or cardiac arrest / gasping respirations\n• Severe encephalopathy (GCS < 8, unless direct CO2 narcosis)\n• Inability to clear copious secretions / high aspiration risk\n• Facial trauma, burns, or recent upper gastrointestinal surgery\n• Hemodynamic instability requiring escalating vasopressors", fill=f"#{COLOR_CHARCOAL_HEX}")

    # Re-evaluation step
    draw.rounded_rectangle([200, 360, 1000, 450], radius=10, fill=f"#{COLOR_NAVY_HEX}")
    draw.text((230, 380), "STEP 3: MANDATORY CLINICAL & REPEAT ABG REVIEW AT 1 TO 2 HOURS", fill="#FFFFFF")
    draw.text((230, 410), "Is pH improving (> 7.30), PaCO2 decreasing, and respiratory rate normalizing?", fill="#E2E8F0")

    # Outcomes
    draw.line([(350, 450), (250, 490)], fill="#16A34A", width=4)
    draw.rounded_rectangle([100, 490, 480, 570], radius=8, fill="#DCFCE7", outline="#16A34A", width=2)
    draw.text((120, 505), "IMPROVING (SUCCESS):", fill="#16A34A")
    draw.text((120, 530), "Continue NIV sessions for 16-24h/day;\ngradually wean as acidosis resolves.", fill=f"#{COLOR_CHARCOAL_HEX}")

    draw.line([(850, 450), (950, 490)], fill=f"#{COLOR_BURGUNDY_HEX}", width=4)
    draw.rounded_rectangle([720, 490, 1100, 570], radius=8, fill="#FEE2E2", outline=f"#{COLOR_BURGUNDY_HEX}", width=2)
    draw.text((740, 505), "FAILING (FAILURE):", fill=f"#{COLOR_BURGUNDY_HEX}")
    draw.text((740, 530), "No delay! Immediate Endotracheal Intubation\n& Invasive Lung-Protective Mechanical Ventilation.", fill=f"#{COLOR_CHARCOAL_HEX}")

    img.save(output_path, "PNG", dpi=(300, 300))
    print(f"  [+] Generated: {output_path.name}")

# Generate all diagram assets
diag_path1 = ASSETS_DIR / "copd_pathology_diagram.png"
diag_path2 = ASSETS_DIR / "equal_pressure_point_diagram.png"
diag_path3 = ASSETS_DIR / "gold_abe_algorithm_diagram.png"
diag_path4 = ASSETS_DIR / "inhaler_device_selection_diagram.png"
diag_path5 = ASSETS_DIR / "niv_management_algorithm_diagram.png"

create_pathology_diagram(diag_path1)
create_epp_diagram(diag_path2)
create_gold_abe_diagram(diag_path3)
create_inhaler_selection_diagram(diag_path4)
create_niv_algorithm_diagram(diag_path5)

# ---------------------------------------------------------------------------
# PART 2: WORD DOCUMENT (.DOCX) COMPOSITION & ADVANCED FORMATTING
# ---------------------------------------------------------------------------

doc = Document()

# Set standard 1-inch (2.54 cm) margins
for sec in doc.sections:
    sec.top_margin = Inches(1.0)
    sec.bottom_margin = Inches(1.0)
    sec.left_margin = Inches(1.0)
    sec.right_margin = Inches(1.0)
    sec.header.is_linked_to_previous = False
    sec.footer.is_linked_to_previous = False

# Helper XML styling functions
def set_cell_background(cell, hex_color):
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    cell._tc.get_or_add_tcPr().append(shd)

def set_cell_padding(cell, top=140, bottom=140, left=180, right=180):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def add_callout(document, title, text, alert_type="NOTE"):
    tbl = document.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    cell = tbl.cell(0, 0)
    cell.width = Inches(6.5)
    
    border_color = COLOR_TEAL_HEX if alert_type == "NOTE" else (COLOR_BURGUNDY_HEX if alert_type == "ALERT" else "16A34A")
    bg_color = "F0FDF4" if alert_type == "EVIDENCE" else ("FEF2F2" if alert_type == "ALERT" else "F4F6F8")
    
    set_cell_background(cell, bg_color)
    set_cell_padding(cell, top=160, bottom=160, left=220, right=200)
    
    # Left thick border
    tcPr = cell._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        f'<w:top w:val="none"/>'
        f'<w:left w:val="single" w:sz="36" w:space="0" w:color="{border_color}"/>'
        f'<w:bottom w:val="none"/>'
        f'<w:right w:val="none"/>'
        f'</w:tcBorders>'
    )
    tcPr.append(borders)
    
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(4)
    run_t = p.add_run(f"[{alert_type}] {title}\n")
    run_t.font.name = "Calibri"
    run_t.font.size = Pt(11)
    run_t.font.bold = True
    run_t.font.color.rgb = RGBColor.from_string(border_color)
    
    run_body = p.add_run(text)
    run_body.font.name = "Calibri"
    run_body.font.size = Pt(10)
    run_body.font.italic = False
    run_body.font.color.rgb = RGBColor.from_string(COLOR_CHARCOAL_HEX)
    
    p_after = document.add_paragraph()
    p_after.paragraph_format.space_before = Pt(2)
    p_after.paragraph_format.space_after = Pt(4)

def format_table(tbl, col_widths, headers, data):
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    # Format Header Row
    hdr_cells = tbl.rows[0].cells
    for i, header_text in enumerate(headers):
        hdr_cells[i].text = header_text
        hdr_cells[i].width = col_widths[i]
        set_cell_background(hdr_cells[i], COLOR_NAVY_HEX)
        set_cell_padding(hdr_cells[i], top=120, bottom=120, left=140, right=140)
        p = hdr_cells[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        for run in p.runs:
            run.font.name = "Calibri"
            run.font.size = Pt(9.5)
            run.font.bold = True
            run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            
    # Format Data Rows
    for r_idx, row_data in enumerate(data):
        row_cells = tbl.rows[r_idx + 1].cells
        bg_col = "FFFFFF" if r_idx % 2 == 0 else "F8FAFC"
        for c_idx, cell_value in enumerate(row_data):
            row_cells[c_idx].text = str(cell_value)
            row_cells[c_idx].width = col_widths[c_idx]
            set_cell_background(row_cells[c_idx], bg_col)
            set_cell_padding(row_cells[c_idx], top=100, bottom=100, left=140, right=140)
            p = row_cells[c_idx].paragraphs[0]
            for run in p.runs:
                run.font.name = "Calibri"
                run.font.size = Pt(9)
                run.font.color.rgb = RGBColor.from_string(COLOR_CHARCOAL_HEX)

print("Writing Textbook Chapters...")

# ---------------------------------------------------------------------------
# COVER PAGE & FRONT MATTER
# ---------------------------------------------------------------------------
p_cov = doc.add_paragraph()
p_cov.paragraph_format.space_before = Pt(40)
p_cov.paragraph_format.space_after = Pt(10)
p_cov.alignment = WD_ALIGN_PARAGRAPH.CENTER
r_pre = p_cov.add_run("POSTGRADUATE MONOGRAPH SERIES IN RESPIRATORY MEDICINE\n")
r_pre.font.name = "Calibri"
r_pre.font.size = Pt(13)
r_pre.font.bold = True
r_pre.font.color.rgb = RGBColor.from_string(COLOR_TEAL_HEX)

r_main = p_cov.add_run("CHRONIC OBSTRUCTIVE PULMONARY DISEASE\n(COPD)\n")
r_main.font.name = "Calibri"
r_main.font.size = Pt(30)
r_main.font.bold = True
r_main.font.color.rgb = RGBColor.from_string(COLOR_NAVY_HEX)

r_sub = p_cov.add_run("A Comprehensive Academic & Clinical Reference Manual for MD & DNB Candidates\nIncorporating GOLD 2024/2025, GINA, ERS/ATS, and Indian Chest Society Guidelines\n")
r_sub.font.name = "Calibri"
r_sub.font.size = Pt(14)
r_sub.font.color.rgb = RGBColor.from_string(COLOR_MUTED_HEX)

doc.add_paragraph()
doc.add_paragraph()

# Author & Institutional Box
p_auth = doc.add_paragraph()
p_auth.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_auth.paragraph_format.space_before = Pt(80)
r_ed = p_auth.add_run("DEPARTMENT OF RESPIRATORY MEDICINE & CRITICAL CARE\nFIRST EDITION | COMPREHENSIVE TEXTBOOK VOLUME\n")
r_ed.font.name = "Calibri"
r_ed.font.size = Pt(11)
r_ed.font.bold = True
r_ed.font.color.rgb = RGBColor.from_string(COLOR_BURGUNDY_HEX)

doc.add_page_break()

# ---------------------------------------------------------------------------
# PREFACE & CURRICULUM SYLLABUS
# ---------------------------------------------------------------------------
h_pref = doc.add_heading("Preface & Educational Blueprint", level=1)
h_pref.paragraph_format.space_before = Pt(18)
h_pref.paragraph_format.space_after = Pt(8)
for r in h_pref.runs:
    r.font.name = "Calibri"
    r.font.size = Pt(18)
    r.font.bold = True
    r.font.color.rgb = RGBColor.from_string(COLOR_NAVY_HEX)

doc.add_paragraph(
    "Chronic Obstructive Pulmonary Disease (COPD) remains one of the greatest global and Indian public health challenges, "
    "ranking consistently as the third leading cause of mortality worldwide and the second leading contributor to disability-adjusted "
    "life years (DALYs) in the Indian subcontinent. For postgraduate residents undertaking the Doctor of Medicine (MD) and Diplomate of "
    "National Board (DNB) in Respiratory Medicine, mastering COPD requires far more than rote memorization of spirometric cutoffs. It demands "
    "an integrated grasp of basic cellular biology, aerosol physics, respiratory mechanics, critical care physiology, and real-world clinical decision-making."
)
doc.add_paragraph(
    "This textbook volume is structured to bridge foundational knowledge with bedside postgraduate practice. Every chapter synthesizes the definitive "
    "authorities—Murray & Nadel's Textbook of Respiratory Medicine, Crofton & Douglas's Respiratory Diseases, and Harrison's Principles of Internal Medicine—"
    "with the latest clinical guidelines from the Global Initiative for Chronic Obstructive Lung Disease (GOLD 2024/2025 Updates), the Indian Chest Society (ICS), "
    "and the National College of Chest Physicians (NCCP)."
)

add_callout(
    doc,
    "CORE EVIDENCE FOUNDATION & CLINICAL SAFETY MANDATE",
    "All clinical recommendations and pharmacological dosages presented herein are derived from prospective multicenter clinical trials "
    "(TORCH, UPLIFT, FLAME, IMPACT, ETHOS, REDUCE, NETT, BOREAS) and authorized consensus guidelines. Pharmacotherapy in clinical practice must always "
    "be verified against individual patient renal, hepatic, and cardiac comorbidities. This manual serves as an advanced educational and clinical decision-support resource.",
    "ALERT"
)

# Table of Contents Summary
h_toc = doc.add_heading("Table of Contents & Chapter Organization", level=2)
for r in h_toc.runs:
    r.font.name = "Calibri"; r.font.size = Pt(14); r.font.bold = True; r.font.color.rgb = RGBColor.from_string(COLOR_TEAL_HEX)

toc_data = [
    ["Chapter 1", "Definition, Evolution, and Global & Indian Epidemiology", "Page 4"],
    ["Chapter 2", "Etiopathogenesis, Molecular Biology, and Cellular Pathology", "Page 11"],
    ["Chapter 3", "Respiratory Mechanics & Pathophysiology: Flow Limitation & EPP", "Page 19"],
    ["Chapter 4", "Clinical Phenotypes & Bedside Semiology (Pink Puffer vs. Blue Bloater)", "Page 27"],
    ["Chapter 5", "Comprehensive Pulmonary Function Testing: Spirometry, Plethysmography, DLCO", "Page 35"],
    ["Chapter 6", "Thoracic Radiology & Biomarkers: CXR Signs, HRCT Phenotypes, CTPA", "Page 44"],
    ["Chapter 7", "Stable COPD Assessment & Staging: The GOLD 2024/2025 ABE Paradigm & BODE Index", "Page 52"],
    ["Chapter 8", "Comprehensive Pharmacotherapy: Inhalation Devices, Dual Bronchodilation, SITT, Biologics", "Page 60"],
    ["Chapter 9", "Non-Pharmacological Management: Pulmonary Rehabilitation, LTOT, Nocturnal NIV, Vaccines", "Page 70"],
    ["Chapter 10", "Interventional Pulmonology & Surgery: Endobronchial Valves (BLVR), LVRS, Bullectomy", "Page 78"],
    ["Chapter 11", "Acute Exacerbations of COPD (AECOPD): Triage, Inpatient Care, NIV Protocols, Invasive Ventilation", "Page 85"],
    ["Chapter 12", "Systemic Manifestations, Comorbidities, Overlap Syndromes (ACO, Overlap with OSA), Palliative Care", "Page 93"],
    ["Appendix", "Postgraduate MD/DNB Table Viva Spotters, Landmark Trial Summaries & 10-Marker Theory Models", "Page 101"]
]
t_toc = doc.add_table(rows=len(toc_data) + 1, cols=3)
format_table(t_toc, [Inches(1.5), Inches(4.0), Inches(1.0)], ["Chapter #", "Syllabus Module & Chapter Title", "Target"], toc_data)

doc.add_page_break()

# ---------------------------------------------------------------------------
# CHAPTER 1: DEFINITION, EVOLUTION & EPIDEMIOLOGY
# ---------------------------------------------------------------------------
h_c1 = doc.add_heading("Chapter 1: Definition, Evolution, and Epidemiology of COPD", level=1)
for r in h_c1.runs:
    r.font.name = "Calibri"; r.font.size = Pt(16); r.font.bold = True; r.font.color.rgb = RGBColor.from_string(COLOR_NAVY_HEX)

doc.add_paragraph(
    "1.1 Historical Evolution of the Concept\n"
    "The clinical and anatomical recognition of chronic obstructive pulmonary disease evolved over three centuries. "
    "In 1679, Théophile Bonet described 'voluminous lungs' in post-mortem specimens. René Laennec in 1819, utilizing his invention of the stethoscope, "
    "distinguished emphysema ('an enlargement of air vesicles with loss of elasticity') from chronic bronchitis. In 1959, the Ciba Guest Symposium "
    "established the first formal consensus definitions separating emphysema (an anatomical, destructive process) from chronic bronchitis "
    "(a clinical definition based on chronic cough and sputum expectoration for at least 3 months in 2 consecutive years)."
)
doc.add_paragraph(
    "1.2 Modern GOLD 2024/2025 Definition\n"
    "The Global Initiative for Chronic Obstructive Lung Disease (GOLD) defines COPD as: 'A heterogeneous lung condition characterized by chronic "
    "respiratory symptoms (dyspnea, cough, sputum production and/or exacerbations) due to abnormalities of the airways (bronchitis, bronchiolitis) "
    "and/or alveoli (emphysema) that cause persistent, often progressive, airflow limitation.' This updated definition intentionally deemphasizes "
    "the historic dichotomization of chronic bronchitis versus emphysema, recognizing that the overwhelming majority of patients harbor a combined, "
    "variable spectrum of both small airway disease and parenchymal destruction."
)
doc.add_paragraph(
    "1.3 Global Epidemiology & Disease Burden\n"
    "According to the Global Burden of Disease (GBD) study, approximately 392 million individuals worldwide live with COPD, with an overall global "
    "prevalence of 10.3% in adults aged 30 to 79 years. Over 3.23 million deaths occur annually from COPD, representing 81.7% of all chronic respiratory disease "
    "deaths. Over 80% of these deaths occur in low- and middle-income countries (LMICs), reflecting profound inequalities in exposure to environmental "
    "pollutants, diagnostic spirometry access, and availability of essential inhaled maintenance therapies."
)
doc.add_paragraph(
    "1.4 The Indian Epidemiological Landscape: The Biomass Smoke & Air Pollution Reality\n"
    "In India, COPD presents distinct epidemiological characteristics compared to Western cohorts. The ICMR-PHFI State-Level Disease Burden study "
    "demonstrated that COPD is the second leading cause of disease burden and mortality in India, responsible for approximately 75.6% of all chronic "
    "respiratory disability. Critically, approximately 50% of COPD cases in India occur in NON-SMOKERS. The key driving risk factors include:\n"
    "1. Biomass Fuel Exposure (The Chulha Phenomenon): Over 500 million rural individuals in India remain exposed to domestic combustion of unprocessed solid biomass fuels "
    "(wood, cow dung cakes, agricultural residues) in poorly ventilated indoor kitchens. Women cooking over traditional chulhas endure extreme particulate concentrations "
    "(PM2.5 often exceeding 1000 µg/m3), developing a severe phenotype characterized by predominant small airway disease, anthracofibrosis, and severe chronic bronchitis.\n"
    "2. Ambient Particulate Matter Air Pollution: Year-round exposure to hazardous urban ambient air pollution (vehicular emissions, crop residue burning, industrial dust) "
    "stunts maximal childhood lung growth and accelerates adulthood lung function decline.\n"
    "3. Post-Tuberculosis Structural Airway Disease: Tuberculosis leaves behind residual cicatricial scarring, distorted conducting airways, and traction bronchiectasis, "
    "which chronically limits airflow and frequently mimics or co-exists with smoker's COPD."
)

epi_table_data = [
    ["Parameter", "Western Smoker COPD Cohorts", "Indian Biomass / Non-Smoker COPD Cohorts"],
    ["Primary Etiology", "Cigarette smoking (> 20 - 30 pack-years)", "Biomass fuel smoke (chulha > 200 hour-years), ambient PM2.5, bidi smoking"],
    ["Gender Ratio", "Historical male predominance; now near equal (1:1)", "High proportion of rural females (up to 40 - 50% non-smoking females)"],
    ["Pathological Spectrum", "Extensive centrilobular emphysema with bullae", "Prominent small airway bronchiolitis, peribronchiolar anthracosis, less emphysema"],
    ["Airflow Limitation", "Severe post-BD FEV1 reduction; rapid FEV1 decline", "Moderate-to-severe obstruction with marked air-trapping; preserved DLCO"],
    ["Comorbidities", "Severe coronary artery disease, peripheral vascular disease", "Cor pulmonale, pulmonary hypertension, recurrent biomass-related infections"]
]
t_epi = doc.add_table(rows=len(epi_table_data), cols=3)
format_table(t_epi, [Inches(1.8), Inches(2.4), Inches(2.3)], epi_table_data[0], epi_table_data[1:])

doc.add_paragraph()
add_callout(
    doc,
    "GENETIC SUSCEPTIBILITY: α1-ANTITRYPSIN DEFICIENCY (AATD)",
    "Alpha-1 Antitrypsin Deficiency is the best-characterized genetic cause of COPD, caused by mutations in the SERPINA1 gene on chromosome 14q32. "
    "Normal allele is Pi*M. The Pi*Z mutation (Glu342Lys) leads to protein misfolding in hepatocytes (causing hepatic cirrhosis) and severe serum AAT deficiency "
    "(< 11 µmol/L or < 50 mg/dL). Patients develop early-onset (ages 30-45) panacinar emphysema with lower-lobe predominance, accelerated tenfold by smoking. "
    "WHO and GOLD mandate screening all newly diagnosed COPD patients once in their lifetime with serum AAT concentration. Confirmatory testing involves "
    "isoelectric focusing phenotyping and SERPINA1 targeted genotyping.",
    "EVIDENCE"
)

doc.add_page_break()

# ---------------------------------------------------------------------------
# CHAPTER 2: ETIOPATHOGENESIS, PATHOLOGY & MOLECULAR BIOLOGY
# ---------------------------------------------------------------------------
h_c2 = doc.add_heading("Chapter 2: Etiopathogenesis, Cellular Mechanisms, and Pathology", level=1)
for r in h_c2.runs:
    r.font.name = "Calibri"; r.font.size = Pt(16); r.font.bold = True; r.font.color.rgb = RGBColor.from_string(COLOR_NAVY_HEX)

doc.add_paragraph(
    "2.1 The Multi-Hit Pathogenetic Framework\n"
    "The pathogenesis of COPD represents an amplified, chronic inflammatory response coupled with structural tissue remodeling, cellular senescence, "
    "and defective repair mechanisms triggered by inhaled noxious particulates. Three classic, interlocking hypotheses form the core pathophysiological triad:"
)
doc.add_paragraph(
    "1. The Protease-Antiprotease Imbalance Hypothesis:\n"
    "Endogenous proteolytic enzymes (neutrophil elastase, proteinase 3, cathepsin G, and Matrix Metalloproteinases MMP-8, MMP-9, MMP-12) digest elastin, collagen, "
    "and extracellular matrix components forming the alveolar framework. Under physiological conditions, these enzymes are tightly regulated by antiproteases "
    "(alpha-1 antitrypsin, secretory leukoprotease inhibitor [SLPI], and tissue inhibitors of metalloproteinases [TIMPs]). In COPD, intense recruitment of activated "
    "neutrophils and alveolar macrophages overwhelms antiprotease defenses, causing permanent alveolar wall destruction and airspace enlargement."
)
doc.add_paragraph(
    "2. The Oxidant-Antioxidant Imbalance & Cellular Senescence:\n"
    "Inhaled smoke contains over 10^15 free radicals per puff (superoxide anions, hydroxyl radicals, hydrogen peroxide, peroxynitrite). Endogenous inflammatory cells "
    "further generate reactive oxygen species (ROS). Oxidative stress inactivates antiproteases (oxidizing the active Met358 residue of AAT), upregulates pro-inflammatory "
    "transcription factor NF-κB, impairs nuclear factor erythroid 2-related factor 2 (Nrf2) antioxidant pathways, and induces double-strand DNA damage. This drives "
    "respiratory cells into the Senescence-Associated Secretory Phenotype (SASP), characterized by irreversible cell cycle arrest and continuous release of IL-1β, IL-6, "
    "and TNF-α."
)
doc.add_paragraph(
    "3. Chronic Inflammatory Cellular Orchestration & Corticosteroid Resistance:\n"
    "Alveolar macrophages act as the primary conductor of the immune response, releasing CXCL8 (IL-8), Leukotriene B4 (LTB4), and TNF-α, which chemoattract neutrophils. "
    "CD8+ cytotoxic Tc1 lymphocytes infiltrate the airway mucosa and alveolar walls, releasing granzyme B and perforin to induce epithelial apoptosis. "
    "A fundamental hallmark of COPD inflammation is Corticosteroid Resistance: Oxidative stress selectively inactivates Histone Deacetylase-2 (HDAC2), the critical enzyme "
    "recruited by activated glucocorticoid receptors to switch off pro-inflammatory genes. Without functional HDAC2, systemic and inhaled corticosteroids cannot switch off "
    "NF-κB-driven cytokine production in classic neutrophilic COPD."
)

# Insert Pathology Diagram
p_img1 = doc.add_paragraph()
p_img1.alignment = WD_ALIGN_PARAGRAPH.CENTER
doc.add_picture(str(diag_path1), width=Inches(6.2))
p_cap1 = doc.add_paragraph("Figure 2.1: Histopathological spectrum of COPD: Chronic Bronchitis vs. Centriacinar vs. Panacinar Emphysema.")
p_cap1.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_cap1.runs[0].font.size = Pt(9.5)
p_cap1.runs[0].font.italic = True
p_cap1.runs[0].font.color.rgb = RGBColor.from_string(COLOR_MUTED_HEX)

doc.add_paragraph(
    "2.2 Structural Pathology & Airway Remodeling\n"
    "1. Large Conducting Airways (> 2 mm internal diameter): Goblet cell hyperplasia, enlargement and hypertrophy of submucosal bronchial mucus glands "
    "(quantified histopathologically by the Reid Index: ratio of gland thickness to total bronchial wall thickness from epithelium to perichondrium; normal < 0.40, "
    "in chronic bronchitis > 0.50 - 0.60), loss of ciliated surface epithelium, and squamous metaplasia.\n"
    "2. Small Conducting Airways (< 2 mm internal diameter - The Main Site of Resistance): Termed 'silent zone' bronchiolitis. Characterized by luminal mucus plugging, "
    "inflammatory cellular infiltration, micro-ulceration, and progressive peribronchiolar adventitial fibrosis that permanently narrows the internal airway lumen.\n"
    "3. Lung Parenchyma (Emphysema): Defined strictly on anatomical criteria as abnormal, permanent enlargement of airspaces distal to the terminal bronchioles, "
    "accompanied by destruction of their walls without obvious fibrosis."
)

doc.add_page_break()

# ---------------------------------------------------------------------------
# CHAPTER 3: PULMONARY PATHOPHYSIOLOGY & RESPIRATORY MECHANICS
# ---------------------------------------------------------------------------
h_c3 = doc.add_heading("Chapter 3: Respiratory Mechanics & Pathophysiology of Flow Limitation", level=1)
for r in h_c3.runs:
    r.font.name = "Calibri"; r.font.size = Pt(16); r.font.bold = True; r.font.color.rgb = RGBColor.from_string(COLOR_NAVY_HEX)

doc.add_paragraph(
    "3.1 Mechanisms of Expiratory Flow Limitation\n"
    "Expiratory airflow is governed by the driving pressure divided by airway resistance: V' = (Palv - Pao) / Raw. "
    "In COPD, expiratory flow limitation arises from two simultaneous mechanical abnormalities:\n"
    "1. Increased Upstream Airway Resistance: Narrowing of small conducting bronchioles (< 2 mm) by mucosal thickening, mucus exudate, and fibrous remodeling.\n"
    "2. Loss of Alveolar Elastic Recoil Pressure: Destruction of elastic fibers reduces the static alveolar driving pressure (Pst,l) that drives air out during passive exhalation."
)
doc.add_paragraph(
    "3.2 The Equal Pressure Point (EPP) Theory & Dynamic Airway Compression\n"
    "During quiet expiration, intrapleural pressure (Ppl) remains negative (-5 cmH2O). However, during forced expiration, active contraction of expiratory abdominal "
    "and internal intercostal muscles raises Ppl to highly positive levels (e.g., +20 cmH2O). The total pressure inside the alveolus (Palv) equals the sum of intrapleural "
    "pressure plus the static lung elastic recoil pressure: Palv = Ppl + Pel.\n"
    "As air flows from alveoli toward the mouth, pressure drops progressively due to frictional resistance. The anatomical site along the airway where intraluminal pressure "
    "drops to precisely equal surrounding intrapleural pressure (Pairway = Ppl) is the EQUAL PRESSURE POINT (EPP)."
)

# Insert EPP Diagram
p_img2 = doc.add_paragraph()
p_img2.alignment = WD_ALIGN_PARAGRAPH.CENTER
doc.add_picture(str(diag_path2), width=Inches(6.2))
p_cap2 = doc.add_paragraph("Figure 3.1: Mechanism of the Equal Pressure Point (EPP) and upstream dynamic airway collapse in healthy vs. emphysematous lungs.")
p_cap2.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_cap2.runs[0].font.size = Pt(9.5); p_cap2.runs[0].font.italic = True; p_cap2.runs[0].font.color.rgb = RGBColor.from_string(COLOR_MUTED_HEX)

doc.add_paragraph(
    "3.3 Dynamic Hyperinflation & Intrinsic PEEP (auto-PEEP)\n"
    "Because expiratory flow is severely limited, exhalation is markedly prolonged. When a COPD patient exercises or experiences tachypnea, the expiratory time (Te) "
    "becomes shorter than the time required for the respiratory system to empty down to its baseline resting functional residual capacity (FRC). Subsequent inspiration "
    "begins before complete exhalation has finished. This phenomenon is DYNAMIC HYPERINFLATION.\n"
    "Consequences of Dynamic Hyperinflation:\n"
    "1. Development of Intrinsic Positive End-Expiratory Pressure (auto-PEEP / PEEPi): A positive alveolar pressure remains at end-expiration (often 5 to 15 cmH2O). "
    "The inspiratory muscles must generate a negative intrapleural pressure exceeding this auto-PEEP threshold just to initiate inspiratory airflow (an enormous threshold load).\n"
    "2. Decreased Diaphragmatic Curvature: The diaphragm is pushed downward into a flattened, shortened configuration, shifting muscle fibers out of their optimal "
    "length-tension relationship. Flattened fibers pull the lower costal margin inward during inspiration (Hoover's sign) rather than lifting it upward.\n"
    "3. Neuromechanical Dissociation: The brain's central motor respiratory drive commands high ventilation, but the stiff, hyperinflated mechanical lung cannot respond "
    "with corresponding volume expansion. This mismatch is perceived clinically as the suffocating sensation of dyspnea."
)
doc.add_paragraph(
    "3.4 Ventilation-Perfusion (V/Q) Mismatch and Gas Exchange Abnormalities\n"
    "COPD produces wide heterogeneity of ventilation and perfusion throughout the pulmonary vascular bed:\n"
    "1. Low V/Q Regions: Poorly ventilated, mucus-obstructed airways with preserved pulmonary blood flow cause physiological venous admixture (shunt effect), "
    "manifesting as arterial hypoxemia (low PaO2).\n"
    "2. High V/Q Regions: Severely emphysematous bullae with destroyed capillary beds maintain ventilation without perfusion, increasing physiological dead space (Vd/Vt), "
    "impairing CO2 clearance, and resulting in chronic hypercapnia (elevated PaCO2)."
)

doc.add_page_break()

# ---------------------------------------------------------------------------
# CHAPTER 4: CLINICAL PHENOTYPES & BEDSIDE SEMIOLOGY
# ---------------------------------------------------------------------------
h_c4 = doc.add_heading("Chapter 4: Clinical Phenotypes and Bedside Physical Examination", level=1)
for r in h_c4.runs:
    r.font.name = "Calibri"; r.font.size = Pt(16); r.font.bold = True; r.font.color.rgb = RGBColor.from_string(COLOR_NAVY_HEX)

doc.add_paragraph(
    "4.1 Clinical Symptomatology\n"
    "The clinical presentation of COPD rests upon a triad of symptoms developing insidiously over years, typically becoming clinically manifest after age 40:\n"
    "1. Progressive Exertional Dyspnea: Initially noticed only during strenuous exertion (climbing hills/stairs); relentlessly progresses over years until dressing, "
    "shaving, or talking produces breathlessness. Quantified using the validated modified Medical Research Council (mMRC) dyspnea scale (Grade 0 to 4).\n"
    "2. Chronic Productive Cough: Often the earliest presenting symptom, traditionally dismissed by smokers as 'smoker's cough'. Characteristically worse upon awakening "
    "with expectoration of small quantities of tenacious mucoid sputum.\n"
    "3. Episodic Wheezing & Chest Tightness: Fluctuates with viral respiratory tract infections, cold weather exposure, or exposure to environmental irritants."
)

# mMRC Table
mmrc_data = [
    ["mMRC Grade", "Severity Description", "Clinical Disability Context"],
    ["Grade 0", "Dyspnea only with strenuous exercise", "Normal physical capacity for age"],
    ["Grade 1", "Hurrying on level ground or walking up a slight hill", "Mild limitation; avoids brisk walking"],
    ["Grade 2", "Walks slower than peers on level ground or stops for breath", "Cannot keep pace with contemporaries"],
    ["Grade 3", "Stops for breath after walking ~100 meters or after a few minutes", "Severe limitation; restricted to home vicinity"],
    ["Grade 4", "Too breathless to leave house or breathless dressing/undressing", "Extreme functional disability; housebound"]
]
t_mmrc = doc.add_table(rows=len(mmrc_data), cols=3)
format_table(t_mmrc, [Inches(1.2), Inches(3.2), Inches(2.1)], mmrc_data[0], mmrc_data[1:])

doc.add_paragraph()
doc.add_paragraph(
    "4.2 The Classical Clinical Archetypes: Pink Puffer vs. Blue Bloater\n"
    "In classical clinical semiology, extreme phenotypes of COPD were characterized as Type A ('Pink Puffer') and Type B ('Blue Bloater'). While most modern patients "
    "harbor features of both, recognizing these archetypes remains high-yield for clinical examinations:"
)

pp_bb_data = [
    ["Clinical Feature", "Type A: Pink Puffer (Predominant Emphysema)", "Type B: Blue Bloater (Predominant Chronic Bronchitis)"],
    ["Primary Pathology", "Parenchymal alveolar destruction & loss of recoil", "Chronic airway mucosal inflammation & mucus hypersecretion"],
    ["Body Habitus & Weight", "Thin, asthenic, marked cachexia and muscle wasting", "Stocky, plethoric, overweight or obese"],
    ["Respiratory Drive", "Intact, high respiratory drive; breathless but not cyanotic", "Blunted, depressed chemoreceptor respiratory drive"],
    ["Breath Sounds", "Diminished / quiet breath sounds throughout; prolonged expiration", "Coarse crackles, rhonchi, polyphonic wheezing throughout"],
    ["Arterial Blood Gases", "Mild hypoxemia (PaO2 65-75), normal PaCO2 until terminal", "Severe hypoxemia (PaO2 < 55), chronic hypercapnia (PaCO2 > 50)"],
    ["Hematocrit", "Normal (40 - 45%)", "Elevated (Secondary polycythemia, hematocrit > 55%)"],
    ["Cor Pulmonale", "Late terminal development", "Early, frequent cor pulmonale with marked peripheral edema"],
    ["Diffusing Capacity (DLCO)", "Markedly reduced (< 50% predicted)", "Normal to mildly reduced"]
]
t_ppbb = doc.add_table(rows=len(pp_bb_data), cols=3)
format_table(t_ppbb, [Inches(1.6), Inches(2.4), Inches(2.5)], pp_bb_data[0], pp_bb_data[1:])

doc.add_paragraph()
doc.add_paragraph(
    "4.3 Bedside Physical Signs: Technique & Pathophysiological Mechanism\n"
    "1. Hoover's Sign: Inward paradoxical movement of the lower lateral ribcage (costal margin) during inspiration. Mechanism: In severe hyperinflation, the flattened "
    "diaphragm has muscle fibers oriented transversely toward the central tendon. Contraction pulls the lower ribs inward toward the midline rather than lifting them outward.\n"
    "2. Campbell's Sign / Fowler's Sign: Downward tracheal movement during inspiration. Mechanism: Powerful diaphragmatic descent pulls downward on the intrathoracic mediastinum "
    "and tracheobronchial tree.\n"
    "3. Reduced Cricosternal Distance: Normal distance between the thyroid cartilage and suprasternal notch is 3 to 4 fingerbreadths. In hyperinflated COPD, it is reduced "
    "to < 2 fingerbreadths due to permanent elevation of the sternum.\n"
    "4. Dahl's Sign: Patches of hyperpigmentation and calluses over the anterior distal thighs, caused by constant resting of the elbows on the knees in the classic tripod position.\n"
    "5. Pursed-Lip Breathing: Exhalation through puckered lips. Mechanism: Imposes an artificial downstream resistance at the lips, raising intraluminal airway pressure "
    "throughout the bronchial tree and preventing dynamic collapse upstream of the equal pressure point."
)

doc.add_page_break()

# ---------------------------------------------------------------------------
# CHAPTER 5: COMPREHENSIVE PULMONARY FUNCTION TESTING
# ---------------------------------------------------------------------------
h_c5 = doc.add_heading("Chapter 5: Comprehensive Pulmonary Function Testing (PFT)", level=1)
for r in h_c5.runs:
    r.font.name = "Calibri"; r.font.size = Pt(16); r.font.bold = True; r.font.color.rgb = RGBColor.from_string(COLOR_NAVY_HEX)

doc.add_paragraph(
    "5.1 Spirometric Quality Control & Standardization (ATS/ERS 2019 Criteria)\n"
    "Post-bronchodilator spirometry is the mandatory gold standard for diagnosing COPD. Before clinical interpretation, three technical criteria must be strictly verified:\n"
    "1. Acceptability at Start: Rapid rise to Peak Expiratory Flow (PEF). Extrapolated volume (VBE) must be < 5% of FVC or < 0.100 L (100 mL), whichever is greater.\n"
    "2. Acceptability during Maneuver: Minimum Forced Expiratory Time (FET) >= 6 seconds in adults, free of coughing or early glottic closure in the first second.\n"
    "3. End-of-Test Criteria: An expiratory plateau (< 0.025 L volume change over the final 1.0 second of expiration).\n"
    "4. Repeatability: Patient performs a minimum of 3 acceptable maneuvers. The difference between the two largest FVC values and the two largest FEV1 values must be <= 0.150 L (150 mL)."
)
doc.add_paragraph(
    "5.2 The Diagnostic Cutoff: LLN vs. The Fixed 0.70 Ratio Debate\n"
    "GOLD mandates a post-bronchodilator FEV1/FVC < 0.70 to establish persistent airflow limitation. However, physiological societies (ATS/ERS) advocate using the "
    "Lower Limit of Normal (LLN - 5th percentile of a healthy age-matched reference population, z-score < -1.645):\n"
    "• In the Elderly (> 65 years): Normal aging causes loss of elastic recoil, lowering physiological FEV1/FVC. Using the fixed 0.70 ratio results in up to 20-30% false-positive "
    "overdiagnosis of COPD.\n"
    "• In Young Adults (< 45 years): Healthy normal FEV1/FVC is 0.80 to 0.85. A ratio of 0.73 represents marked airway obstruction in a 35-year-old, yet is misclassified "
    "as 'normal' under the fixed 0.70 criterion."
)
doc.add_paragraph(
    "5.3 GOLD Spirometric Staging of Airflow Limitation (In Patients with FEV1/FVC < 0.70):\n"
    "• GOLD 1 (Mild): Post-BD FEV1 >= 80% predicted.\n"
    "• GOLD 2 (Moderate): 50% <= Post-BD FEV1 < 80% predicted.\n"
    "• GOLD 3 (Severe): 30% <= Post-BD FEV1 < 50% predicted.\n"
    "• GOLD 4 (Very Severe): Post-BD FEV1 < 30% predicted."
)
doc.add_paragraph(
    "5.4 Bronchodilator Reversibility (BDR) Testing\n"
    "Protocol: Administer 400 µg of Salbutamol (or 160 µg Ipratropium) via a valved holding chamber; repeat spirometry after 15 to 20 minutes.\n"
    "• Modern ATS/ERS 2021 Criterion: An increase in FEV1 or FVC by > 10% OF THE PREDICTED VALUE.\n"
    "• Traditional Criterion: An increase in FEV1 or FVC by > 12% AND > 200 mL from pre-bronchodilator baseline.\n"
    "Clinical Pearl: High reversibility does NOT exclude COPD. If post-bronchodilator FEV1/FVC remains < 0.70, persistent obstruction is confirmed, and the diagnosis is COPD or Asthma-COPD Overlap (ACO)."
)
doc.add_paragraph(
    "5.5 Lung Volumes & Diffusing Capacity (DLCO)\n"
    "• Body Plethysmography: Demonstrates elevated Functional Residual Capacity (FRC), markedly elevated Residual Volume (RV), and elevated Total Lung Capacity (TLC). "
    "An RV/TLC ratio > 120-140% of predicted confirms severe air-trapping.\n"
    "• DLCO (Single-Breath Carbon Monoxide Uptake): Corrected for baseline hemoglobin. In emphysema, DLCO is significantly reduced (< 70% predicted) due to capillary bed destruction. "
    "In pure chronic bronchitis and uncomplicated bronchial asthma, DLCO remains normal or elevated."
)

pft_diff_data = [
    ["Parameter", "Normal", "COPD (Emphysema)", "COPD (Chronic Bronchitis)", "Bronchial Asthma"],
    ["FEV1 / FVC Ratio", "> 0.70 (> LLN)", "< 0.70 (< LLN)", "< 0.70 (< LLN)", "Normal or < 0.70 (Variable)"],
    ["BDR Response", "Negative", "Variable / Partial", "Variable / Partial", "Marked (Often normalizes)"],
    ["TLC (Plethysmography)", "80 - 120% pred", "Markedly Increased (> 120%)", "Normal to mildly increased", "Normal to mildly increased"],
    ["RV / TLC Ratio", "< 35%", "Markedly Elevated (> 45-55%)", "Elevated (35 - 45%)", "Transiently elevated in attacks"],
    ["DLCO (Hb-corrected)", "> 80% pred", "Markedly Reduced (< 60%)", "Normal (80 - 100%)", "Normal or Supranormal (> 100%)"]
]
t_pft = doc.add_table(rows=len(pft_diff_data), cols=5)
format_table(t_pft, [Inches(1.5), Inches(1.1), Inches(1.3), Inches(1.3), Inches(1.3)], pft_diff_data[0], pft_diff_data[1:])

doc.add_page_break()

# ---------------------------------------------------------------------------
# CHAPTER 6: THORACIC RADIOLOGY & BIOMARKERS
# ---------------------------------------------------------------------------
h_c6 = doc.add_heading("Chapter 6: Thoracic Imaging and Laboratory Biomarkers", level=1)
for r in h_c6.runs:
    r.font.name = "Calibri"; r.font.size = Pt(16); r.font.bold = True; r.font.color.rgb = RGBColor.from_string(COLOR_NAVY_HEX)

doc.add_paragraph(
    "6.1 Chest Radiography (CXR)\n"
    "While CXR has low sensitivity for early COPD, it is critical during acute exacerbations to rule out complications (pneumothorax, pneumonia, pulmonary edema). "
    "Characteristic radiographic signs of hyperinflation and vascular attenuation include:\n"
    "1. Diaphragmatic Flattening: Flattened hemidiaphragms with blunted costophrenic angles. On PA view, the right dome lies below the 6th anterior or 10th posterior rib. "
    "On lateral view, the diaphragmatic curvature is lost (sternodiaphragmatic angle > 90°).\n"
    "2. Increased Retrosternal Clear Space: On lateral CXR, the lucent space between the posterior cortex of the sternum and the ascending aorta exceeds 2.5 to 3.0 cm.\n"
    "3. Narrow, Elongated Vertical Heart ('Tubular Heart'): Driven by diaphragmatic descent and low intrathoracic cardiac position.\n"
    "4. Peripheral Vascular Oligemia: Rapid tapering and attenuation of pulmonary vessels from hilum to periphery ('pruning')."
)
doc.add_paragraph(
    "6.2 High-Resolution Computed Tomography (HRCT) of the Thorax\n"
    "HRCT is the imaging modality of choice for phenotyping COPD and evaluating interventional candidacy:\n"
    "1. Centrilobular Emphysema: Low-attenuation focal lucencies (<-950 HU) lacking visible walls, clustered around the centrilobular core bronchiole in upper lobes.\n"
    "2. Panlobular Emphysema: Diffuse, uniform geographic destruction of entire secondary pulmonary lobules with severe vascular attenuation in lower lobes.\n"
    "3. Paraseptal Emphysema: Subpleural single-layer cystic lucencies bounded by interlobular septa along the visceral pleura and fissures; prone to bullous formation.\n"
    "4. Quantitative CT (QCT): Uses software algorithms to calculate the 'Emphysema Index' (percentage of voxels with attenuation <-950 HU on inspiratory CT) and "
    "'Pi10' (airway wall thickness at a standardized internal perimeter of 10 mm)."
)
doc.add_paragraph(
    "6.3 Circulating Biomarkers in COPD\n"
    "1. Blood Eosinophil Count: The most clinically actionable biomarker in stable COPD. Counts >= 300 cells/µL predict significant exacerbation reduction from inhaled corticosteroids (ICS). "
    "Counts < 100 cells/µL indicate minimal ICS response and increased risk of bacterial pneumonia.\n"
    "2. Serum Procalcitonin: Used during acute exacerbations to guide antibiotic stewardship (procalcitonin < 0.1 µg/L strongly argues against bacterial infection).\n"
    "3. Fibrinogen & hs-CRP: Elevated systemic inflammatory markers associated with increased cardiovascular events and hospitalization risk."
)

doc.add_page_break()

# ---------------------------------------------------------------------------
# CHAPTER 7: STABLE COPD ASSESSMENT & STAGING
# ---------------------------------------------------------------------------
h_c7 = doc.add_heading("Chapter 7: Stable COPD Assessment: The GOLD ABE Paradigm & BODE Index", level=1)
for r in h_c7.runs:
    r.font.name = "Calibri"; r.font.size = Pt(16); r.font.bold = True; r.font.color.rgb = RGBColor.from_string(COLOR_NAVY_HEX)

doc.add_paragraph(
    "7.1 The Paradigm Shift: From ABCD to the GOLD ABE Assessment Scheme\n"
    "In 2023-2025, GOLD permanently retired the historical ABCD classification. Groups C and D were merged into GROUP E (Exacerbation group) to recognize that "
    "exacerbation frequency and severity carry supreme clinical urgency regardless of baseline symptom scores:"
)

# Insert GOLD ABE Diagram
p_img3 = doc.add_paragraph()
p_img3.alignment = WD_ALIGN_PARAGRAPH.CENTER
doc.add_picture(str(diag_path3), width=Inches(6.2))
p_cap3 = doc.add_paragraph("Figure 7.1: The GOLD 2024/2025 ABE Assessment Scheme and Initial Pharmacotherapy Pathways.")
p_cap3.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_cap3.runs[0].font.size = Pt(9.5); p_cap3.runs[0].font.italic = True; p_cap3.runs[0].font.color.rgb = RGBColor.from_string(COLOR_MUTED_HEX)

doc.add_paragraph(
    "7.2 Initial Pharmacotherapy Pathways:\n"
    "• Group A: Single Bronchodilator (SABA, LABA, or LAMA). Choice based on symptom relief and patient preference.\n"
    "• Group B: Dual Long-Acting Bronchodilation (LABA + LAMA). Superior to monotherapy in improving lung function, symptoms, and health status.\n"
    "• Group E: Initial therapy is LABA + LAMA combination. If blood eosinophils >= 300 cells/µL, initiate UPFRONT TRIPLE THERAPY (LABA + LAMA + ICS)."
)
doc.add_paragraph(
    "7.3 Multidimensional Prognostication: The BODE Index\n"
    "Because FEV1 alone correlates poorly with mortality and systemic manifestations, Celli et al. (NEJM 2004) validated the BODE Index (Body mass index, "
    "Obstruction, Dyspnea, Exercise capacity). It scores patients from 0 to 10 points and provides highly accurate 52-month mortality predictions."
)

bode_table_data = [
    ["Variable", "0 Points", "1 Point", "2 Points", "3 Points"],
    ["B — Body Mass Index (BMI, kg/m²)", "> 21", "≤ 21", "—", "—"],
    ["O — Obstruction (FEV1 % predicted)", "≥ 65%", "50 - 64%", "36 - 49%", "≤ 35%"],
    ["D — Dyspnea (mMRC scale)", "Grade 0 - 1", "Grade 2", "Grade 3", "Grade 4"],
    ["E — Exercise (6MWD in meters)", "≥ 350 m", "250 - 349 m", "150 - 249 m", "≤ 149 m"]
]
t_bode = doc.add_table(rows=len(bode_table_data), cols=5)
format_table(t_bode, [Inches(2.0), Inches(1.1), Inches(1.1), Inches(1.1), Inches(1.2)], bode_table_data[0], bode_table_data[1:])

doc.add_paragraph()
doc.add_paragraph(
    "Interpretation of BODE Quartiles:\n"
    "• Quartile 1 (Score 0-2): 52-month mortality approximately 20%.\n"
    "• Quartile 2 (Score 3-4): 52-month mortality approximately 30%.\n"
    "• Quartile 3 (Score 5-6): 52-month mortality approximately 40%.\n"
    "• Quartile 4 (Score 7-10): 52-month mortality approximately 80%."
)

doc.add_page_break()

# ---------------------------------------------------------------------------
# CHAPTER 8: COMPREHENSIVE PHARMACOTHERAPY
# ---------------------------------------------------------------------------
h_c8 = doc.add_heading("Chapter 8: Comprehensive Pharmacotherapy and Inhaler Delivery", level=1)
for r in h_c8.runs:
    r.font.name = "Calibri"; r.font.size = Pt(16); r.font.bold = True; r.font.color.rgb = RGBColor.from_string(COLOR_NAVY_HEX)

doc.add_paragraph(
    "8.1 Inhalation Science & Device Ergonomics\n"
    "The clinical efficacy of inhaled pharmacotherapy is entirely governed by aerosol deposition within conducting airways (aerodynamic diameter 1 to 5 µm). "
    "Prescribing the correct inhaler device matched to the patient's peak inspiratory flow (PIF) and manual dexterity is as vital as choosing the drug molecule:"
)

# Insert Inhaler Selection Diagram
p_img4 = doc.add_paragraph()
p_img4.alignment = WD_ALIGN_PARAGRAPH.CENTER
doc.add_picture(str(diag_path4), width=Inches(6.2))
p_cap4 = doc.add_paragraph("Figure 8.1: Peak Inspiratory Flow (PIF) based Inhaler Selection Algorithm.")
p_cap4.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_cap4.runs[0].font.size = Pt(9.5); p_cap4.runs[0].font.italic = True; p_cap4.runs[0].font.color.rgb = RGBColor.from_string(COLOR_MUTED_HEX)

doc.add_paragraph(
    "8.2 Core Bronchodilator Classes\n"
    "1. Long-Acting Muscarinic Antagonists (LAMAs - Tiotropium, Glycopyrronium, Umeclidinium): Block M3 muscarinic receptors on airway smooth muscle. "
    "In the landmark POET-COPD trial, Tiotropium was superior to Salmeterol in delaying time to first exacerbation and reducing hospitalization rates.\n"
    "2. Long-Acting Beta2-Agonists (LABAs - Formoterol, Salmeterol, Vilanterol, Indacaterol): Stimulate beta2-adrenergic receptors, increasing cyclic AMP. "
    "Ultra-LABAs (Vilanterol, Indacaterol) provide full 24-hour bronchodilation for once-daily dosing.\n"
    "3. Dual Bronchodilation (LABA + LAMA): Dual bronchodilators yield superior FEV1 improvements, greater reduction in hyperinflation, and lower exacerbation rates "
    "than monotherapies (demonstrated in FLAME and SPARK clinical trials)."
)
doc.add_paragraph(
    "8.3 Inhaled Corticosteroids (ICS) & Single-Inhaler Triple Therapy (SITT)\n"
    "The role of ICS in COPD is restricted strictly to patients with an eosinophilic or frequent exacerbator phenotype. The IMPACT (2018) and ETHOS (2020) trials "
    "definitively established that Single-Inhaler Triple Therapy (Fluticasone furoate/Umeclidinium/Vilanterol and Budesonide/Glycopyrronium/Formoterol) achieves:\n"
    "• Significant 15-25% reduction in moderate-to-severe exacerbations compared to dual bronchodilators.\n"
    "• A statistically significant ALL-CAUSE MORTALITY BENEFIT in high-risk exacerbating COPD patients.\n"
    "Pneumonia Risk: High-dose ICS significantly increases the risk of bacterial pneumonia (highest with fluticasone propionate, lower with budesonide). "
    "ICS should NEVER be prescribed as monotherapy in COPD."
)
doc.add_paragraph(
    "8.4 Oral & Targeted Therapies\n"
    "• Phosphodiesterase-4 Inhibitors (Roflumilast 500 µg OD): Reduces intracellular cAMP breakdown in inflammatory cells. Indicated in GOLD 3-4 COPD with chronic bronchitis "
    "and frequent exacerbations despite triple therapy. Reduces exacerbation rates by ~17% (REACT and RE2SPOND trials).\n"
    "• Long-Term Macrolide Prophylaxis (Azithromycin 250 mg OD or 500 mg 3x/week): Anti-inflammatory and immunomodulatory properties reduce exacerbations in ex-smokers. "
    "Mandates baseline and periodic audiometry and ECG (QTc interval) monitoring.\n"
    "• Novel Biologics (Dupilumab): In 2023-2024, the BOREAS and NOTUS phase III randomized trials demonstrated that Dupilumab (anti-IL-4Rα monoclonal antibody) "
    "significantly reduced exacerbations by 30-34% and improved FEV1 in patients with COPD, elevated blood eosinophils (>= 300 cells/µL), and chronic bronchitis on triple therapy."
)

doc.add_page_break()

# ---------------------------------------------------------------------------
# CHAPTER 9: NON-PHARMACOLOGICAL MANAGEMENT
# ---------------------------------------------------------------------------
h_c9 = doc.add_heading("Chapter 9: Non-Pharmacological Management and Rehabilitation", level=1)
for r in h_c9.runs:
    r.font.name = "Calibri"; r.font.size = Pt(16); r.font.bold = True; r.font.color.rgb = RGBColor.from_string(COLOR_NAVY_HEX)

doc.add_paragraph(
    "9.1 Smoking Cessation: The Cornerstone Intervention\n"
    "Smoking cessation is the only intervention (alongside LTOT and surgical LVRS) proven to alter the progressive downhill rate of FEV1 decline in COPD (Fletcher-Peto curve). "
    "Pharmacotherapy combined with behavioral counseling achieves the highest long-term quit rates:\n"
    "• Varenicline (Partial α4β2 nicotinic acetylcholine receptor agonist): 0.5 mg OD for 3 days, 0.5 mg BD for 4 days, then 1 mg BD for 12 weeks. Most effective monotherapy.\n"
    "• Bupropion (Atypical antidepressant, dopamine/norepinephrine reuptake inhibitor): 150 mg OD for 3 days then 150 mg BD for 7-12 weeks.\n"
    "• Nicotine Replacement Therapy (NRT): Transdermal patches combined with short-acting gum/lozenges."
)
doc.add_paragraph(
    "9.2 Comprehensive Pulmonary Rehabilitation\n"
    "Pulmonary rehabilitation is a supervised, multidisciplinary intervention lasting 6 to 12 weeks with Level A evidence for improving symptoms, exercise tolerance, "
    "and health-related quality of life, and reducing hospital readmissions following an exacerbation:\n"
    "1. Aerobic Endurance Training: Treadmill walking or cycle ergometry at 60-80% of peak work rate for 20-30 minutes per session.\n"
    "2. Resistance / Strength Training: Upper and lower extremity weight training targeting quadriceps, biceps, and deltoids.\n"
    "3. Inspiratory Muscle Training (IMT): Using threshold load devices in patients with documented inspiratory muscle weakness (MIP < 60 cmH2O)."
)
doc.add_paragraph(
    "9.3 Long-Term Oxygen Therapy (LTOT)\n"
    "Based on the landmark Nocturnal Oxygen Therapy Trial (NOTT) and Medical Research Council (MRC) trials, LTOT improves survival when prescribed for >= 15 hours/day "
    "(preferably 24 hours/day) in patients meeting rigid arterial blood gas criteria in stable state (at least 3-4 weeks post-exacerbation):\n"
    "• Absolute Indication: Resting PaO2 <= 55 mmHg (SpO2 <= 88%).\n"
    "• Conditional Indication: Resting PaO2 56 - 59 mmHg (SpO2 89%) if accompanied by secondary polycythemia (hematocrit > 55%), pulmonary hypertension (P pulmonale on ECG), "
    "or clinical cor pulmonale with dependent pedal edema.\n"
    "Target: Titrate flow rate (typically 1 to 2 L/min via nasal prongs) to maintain PaO2 60 to 65 mmHg (SpO2 90 - 92%). Avoid over-oxygenation (SpO2 > 93%) to prevent acute CO2 narcosis."
)
doc.add_paragraph(
    "9.4 Nocturnal Non-Invasive Ventilation (NIV) in Chronic Hypercapnic COPD\n"
    "Köhnlein et al. (Lancet Respir Med 2014) established that High-Intensity Nocturnal NIV targeting a marked reduction in daytime PaCO2 (by at least 20% or to < 48 mmHg) "
    "significantly improves 1-year survival (mortality 12% in NIV group vs. 33% in control group) in patients with severe stable COPD and persistent daytime hypercapnia (PaCO2 >= 52 - 53 mmHg)."
)
doc.add_paragraph(
    "9.5 Adult Immunization Recommendations in COPD (GOLD & CDC Guidelines)\n"
    "• Influenza Vaccine: Annual administration reduces serious lower respiratory illness, hospitalization, and mortality.\n"
    "• Pneumococcal Vaccine: PCV20 (single dose) OR PCV15 followed by PPSV23 after >= 1 year; protects against bacteremia and invasive pneumococcal pneumonia.\n"
    "• Tdap Booster: Protects against Bordetella pertussis in adults not vaccinated in adolescence.\n"
    "• RSV Vaccine: Single dose of recombinant RSVPreF3 or RSVpreF recommended in adults >= 60 years with chronic lung disease.\n"
    "• Herpes Zoster (Shingrix): Recombinant zoster vaccine in adults >= 50 years (two doses)."
)

doc.add_page_break()

# ---------------------------------------------------------------------------
# CHAPTER 10: INTERVENTIONAL & SURGICAL MANAGEMENT
# ---------------------------------------------------------------------------
h_c10 = doc.add_heading("Chapter 10: Interventional Pulmonology and Surgical Approaches", level=1)
for r in h_c10.runs:
    r.font.name = "Calibri"; r.font.size = Pt(16); r.font.bold = True; r.font.color.rgb = RGBColor.from_string(COLOR_NAVY_HEX)

doc.add_paragraph(
    "10.1 Bronchoscopic Lung Volume Reduction (BLVR)\n"
    "BLVR provides a minimally invasive alternative to surgical resection for reducing severe hyperinflation in advanced emphysema:\n"
    "1. Endobronchial Valves (EBVs - Zephyr, Spiration): One-way silicone duckbill valves deployed via flexible bronchoscopy into all segmental bronchi of the target lobe. "
    "The valves permit air and secretions to exit during expiration while blocking inspiratory airflow, leading to complete target lobar atelectasis.\n"
    "2. The Collateral Ventilation Prerequisite (Chartis System): EBVs strictly require ABSENCE OF COLLATERAL VENTILATION (CV-negative). If collateral channels "
    "(canals of Lambert, pores of Kohn) exist between adjacent lobes across incomplete fissures, air will continue to fill the target lobe via collateral drift, "
    "and atelectasis will fail. High-resolution CT fissure analysis (intact fissure > 95%) and balloon-catheter Chartis assessment are mandatory prior to valve deployment.\n"
    "3. Primary Complication: Pneumothorax occurs in 20 to 25% of patients within 72 hours post-procedure, resulting from rapid expansion of the non-target ipsilateral lobe. "
    "Patients must remain hospitalized on strict bed rest with chest drainage readily available."
)
doc.add_paragraph(
    "10.2 Lung Volume Reduction Surgery (LVRS)\n"
    "Based on the National Emphysema Treatment Trial (NETT - NEJM 2003), bilateral surgical resection of 20-30% of the most destroyed emphysematous tissue via VATS or median sternotomy "
    "improves survival, FEV1, and quality of life in a specific subgroup:\n"
    "• Ideal Candidate: Upper-lobe predominant emphysema with LOW baseline post-rehabilitation exercise capacity.\n"
    "• Contraindicated (High Mortality Group): Non-upper lobe (diffuse/basilar) emphysema with FEV1 < 20% predicted and either DLCO < 20% predicted or homogenous emphysema."
)
doc.add_paragraph(
    "10.3 Bullectomy & Lung Transplantation\n"
    "• Bullectomy: Surgical excision of a single giant bulla occupying > 30-50% of the hemithorax, which compresses relatively normal adjacent lung parenchyma.\n"
    "• Lung Transplantation: Single or bilateral lung transplantation indicated in severe end-stage COPD (BODE score 7-10, FEV1 < 20% predicted, DLCO < 20%, secondary pulmonary hypertension, "
    "or frequent life-threatening exacerbations)."
)

doc.add_page_break()

# ---------------------------------------------------------------------------
# CHAPTER 11: ACUTE EXACERBATIONS OF COPD (AECOPD)
# ---------------------------------------------------------------------------
h_c11 = doc.add_heading("Chapter 11: Acute Exacerbations of COPD (AECOPD) and Critical Care", level=1)
for r in h_c11.runs:
    r.font.name = "Calibri"; r.font.size = Pt(16); r.font.bold = True; r.font.color.rgb = RGBColor.from_string(COLOR_NAVY_HEX)

doc.add_paragraph(
    "11.1 Definition, Etiology & The Rome Classification\n"
    "An exacerbation of COPD is an event characterized by dyspnea and/or cough and sputum that worsen over <= 14 days, which may be accompanied by tachypnea and/or tachycardia "
    "and is often associated with increased local and systemic inflammation caused by airway infection, pollution, or other insults.\n"
    "• Viral Triggers (30-40%): Human Rhinovirus, Influenza, Respiratory Syncytial Virus (RSV).\n"
    "• Bacterial Triggers (40-50%): Haemophilus influenzae, Streptococcus pneumoniae, Moraxella catarrhalis; in severe obstruction (FEV1 < 35%), Pseudomonas aeruginosa.\n"
    "• Environmental Triggers (10-15%): Temperature inversions, high ambient PM2.5, ozone spikes."
)
doc.add_paragraph(
    "11.2 Anthonisen Criteria for Antibiotic Therapy:\n"
    "• Type 1 (Severe): All 3 cardinal symptoms present (worsening dyspnea, increased sputum volume, increased sputum purulence) -> ANTIBIOTICS MANDATED.\n"
    "• Type 2 (Moderate): 2 of 3 cardinal symptoms -> Antibiotics indicated IF one symptom is increased sputum purulence.\n"
    "• Type 3 (Mild): 1 cardinal symptom plus at least one supporting sign (fever, wheezing, leukocytosis)."
)
doc.add_paragraph(
    "11.3 Hospital Medical Management Protocol:\n"
    "1. Controlled Oxygen Therapy: Delivered via 24% or 28% Venturi mask or low-flow nasal prongs (1-2 L/min) targeting an SpO2 of 88% to 92%. Avoid hyperoxia (SpO2 > 93%) "
    "to prevent worsening hypercapnia via the Haldane effect, dead-space expansion, and blunted hypoxic ventilatory drive.\n"
    "2. Inhaled Bronchodilators: Short-acting beta2-agonists (Salbutamol 2.5 mg) combined with anticholinergics (Ipratropium 0.5 mg) via nebulization or pMDI with spacer.\n"
    "3. Systemic Corticosteroids: Oral Prednisolone 40 mg daily for precisely 5 DAYS (demonstrated non-inferior to 14 days in the REDUCE trial; shortens recovery, "
    "reduces treatment failure, and minimizes steroid adverse events).\n"
    "4. Targeted Antibiotic Therapy: 5 to 7 days duration. Uncomplicated: Co-amoxiclav, Azithromycin, or Cefuroxime. If Pseudomonas risk factors present "
    "(frequent hospitalizations, bronchiectasis, prior Pseudomonas isolation): Levofloxacin, Ciprofloxacin, or Piperacillin-Tazobactam."
)

# Insert NIV Diagram
p_img5 = doc.add_paragraph()
p_img5.alignment = WD_ALIGN_PARAGRAPH.CENTER
doc.add_picture(str(diag_path5), width=Inches(6.2))
p_cap5 = doc.add_paragraph("Figure 11.1: Non-Invasive Ventilation (NIV) Decision Algorithm in Acute Hypercapnic Respiratory Failure.")
p_cap5.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_cap5.runs[0].font.size = Pt(9.5); p_cap5.runs[0].font.italic = True; p_cap5.runs[0].font.color.rgb = RGBColor.from_string(COLOR_MUTED_HEX)

doc.add_paragraph(
    "11.4 Invasive Mechanical Ventilation Principles in Severe COPD\n"
    "When NIV fails, patients require immediate endotracheal intubation. The primary mechanical challenge is dynamic hyperinflation and intrinsic PEEP (auto-PEEP):\n"
    "• Ventilator Settings: Mode: Assist-Control Volume Control (AC-VC) or PRVC. Tidal Volume: 6 to 8 mL/kg of PREDICTED BODY WEIGHT (never actual body weight). "
    "Respiratory Rate: Keep LOW (10 to 12 breaths/min) to allow long expiratory time (Te >= 3.5 - 4.5 s; I:E ratio 1:3 to 1:4). Peak Inspiratory Flow: High (70 to 80 L/min).\n"
    "• PEEP Setting: Apply extrinsic PEEP equal to approximately 70-80% of measured intrinsic auto-PEEP to reduce the inspiratory trigger work of breathing without expanding total lung volume.\n"
    "• Permissive Hypercapnia: Tolerate elevated PaCO2 (60-70 mmHg) as long as pH remains >= 7.20 - 7.25."
)

doc.add_page_break()

# ---------------------------------------------------------------------------
# CHAPTER 12: COMORBIDITIES & OVERLAP SYNDROMES
# ---------------------------------------------------------------------------
h_c12 = doc.add_heading("Chapter 12: Comorbidities, Overlap Syndromes, and Palliative Care", level=1)
for r in h_c12.runs:
    r.font.name = "Calibri"; r.font.size = Pt(16); r.font.bold = True; r.font.color.rgb = RGBColor.from_string(COLOR_NAVY_HEX)

doc.add_paragraph(
    "12.1 Cardiovascular Comorbidities\n"
    "Cardiovascular disease accounts for up to 30-40% of all deaths in COPD. Chronic systemic inflammation (hs-CRP, IL-6), arterial stiffness, and recurrent hypoxemic episodes "
    "drive coronary artery atherosclerosis, heart failure (HFpEF and HFrEF), and cardiac arrhythmias.\n"
    "Cardioselective Beta-Blocker Safety: Cardioselective beta1-blockers (Bisoprolol, Metoprolol succinate) are completely safe and proven to reduce cardiovascular mortality in COPD patients with CAD. "
    "Withholding beta-blockers due to unfounded fear of bronchospasm is a major clinical error."
)
doc.add_paragraph(
    "12.2 Asthma-COPD Overlap (ACO)\n"
    "ACO represents patients who display clinical features of both asthma and COPD:\n"
    "• Features: Persistent airflow limitation (post-BD FEV1/FVC < 0.70) in a patient aged >= 40 years with a heavy smoking or biomass exposure history, but accompanied by "
    "documented marked bronchodilator reversibility (FEV1 increase > 12% and > 400 mL, or > 15% predicted), personal history of asthma before age 40, elevated blood eosinophils (>= 300 cells/µL), "
    "or atopic background.\n"
    "• MANDATORY MANAGEMENT RULE: Inhaled Corticosteroids (ICS) form the compulsory backbone of therapy in ACO. LABA monotherapy is strictly contraindicated."
)
doc.add_paragraph(
    "12.3 COPD-OSA Overlap Syndrome\n"
    "Co-existence of COPD and Obstructive Sleep Apnea (OSA) occurs in ~10% of COPD patients. Overlap patients experience profound, prolonged nocturnal oxygen desaturations, "
    "accelerated development of pulmonary arterial hypertension, and significantly higher mortality than patients with either disease alone. "
    "Continuous Positive Airway Pressure (CPAP) therapy markedly improves survival and reduces hospitalizations."
)
doc.add_paragraph(
    "12.4 Palliative & End-of-Life Care in Terminal COPD\n"
    "For patients with refractory, end-stage dyspnea despite maximal medical and non-pharmacological therapies:\n"
    "• Low-Dose Opioids: Oral Morphine 5 to 10 mg twice daily (or oral suspension 2.5-5 mg q4h PRN) effectively blunts the sensation of breathlessness via central respiratory modulation "
    "without precipitating fatal hypercapnia when titrated carefully.\n"
    "• Facial Fan Therapy: A handheld fan directed at the trigeminal sensory nerve distribution (cheeks and nose) provides rapid dyspnea relief via neurosensory stimulation."
)

doc.add_page_break()

# ---------------------------------------------------------------------------
# APPENDIX: POSTGRADUATE VIVA & LANDMARK TRIAL CHEAT SHEET
# ---------------------------------------------------------------------------
h_app = doc.add_heading("Appendix: Landmark Clinical Trials & High-Yield Viva Models", level=1)
for r in h_app.runs:
    r.font.name = "Calibri"; r.font.size = Pt(16); r.font.bold = True; r.font.color.rgb = RGBColor.from_string(COLOR_NAVY_HEX)

doc.add_paragraph(
    "A.1 Landmark Pulmonology Clinical Trials in COPD\n"
    "Postgraduate candidates must cite landmark evidence to justify clinical management decisions in university examinations:"
)

trials_data = [
    ["Trial Acronym & Year", "Intervention / Comparison", "Primary Clinical Finding & Practice Impact"],
    ["TORCH (NEJM 2007)", "Salmeterol/Fluticasone vs. monotherapies vs. placebo (3 yrs)", "Reduced exacerbations & improved health status; 17% reduction in mortality (p=0.052)."],
    ["UPLIFT (NEJM 2008)", "Tiotropium HandiHaler vs. control over 4 years", "Reduced exacerbations and hospitalizations; failed to alter rate of post-BD FEV1 decline."],
    ["FLAME (NEJM 2016)", "Indacaterol/Glycopyrronium (LABA/LAMA) vs. Salm/Fluticasone", "LABA/LAMA was superior in reducing exacerbations with lower pneumonia rates."],
    ["IMPACT (NEJM 2018)", "Fluticasone furoate/Umeclidinium/Vilanterol (Triple) vs. Duals", "Single-inhaler triple therapy reduced exacerbations and demonstrated ALL-CAUSE MORTALITY BENEFIT."],
    ["ETHOS (NEJM 2020)", "Budesonide/Glycopyrronium/Formoterol (Triple) vs. Duals", "Confirmed all-cause mortality reduction in symptomatic patients with frequent exacerbations."],
    ["REDUCE (JAMA 2013)", "Prednisone 40 mg daily for 5 days vs. 14 days in AECOPD", "5-day regimen was non-inferior in time to re-exacerbation, minimizing steroid toxicity."],
    ["NETT (NEJM 2003)", "Bilateral LVRS vs. Medical Therapy in severe emphysema", "Survival benefit in upper-lobe predominant emphysema with low baseline exercise capacity."],
    ["LIBERATE (Lancet 2018)", "Zephyr Endobronchial Valves (BLVR) vs. Standard Care", "Significant FEV1, 6MWD, and SGRQ improvements in patients without collateral ventilation."],
    ["BOREAS (NEJM 2023)", "Dupilumab vs. Placebo in eosinophilic COPD on triple therapy", "30% reduction in annualized exacerbations; established biologic role in Type-2 high COPD."]
]
t_trials = doc.add_table(rows=len(trials_data), cols=3)
format_table(t_trials, [Inches(1.8), Inches(2.2), Inches(2.5)], trials_data[0], trials_data[1:])

doc.add_paragraph()
doc.add_paragraph(
    "A.2 Model 10-Marker University Theory Question Answer Blueprint:\n"
    "Question: 'Discuss the etiopathogenesis, evaluation, and management of acute exacerbation of COPD in a 65-year-old male with type 2 respiratory failure.'\n"
    "Model Structure:\n"
    "1. Formal Definition & Rome Classification (1 Mark): Define AECOPD and classify severity (mild/moderate/severe).\n"
    "2. Etiopathogenesis & Precipitating Triggers (2 Marks): Viral pathogens, bacterial shift, worsening V/Q mismatch, dynamic hyperinflation, and intrinsic PEEP.\n"
    "3. Clinical Evaluation & Emergency Triage (2 Marks): Vitals, altered sensorium (CO2 narcosis), use of accessory muscles, ABG criteria (pH < 7.35, PaCO2 > 45).\n"
    "4. Medical Pharmacotherapy Protocol (2 Marks): Controlled oxygen titration (target SpO2 88-92%), nebulized bronchodilators, Prednisolone 40 mg x 5 days, Anthonisen antibiotics.\n"
    "5. Non-Invasive Ventilation Protocol (2 Marks): Mode (BiPAP S/T), initial IPAP (10-12) and EPAP (4-5), titration, failure criteria, and intubation threshold.\n"
    "6. Discharge Criteria & Recent Evidence (1 Mark): Inhaler technique review, vaccination, pulmonary rehab referral, and REDUCE trial citation."
)

# Save Document
doc.save(str(SCRATCH_DOCX))
print(f"Scratch copy saved to: {SCRATCH_DOCX}")
try:
    doc.save(str(OUTPUT_DOCX))
    print(f"Primary copy saved to: {OUTPUT_DOCX}")
except Exception as e:
    print(f"Direct save to {OUTPUT_DOCX} deferred: {e}")
print("Textbook generation completed successfully!")
